import { CitationContext, CitationsMap } from "../typings/style"
import { applyIntextCitationStyle, asStyleIdentifier, collectIntextCitationFieldsInRange, collectNoteCitationFootnotesInRange, isBibliographyEntry, isBibliographyTitle, isIntextCitation, isNoteCitation, migrateIntextCitationsToNotes, migrateNoteCitationsToIntext, rebuildNoteCitationAtRange, renderStyledField } from "../utils/field"
import { RefreshResponseData } from "../typings/http"
import { request, getDocumentId } from "../utils/http"
import { useI10n } from "../utils/i10n"
import { logWarn } from "../utils/log"
import { withProgress } from "../utils/progress"
import type { PrefStyle } from "./preference"
import { getPreference } from "./preference"
import { findPreviousChapterBreak, getUpdateRange } from "./chapter-break"
import { notifyTaskpaneCitationsRefreshed } from "./taskpane"
import { collectBibliographyFieldsInRange, deleteExistingBibliography, insertBibliography } from "./bibliography"

const REFRESH_MESSAGE_ZH = {
  progressReason: "正在刷新引注...",
}

const REFRESH_MESSAGE_EN = {
  progressReason: "Refreshing citations...",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { refresh: REFRESH_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { refresh: REFRESH_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { refresh: REFRESH_MESSAGE_EN },
})

function isSameStyle(
  previousStyle: PrefStyle | undefined,
  nextStyle: PrefStyle,
): boolean {
  if (!previousStyle) {
    return false
  }

  return previousStyle.id === nextStyle.id
    && previousStyle.title === nextStyle.title
    && previousStyle.citationType === nextStyle.citationType
}

function moveCaretBeforeField(field: Wps.Field): void {
  const docStart = wps.ActiveDocument.Content.Start
  const caret = wps.ActiveDocument.Range().Duplicate
  let position = Math.min(field.Code.Start, field.Result.Start) - 1

  while (position > docStart) {
    caret.SetRange(position, position)
    if (caret.Fields.Count === 0) {
      caret.Select()
      return
    }
    position -= 1
  }

  caret.SetRange(docStart, docStart)
  caret.Select()
}

async function refreshBibliographyInRange(
  range: Wps.Range,
  respond: RefreshResponseData,
  prefs: NonNullable<Awaited<ReturnType<typeof getPreference>>>,
) {
  if (!respond.bibliography || respond.bibliography.length === 0) {
    return
  }

  const lines = respond.bibliography.filter((line) =>
    isBibliographyTitle(line) || isBibliographyEntry(line)
  )
  if (lines.length === 0) {
    return
  }

  const bibliographyFields = collectBibliographyFieldsInRange(range)
  const firstBibliographyField = bibliographyFields[0]
  if (!firstBibliographyField) {
    return
  }

  const insertRange = firstBibliographyField.Result.Duplicate
  insertRange.Collapse(wps.Enum.wdCollapseStart)
  deleteExistingBibliography(range)
  insertBibliography(insertRange, lines, prefs)
}

export async function onRefreshEvent() {
  await withProgress(t("refresh.progressReason"), async () => {
    const pref = await getPreference()
    if (!pref) {
      logWarn("Refresh", "No style found, stopping refresh.")
      return
    }
    if (pref.refreshAll) {
      await refreshAll(pref.syncItems)
    }
    else {
      await refresh(undefined, pref.syncItems)
    }
    notifyTaskpaneCitationsRefreshed()
  })
}

export async function refreshForStyleChange(
  previousStyle: PrefStyle | undefined,
  nextStyle: PrefStyle,
): Promise<boolean> {
  if (isSameStyle(previousStyle, nextStyle)) {
    return false
  }

  return withProgress(t("refresh.progressReason"), async () => {
    if (previousStyle && previousStyle.citationType !== nextStyle.citationType) {
      const rangeToMigrate = getUpdateRange()
      if (nextStyle.citationType === "note-citation") {
        migrateIntextCitationsToNotes(rangeToMigrate)
      }
      else {
        migrateNoteCitationsToIntext(rangeToMigrate)
      }
    }

    const refreshed = await refresh(getUpdateRange())
    notifyTaskpaneCitationsRefreshed()
    return refreshed
  })
}

export async function refresh(range?: Wps.Range, syncItems?: boolean): Promise<boolean> {
  range = range ?? getUpdateRange()
  const prefs = await getPreference()
  if (!prefs) {
    logWarn("Refresh", "No style found, stopping refresh.")
    return false
  }
  if (prefs.style.citationType === "intext-citation" as keyof CitationsMap) {
    const pairs: { field: Wps.Field, context: CitationContext }[] = []
    for (const { field, data } of collectIntextCitationFieldsInRange(range)) {
      // 断言获取页码时返回的是 number 类型
      const page = field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
      pairs.push({
        field,
        context: {
          id: data.id,
          page,
          ...data.source
        },
      })
    }
    if (pairs.length === 0) {
      logWarn("Refresh", "No intext citations found; deleting existing bibliography and stopping refresh.")
      return deleteExistingBibliography(range)
    }
    const respond = await request("refresh", {
      documentId: getDocumentId(),
      style: asStyleIdentifier(prefs.style),
      contexts: pairs.map((pair) => pair.context),
      syncItems: syncItems ?? prefs.syncItems,
    })
    if (!respond) {
      logWarn("Refresh", "Could not get response from /refresh, skipping this chapter.")
      return false
    }
    for (const { field, context } of pairs) {
      const updatedData = respond.citations.find((c) => c.id === context.id)
      if (!updatedData) {
        logWarn("Refresh", `No updated data found for citation with id ${context.id}, skipping.`)
        continue
      }
      if (!isIntextCitation(updatedData)) {
        logWarn("Refresh", `Updated data for citation with id ${context.id} is not a valid intext citation, skipping.`)
        continue
      }
      field.Data = JSON.stringify(updatedData)
      renderStyledField(field, applyIntextCitationStyle, updatedData.content)
    }
    // Refresh bibliography if exists
    await refreshBibliographyInRange(range, respond, prefs)
  }
  else if (prefs.style.citationType === "note-citation" as keyof CitationsMap) {
    const pairs: { note: Wps.Footnote, field: Wps.Field, context: CitationContext }[] = []
    for (const { note: footnote, field, data } of collectNoteCitationFootnotesInRange(range)) {
      // 断言获取页码时返回的是 number 类型
      const page = field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
      pairs.push({
        note: footnote,
        field,
        context: {
          id: data.id,
          page,
          ...data.source
        },
      })
    }
    if (pairs.length === 0) {
      logWarn("Refresh", "No note citations found; deleting existing bibliography and stopping refresh.")
      return deleteExistingBibliography(range)
    }
    const respond = await request("refresh", {
      documentId: getDocumentId(),
      style: asStyleIdentifier(prefs.style),
      contexts: pairs.map((pair) => pair.context),
      syncItems: syncItems ?? prefs.syncItems,
    })
    if (!respond) {
      logWarn("Refresh", "Could not get response from /refresh, skipping this chapter.")
      return false
    }
    for (let i = pairs.length - 1; i >= 0; i -= 1) {
      const { note, field, context } = pairs[i]
      const updatedData = respond.citations.find((c) => c.id === context.id)
      if (!updatedData) {
        logWarn("Refresh", `No updated data found for citation with id ${context.id}, skipping.`)
        continue
      }
      if (!isNoteCitation(updatedData)) {
        logWarn("Refresh", `Updated data for citation with id ${context.id} is not a valid note citation, skipping.`)
        continue
      }
      const rebuilt = rebuildNoteCitationAtRange(note, field, updatedData)
      if (!rebuilt) {
        logWarn("Refresh", `Failed to rebuild note citation with id ${context.id}, skipping.`)
      }
    }
    // Refresh bibliography if exists
    await refreshBibliographyInRange(range, respond, prefs)
  }
  return true
}

export async function refreshAll(syncItems?: boolean): Promise<void> {
  const prefs = await getPreference()
  if (!prefs?.style) {
    return
  }
  const originalRange = wps.Selection.Range.Duplicate

  // 从文档末尾开始，从后往前逐个刷新章节，避免向前修改导致索引漂移
  const docEnd = wps.ActiveDocument.Content.Duplicate
  docEnd.Collapse(wps.Enum.wdCollapseEnd)
  docEnd.Select()

  let previousRangeKey: string | null = null
  while (true) {
    const toUpdate = getUpdateRange()
    const rangeKey = `${toUpdate.Start}:${toUpdate.End}`
    if (rangeKey === previousRangeKey) {
      logWarn("Refresh", "refreshAll detected repeated update range, aborting to avoid infinite loop.", {
        rangeKey,
        selectionStart: wps.Selection.Range.Start,
        selectionEnd: wps.Selection.Range.End,
      })
      break
    }
    previousRangeKey = rangeKey
    await refresh(toUpdate, syncItems ?? prefs.syncItems)
    const prevBreak = findPreviousChapterBreak()
    if (!prevBreak) {
      break
    }
    moveCaretBeforeField(prevBreak.field)
  }

  originalRange.Select()
}

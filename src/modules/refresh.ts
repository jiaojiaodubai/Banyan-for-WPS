import { CitationContext, CitationsMap, IntextCitation, NoteCitation } from "../typings/style"
import {
  applyIntextCitationStyle,
  asStyleIdentifier,
  collectIntextCitationFieldsInRange,
  collectNoteCitationFootnotesInRange,
  fieldContentEquals,
  getWholeFieldRange,
  isBibliographyEntry,
  isBibliographyTitle,
  isIntextCitation,
  isNoteCitation,
  migrateIntextCitationsToNotes,
  migrateNoteCitationsToIntext,
  readFieldDataWithText,
  rebuildNoteCitationAtRange,
  richTextEquals,
  renderStyledFieldWithData,
  restoreMainTextAfterNote,
} from "../utils/field"
import { RefreshResponseData } from "../typings/http"
import { request, getDocumentId } from "../utils/http"
import { useI10n } from "../utils/i10n"
import { logWarn } from "../utils/log"
import { withProgress } from "../utils/progress"
import type { PrefStyle } from "./preference"
import { getPreference } from "./preference"
import { findPreviousChapterBreak, getUpdateRange } from "./chapter-break"
import { notifyTaskpaneCitationsRefreshed } from "./taskpane"
import {
  collectBibliographyFieldsInRange,
  deleteExistingBibliography,
  insertBibliography,
  updateBibliographyInPlace,
} from "./bibliography"
import { withBatchUpdate } from "../utils/batch-update"

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

function indexCitationResponses(citations: RefreshResponseData["citations"]): Map<string, RefreshResponseData["citations"][number]> {
  const index = new Map<string, RefreshResponseData["citations"][number]>()
  for (const citation of citations) {
    index.set(citation.id, citation)
  }
  return index
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
) : Promise<boolean> {
  if (!respond.bibliography || respond.bibliography.length === 0) {
    return false
  }

  const lines = respond.bibliography.filter((line) =>
    isBibliographyTitle(line) || isBibliographyEntry(line)
  )
  if (lines.length === 0) {
    return false
  }

  const bibliographyFields = collectBibliographyFieldsInRange(range)
  const firstBibliographyField = bibliographyFields[0]
  if (!firstBibliographyField) {
    return false
  }

  // 书目身份与顺序未变时保留现有域对象，避免普通刷新（只有源信息变化或什么都
  // 没变）把每条题录删掉重建。
  const inPlaceResult = updateBibliographyInPlace(range, lines, prefs)
  if (inPlaceResult !== null) {
    return inPlaceResult
  }

  // 重建插入点取域起点（`Result.Start` 在域内部，同 VBA 的 ReplaceBibliography）。
  const insertRange = getWholeFieldRange(firstBibliographyField)
  insertRange.Collapse(wps.Enum.wdCollapseStart)
  deleteExistingBibliography(range)
  insertBibliography(insertRange, lines, prefs)
  return true
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

  return withBatchUpdate("Banyan Refresh", () => withProgress(t("refresh.progressReason"), async () => {
    // 记下迁移前的选区（`getUpdateRange()` 以选区为锚），迁移后恢复，见 restoreRefreshSelection。
    const originalRange = wps.Selection.Range.Duplicate
    const originalStory = wps.Selection.StoryType

    if (previousStyle && previousStyle.citationType !== nextStyle.citationType) {
      const rangeToMigrate = getUpdateRange()
      if (nextStyle.citationType === "note-citation") {
        migrateIntextCitationsToNotes(rangeToMigrate)
      }
      else {
        migrateNoteCitationsToIntext(rangeToMigrate)
      }
      restoreRefreshSelection(originalRange, originalStory)
    }

    const refreshed = await refresh(getUpdateRange())
    notifyTaskpaneCitationsRefreshed()
    return refreshed
  }))
}

/**
 * 恢复样式切换前的选区（对应 VBA 的 RestoreRefreshSelection）。
 *
 * `getUpdateRange()` 以选区为锚计算本章范围，而迁移会挪动插入点；原选区不在正文时
 * 退回正文起点。
 */
function restoreRefreshSelection(originalRange: Wps.Range, originalStory: number): void {
  if (originalStory === wps.Enum.wdMainTextStory) {
    originalRange.Select()
    return
  }

  const caret = wps.ActiveDocument.Range().Duplicate
  const contentStart = wps.ActiveDocument.Content.Start
  caret.SetRange(contentStart, contentStart)
  caret.Select()
}

export async function refresh(range?: Wps.Range, syncItems?: boolean): Promise<boolean> {
  return withBatchUpdate("Banyan Refresh", () => refreshInRange(range, syncItems))
}

async function refreshInRange(range?: Wps.Range, syncItems?: boolean): Promise<boolean> {
  range = range ?? getUpdateRange()
  const prefs = await getPreference()
  if (!prefs) {
    logWarn("Refresh", "No style found, stopping refresh.")
    return false
  }
  if (prefs.style.citationType === "intext-citation" as keyof CitationsMap) {
    const collected = collectIntextCitationFieldsInRange(range)
    if (collected.length === 0) {
      logWarn("Refresh", "No intext citations found; deleting existing bibliography and stopping refresh.")
      return deleteExistingBibliography(range)
    }

    // 请求阶段就把每个引注的数据读出来，并把已存文本缓存到 pair 上：更新阶段
    // 据此判定「无变化」而不再触碰 Field.Data（懒解析）。id 一律取收集阶段从
    // 域代码解析出的结果（重复项已在收集时重键）。
    const pairs: IntextRefreshPair[] = []
    const contexts: CitationContext[] = []
    for (const pair of collected) {
      const { text, data } = readFieldDataWithText<IntextCitation>(pair.field)
      if (!data || !isIntextCitation(data)) {
        logWarn("Refresh", `Could not read data of intext citation ${pair.id}, skipping.`)
        continue
      }
      // 断言获取页码时返回的是 number 类型
      const page = pair.field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
      const context: CitationContext = {
        id: pair.id,
        page,
        ...data.source
      }
      contexts.push(context)
      pairs.push({ field: pair.field, data, text, context })
    }
    if (pairs.length === 0) {
      logWarn("Refresh", "No readable intext citation data; skipping this chapter.")
      return false
    }

    const respond = await request("refresh", {
      documentId: getDocumentId(),
      style: asStyleIdentifier(prefs.style),
      contexts,
      syncItems: syncItems ?? prefs.syncItems,
    })
    if (!respond) {
      logWarn("Refresh", "Could not get response from /refresh, skipping this chapter.")
      return false
    }
    const responseIndex = indexCitationResponses(respond.citations)
    let didUpdateCitation = false
    for (const pair of pairs) {
      const updatedData = responseIndex.get(pair.context.id)
      if (!updatedData) {
        logWarn("Refresh", `No updated data found for citation with id ${pair.context.id}, skipping.`)
        continue
      }
      if (!isIntextCitation(updatedData)) {
        logWarn("Refresh", `Updated data for citation with id ${pair.context.id} is not a valid intext citation, skipping.`)
        continue
      }
      const nextJson = JSON.stringify(updatedData)
      if (nextJson.length === 0) {
        logWarn("Refresh", `Could not serialize updated data for citation with id ${pair.context.id}, skipping.`)
        continue
      }
      // 与请求前缓存的文本相同：整个域都不需要触碰。
      if (pair.text === nextJson) continue
      if (applyIntextCitationData(pair.field, updatedData, pair.data)) {
        didUpdateCitation = true
      }
    }
    // 顺带刷新已有的书目
    const bibliographyChanged = await refreshBibliographyInRange(range, respond, prefs)
    return bibliographyChanged || didUpdateCitation
  }
  else if (prefs.style.citationType === "note-citation" as keyof CitationsMap) {
    const collected = collectNoteCitationFootnotesInRange(range)
    if (collected.length === 0) {
      logWarn("Refresh", "No note citations found; deleting existing bibliography and stopping refresh.")
      return deleteExistingBibliography(range)
    }

    // 与正文引注同理：请求阶段缓存已存文本与对象，更新阶段据此跳过无变化的引注。
    const pairs: NoteRefreshPair[] = []
    const contexts: CitationContext[] = []
    for (const pair of collected) {
      const { text, data } = readFieldDataWithText<NoteCitation>(pair.field)
      if (!data || !isNoteCitation(data)) {
        logWarn("Refresh", `Could not read data of note citation ${pair.id}, skipping.`)
        continue
      }
      // 断言获取页码时返回的是 number 类型
      const page = pair.field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
      const context: CitationContext = {
        id: pair.id,
        page,
        ...data.source
      }
      contexts.push(context)
      pairs.push({ note: pair.note, field: pair.field, data, text, context })
    }
    if (pairs.length === 0) {
      logWarn("Refresh", "No readable note citation data; skipping this chapter.")
      return false
    }

    const respond = await request("refresh", {
      documentId: getDocumentId(),
      style: asStyleIdentifier(prefs.style),
      contexts,
      syncItems: syncItems ?? prefs.syncItems,
    })
    if (!respond) {
      logWarn("Refresh", "Could not get response from /refresh, skipping this chapter.")
      return false
    }
    const responseIndex = indexCitationResponses(respond.citations)
    let didUpdateCitation = false
    // 重建脚注会移动范围：倒序遍历。
    for (let i = pairs.length - 1; i >= 0; i -= 1) {
      const pair = pairs[i]
      const updatedData = responseIndex.get(pair.context.id)
      if (!updatedData) {
        logWarn("Refresh", `No updated data found for citation with id ${pair.context.id}, skipping.`)
        continue
      }
      if (!isNoteCitation(updatedData)) {
        logWarn("Refresh", `Updated data for citation with id ${pair.context.id} is not a valid note citation, skipping.`)
        continue
      }
      const nextJson = JSON.stringify(updatedData)
      if (nextJson.length === 0) {
        logWarn("Refresh", `Could not serialize updated data for citation with id ${pair.context.id}, skipping.`)
        continue
      }
      if (pair.text === nextJson) continue
      if (applyNoteCitationData(pair.note, pair.field, updatedData, pair.data)) {
        didUpdateCitation = true
      }
    }
    // 顺带刷新已有的书目
    const bibliographyChanged = await refreshBibliographyInRange(range, respond, prefs)
    return didUpdateCitation || bibliographyChanged
  }
  return true
}

type CitationFieldPairData<T extends IntextCitation | NoteCitation> = {
  field: Wps.Field
  data: T
}

type IntextRefreshPair = CitationFieldPairData<IntextCitation> & {
  text: string
  context: CitationContext
}

type NoteRefreshPair = CitationFieldPairData<NoteCitation> & {
  note: Wps.Footnote
  text: string
  context: CitationContext
}

/**
 * 把响应数据写到正文引注。`Field.Data` 是一个整体序列化字符串，没有可原地修补
 * 的局部，因此除文本完全相同（已在调用处跳过）外一律原样落盘；content 比较
 * 只作为渲染门禁。原样落盘还能让下次刷新的文本短路继续命中。
 */
function applyIntextCitationData(
  field: Wps.Field,
  updatedData: IntextCitation,
  currentData: IntextCitation | null,
): boolean {
  const contentChanged = currentData === null
    ? true
    : !fieldContentEquals(currentData, updatedData)
  try {
    field.Data = JSON.stringify(updatedData)
  }
  catch (error) {
    logWarn("Refresh", `Failed to write updated data for citation with id ${updatedData.id}, skipping render.`, error)
    return false
  }
  if (contentChanged) {
    renderStyledFieldWithData(field, applyIntextCitationStyle, updatedData, updatedData.content, "character")
  }
  return true
}

/**
 * 把响应数据写到脚注引注：先比 reference（更短）、再比内容，只有呈现变化时
 * 才重建域与脚注；否则仅落数据。
 */
function applyNoteCitationData(
  note: Wps.Footnote,
  field: Wps.Field,
  updatedData: NoteCitation,
  currentData: NoteCitation | null,
): boolean {
  let presentationChanged: boolean
  if (currentData === null) {
    presentationChanged = true
  }
  else {
    presentationChanged = !richTextEquals(currentData.reference, updatedData.reference)
    if (!presentationChanged) {
      presentationChanged = !fieldContentEquals(currentData, updatedData)
    }
  }

  if (presentationChanged) {
    const rebuilt = rebuildNoteCitationAtRange(note, field, updatedData)
    if (!rebuilt) {
      logWarn("Refresh", `Failed to rebuild note citation with id ${updatedData.id}, skipping.`)
      return false
    }
    // 重建脚注会把插入点留在脚注里，立即回移到引用之后。
    restoreMainTextAfterNote(rebuilt.note)
    return true
  }

  try {
    field.Data = JSON.stringify(updatedData)
  }
  catch (error) {
    logWarn("Refresh", `Failed to write updated data for note citation with id ${updatedData.id}.`, error)
    return false
  }
  return true
}

export async function refreshAll(syncItems?: boolean): Promise<void> {
  await withBatchUpdate("Banyan Refresh", () => refreshAllInDocument(syncItems))
}

async function refreshAllInDocument(syncItems?: boolean): Promise<void> {
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

import { request, getDocumentId } from "../utils/http"
import { useI10n } from "../utils/i10n"
import { withProgress } from "../utils/progress"
import {
  FIELD_PLACEHOLDER_COLOR,
  addBookmarkToField,
  applyStyleToField,
  asStyleIdentifier,
  getBibliographyBookmarkName,
  isBibliographyEntry,
  isBibliographyTitle,
  isIntextCitation,
  isNoteCitation,
  readFieldData,
  removeFieldSafely,
  renderStyledField,
} from "../utils/field"
import type { CitationContext, BibliographyLine } from "../typings/style"
import { getPreference, savePreference } from "./preference"
import { getUpdateRange } from "./chapter-break"
import { logError } from "../utils/log"

const BIBLIOGRAPHY_MESSAGE_ZH = {
  notInMainText: "参考文献表只能插入在主文档正文中。",
  multipleFields: "检测到多个域，请只选择一个书目条目域或将光标放在要插入书目的位置。",
  notBanyanBibliographyEntry: "请选择 Banyan 创建的书目条目域。",
  noStyle: "请先在设置中选择引用样式。",
  noCitationInRange: "当前章节没有检测到引注，请先添加引注。",
  refreshFailed: "获取参考文献表数据失败。",
  bibliographyFailed: "编辑参考文献条目失败。",
  invalidBibliographyLine: "服务器返回的书目数据无效。",
  error: "操作参考文献表时发生错误：{message}",
  progressReason: "正在处理参考文献表...",
}

const BIBLIOGRAPHY_MESSAGE_EN = {
  notInMainText: "Bibliography can only be inserted in the main text story.",
  multipleFields:
    "Multiple fields detected. Select only one bibliography entry field or place the cursor where you want to insert bibliography.",
  notBanyanBibliographyEntry: "Please select a Banyan created bibliography entry field.",
  noStyle: "Please set citation style in Preferences first.",
  noCitationInRange: "No citation was found in the current section. Please add citations first.",
  refreshFailed: "Failed to fetch bibliography data.",
  bibliographyFailed: "Failed to edit bibliography entry.",
  invalidBibliographyLine: "Invalid bibliography data returned by server.",
  error: "An error occurred while handling bibliography: {{message}}",
  progressReason: "Processing bibliography...",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: {
    bibliography: BIBLIOGRAPHY_MESSAGE_ZH,
  },
  [wps.Enum.msoLanguageIDChineseSingapore]: {
    bibliography: BIBLIOGRAPHY_MESSAGE_ZH,
  },
  [wps.Enum.msoLanguageIDEnglishUS]: {
    bibliography: BIBLIOGRAPHY_MESSAGE_EN,
  },
})

type ActionMode =
  | { type: "add"; range: Wps.Range }
  | { type: "edit"; field: Wps.Field }

const PENDING_BIBLIOGRAPHY_LINE: BibliographyLine = {
  type: "bibliography-title",
  content: {
    text: "{ BIBLIOGRAPHY }",
    marks: [
      {
        type: "color",
        start: 0,
        end: "{ BIBLIOGRAPHY }".length,
        value: FIELD_PLACEHOLDER_COLOR,
      },
    ],
  },
}

export async function onBibliographyEvent(): Promise<void> {
  try {
    if (wps.Selection.StoryType !== wps.Enum.wdMainTextStory) {
      alert(t("bibliography.notInMainText"))
      return
    }
    const pref = await getPreference()
    if (!pref) {
      alert(t("bibliography.noStyle"))
      return
    }
    const style = pref.style
    const mode = getActionMode()
    if (!mode) {
      return
    }
    const updateRange = getUpdateRange()
    const contexts = collectCitationContexts(updateRange)
    if (contexts.length === 0) {
      alert(t("bibliography.noCitationInRange"))
      return
    }
    if (mode.type === "add") {
      const insertRange = mode.range.Duplicate
      let pendingField: Wps.Field | null = createPendingBibliographyField(insertRange, pref)
      try {
        await withProgress(t("bibliography.progressReason"), async () => {
          const response = await request("refresh", {
            documentId: getDocumentId(),
            style: asStyleIdentifier(style),
            contexts,
            syncItems: pref.syncItems,
          })
          if (!response) {
            removeFieldSafely(pendingField)
            pendingField = null
            alert(t("bibliography.refreshFailed"))
            return
          }
          const lines = response.bibliography.filter((line) =>
            isBibliographyTitle(line) || isBibliographyEntry(line),
          )
          if (lines.length === 0) {
            removeFieldSafely(pendingField)
            pendingField = null
            alert(t("bibliography.invalidBibliographyLine"))
            return
          }
          deleteExistingBibliography(updateRange)
          insertBibliography(insertRange, lines, pref)
          pendingField = null
        })
      }
      catch (error) {
        removeFieldSafely(pendingField)
        throw error
      }
      return
    }
    const currentLine = readFieldData(mode.field)
    if (!isBibliographyEntry(currentLine)) {
      alert(t("bibliography.notBanyanBibliographyEntry"))
      return
    }
    const response = await request("bibliography", {
      documentId: getDocumentId(),
      style: asStyleIdentifier(style),
      line: currentLine,
      extraSource: pref.extraSource,
    })
    if (!response) {
      alert(t("bibliography.bibliographyFailed"))
      return
    }
    if (!isBibliographyEntry(response.line)) {
      alert(t("bibliography.invalidBibliographyLine"))
      return
    }
    mode.field.Data = JSON.stringify(response.line)
    renderStyledField(
      mode.field,
      (field) => applyStyleToField(field, pref.bibliographyEntryStyle, "paragraph"),
      response.line.content,
    )
    await savePreference({
      ...pref,
      extraSource: response.extraSource,
    })
  }
  catch (error) {
    logError("Bibliography", "Failed to handle bibliography event.", error)
    alert(t("bibliography.error", { message: getErrorMessage(error) }))
  }
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message
  }
  return String(error)
}

function getActionMode(): ActionMode | null {
  const range = wps.Selection.Range.Duplicate
  if (range.Fields.Count === 0) {
    range.Collapse(wps.Enum.wdCollapseEnd)
    return {
      type: "add",
      range,
    }
  }
  if (range.Fields.Count > 1) {
    alert(t("bibliography.multipleFields"))
    return null
  }
  range.Collapse(wps.Enum.wdCollapseStart)
  const field = range.Fields.Item(1)
  const data = readFieldData(field)
  if (!isBibliographyEntry(data)) {
    alert(t("bibliography.notBanyanBibliographyEntry"))
    return null
  }
  return {
    type: "edit",
    field,
  }
}

function collectCitationContexts(range: Wps.Range): CitationContext[] {
  const contexts = new Map<string, CitationContext>()
  for (let i = 1; i <= range.Fields.Count; i++) {
    const field = range.Fields.Item(i)
    if (field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }
    const data = readFieldData(field)
    if (!isIntextCitation(data)) {
      continue
    }
    const page = field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
    contexts.set(data.id, {
      id: data.id,
      page,
      ...data.source,
    })
  }
  for (let i = 1; i <= range.Footnotes.Count; i++) {
    const note = range.Footnotes.Item(i)
    if (note.Range.Fields.Count === 0) {
      continue
    }
    const field = note.Range.Fields.Item(1)
    const data = readFieldData(field)
    if (!isNoteCitation(data)) {
      continue
    }
    const page = field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
    contexts.set(data.id, {
      id: data.id,
      page,
      ...data.source,
    })
  }
  return [...contexts.values()]
}

export function collectBibliographyFieldsInRange(range: Wps.Range): Wps.Field[] {
  const bibliographyFields: Wps.Field[] = []
  for (let i = 1; i <= range.Fields.Count; i++) {
    const field = range.Fields.Item(i)
    if (field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }
    const data = readFieldData(field)
    if (!isBibliographyTitle(data) && !isBibliographyEntry(data)) {
      continue
    }
    bibliographyFields.push(field)
  }
  return bibliographyFields
}

export function deleteExistingBibliography(range: Wps.Range): boolean {
  const bibliographyFields = collectBibliographyFieldsInRange(range)

  const firstField = bibliographyFields[0]
  if (!firstField) return false

  firstField.Delete()
  for (let i = bibliographyFields.length - 1; i >= 1; i--) {
    // 在 Word API中，仅仅删除 Field 并不会删除整段，但 WPS 这里会
    bibliographyFields[i].Result.Delete()
  }
  return true
}

export function insertBibliography(range: Wps.Range, lines: BibliographyLine[], pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string }) {
  const caret = range.Duplicate
  caret.Collapse(wps.Enum.wdCollapseEnd)
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    const fieldCode = "id" in line
      ? `BANYAN_BIBLIOGRAPHY ${line.id}`
      : "BANYAN_BIBLIOGRAPHY"
    const field = caret.Fields.Add(
      caret,
      wps.Enum.wdFieldAddin,
      fieldCode,
      false,
    )
    if (isBibliographyTitle(line)) {
      field.Data = JSON.stringify(line)
      renderStyledField(
        field,
        (field) => applyStyleToField(field, pref.bibliographyTitleStyle, "paragraph"),
        line.content,
      )
    }
    else if (isBibliographyEntry(line)) {
      field.Data = JSON.stringify(line)
      renderStyledField(
        field,
        (field) => applyStyleToField(field, pref.bibliographyEntryStyle, "paragraph"),
        line.content,
      )
      addBookmarkToField(field, getBibliographyBookmarkName(line.id))
    }

    if (i < lines.length - 1) {
      field.Result.InsertParagraphAfter()
      // 在域后面插入段落标记后，光标会停留在段落标记前面，这样在域后面就有一个额外的偏移量
      caret.SetRange(field.Result.End + 1, field.Result.End + 1)
    }
  }
}

function createPendingBibliographyField(
  range: Wps.Range,
  pref: { bibliographyTitleStyle: string }
): Wps.Field {
  const cursor = range.Duplicate
  cursor.Collapse(wps.Enum.wdCollapseEnd)
  const field = cursor.Fields.Add(
    cursor,
    wps.Enum.wdFieldAddin,
    "BANYAN_BIBLIOGRAPHY",
    false,
  )
  field.Data = JSON.stringify(PENDING_BIBLIOGRAPHY_LINE)
  renderStyledField(
    field,
    (field) => applyStyleToField(field, pref.bibliographyTitleStyle, "paragraph"),
    PENDING_BIBLIOGRAPHY_LINE.content,
  )
  return field
}

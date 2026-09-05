import { request, getDocumentId } from "../utils/http"
import { useI10n } from "../utils/i10n"
import { withProgress } from "../utils/progress"
import { getPreference } from "./preference"
import {
  createEmptyCitationSource,
  createIntextCitationAtRange,
  createNoteCitationAtRange,
  createPlaceholderIntextCitationData,
  createPlaceholderNoteCitationData,
  applyIntextCitationStyle,
  asStyleIdentifier,
  isIntextCitation,
  isNoteCitation,
  readFieldData,
  removeFieldSafely,
  removeFootnoteSafely,
  renderStyledField,
  FieldAndData,
} from "../utils/field"
import { IntextCitation, NoteCitation } from "../typings/style"
import { refresh, refreshAll } from "./refresh"
import { notifyTaskpaneCitationsRefreshed } from "./taskpane"
import { StyleIdentifier } from "../typings/http"

const CITATION_MESSAGE_ZH = {
  notInMainText: "引注只能插入在主文档正文中。",
  notBanyanIntextCitation: "请选择Banyan创建的引注。",
  notBanyanNoteCitation:
    "请选择Banyan创建的脚注引用。",
  noStyle: "请先在设置中选择引用样式。",
  error: "操作引注时发生错误：{message}",
  progressReason: "正在处理引注...",
}
const CITATION_MESSAGE_EN = {
  notInMainText:
    "Citations can only be inserted into the main text of the document.",
  notBanyanIntextCitation: "Please select a Banyan created citation.",
  notBanyanNoteCitation: "Please select a Banyan created note citation.",
  noStyle: "Please set citation style in Preferences first.",
  error: "An error occurred while handling the citation: {message}",
  progressReason: "Processing citation...",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { citation: CITATION_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { citation: CITATION_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { citation: CITATION_MESSAGE_EN },
})

export function insertIntextCitation(range: Wps.Range, data: IntextCitation): Wps.Field {
  return createIntextCitationAtRange(range, data)
}

export function insertNoteCitation(range: Wps.Range, data: NoteCitation): {
  note: Wps.Footnote,
  field: Wps.Field,
} {
  return createNoteCitationAtRange(range, data)
}

export async function onCitationEvent() {
  if (wps.Selection.StoryType !== wps.Enum.wdMainTextStory) {
    alert(t("citation.notInMainText"))
    return
  }
  const pref = await getPreference()
  if (!pref) {
    alert(t("citation.noStyle"))
    return
  }
  const range = wps.Selection.Range.Duplicate
  const caret = range.End
  const paragraph = range.Paragraphs.Last.Range.Duplicate

  let updated = false
  switch (pref.style.citationType) {
    case "intext-citation": {
      let found = false
      for (let i = 1; i <= paragraph.Fields.Count; i++) {
        const field = paragraph.Fields.Item(i)
        if (field.Type !== wps.Enum.wdFieldAddin) continue
        const inField = field.Result.Start <= caret && field.Result.End >= caret
        if (!inField) continue
        const data = readFieldData(field)
        if (isIntextCitation(data)) {
          found = true
          updated = await editIntextCitation({ field, data }, asStyleIdentifier(pref.style))
          break
        }
        else {
          alert(t("citation.notBanyanIntextCitation"))
          return
        }
      }
      if (!found) {
        updated = await addIntextCitation(range, asStyleIdentifier(pref.style))
      }
      break
    }
    case "note-citation": {
      let found = false
      for (let i = 1; i <= paragraph.Footnotes.Count; i++) {
        const footnote = paragraph.Footnotes.Item(i)
        const reference = footnote.Reference
        const inReference = reference.Start <= caret && reference.End >= caret
        if (!inReference) continue
        const note = footnote.Range
        for (let j = 1; j <= note.Fields.Count; j++) {
          const field = note.Fields.Item(j)
          if (field.Type !== wps.Enum.wdFieldAddin) continue
          const data = readFieldData(field)
          if (isNoteCitation(data)) {
            found = true
            updated = await editNoteCitation({ field, data }, asStyleIdentifier(pref.style))
            break
          }
          else {
            alert(t("citation.notBanyanNoteCitation"))
            return
          }
        }
      }
      if (!found) {
        updated = await addNoteCitation(range, asStyleIdentifier(pref.style))
      }
      break
    }
  }

  if (!updated) {
    return
  }

  await withProgress(t("citation.progressReason"), async () => {
    if (pref.refreshAll) {
      await refreshAll()
    }
    else {
      await refresh()
    }
    notifyTaskpaneCitationsRefreshed()
  })
}

async function addIntextCitation(
  range: Wps.Range,
  style: StyleIdentifier
): Promise<boolean> {
  const id = crypto.randomUUID()
  const data = createPlaceholderIntextCitationData(id, createEmptyCitationSource())
  const field = insertIntextCitation(range, data)

  try {
    const source = await request("citation", {
      documentId: getDocumentId(),
      style: style,
    })
    if (!source) {
      removeFieldSafely(field)
      return false
    }
    data.source = source
    field.Data = JSON.stringify(data)
    return true
  }
  catch (error) {
    removeFieldSafely(field)
    throw error
  }
}

async function editIntextCitation(
  fd: FieldAndData<IntextCitation>,
  style: StyleIdentifier
): Promise<boolean> {
  const { field, data } = fd
  const newSource = await request("citation", {
    documentId: getDocumentId(),
    style: style,
    source: data.source,
  })
  if (!newSource) return false
  data.source = newSource
  field.Data = JSON.stringify(data)
  renderStyledField(field, applyIntextCitationStyle)
  return true
}

async function addNoteCitation(
  range: Wps.Range,
  style: StyleIdentifier
): Promise<boolean> {
  const id = crypto.randomUUID()
  const data = createPlaceholderNoteCitationData(id, createEmptyCitationSource())
  const { note, field } = insertNoteCitation(range, data)

  try {
    const source = await request("citation", {
      documentId: getDocumentId(),
      style: style,
    })
    if (!source) {
      removeFootnoteSafely(note)
      return false
    }
    data.source = source
    field.Data = JSON.stringify(data)
    return true
  }
  catch (error) {
    removeFootnoteSafely(note)
    throw error
  }
}

async function editNoteCitation(
  fd: FieldAndData<NoteCitation>,
  style: StyleIdentifier
): Promise<boolean> {
  const { field, data } = fd
  const newSource = await request("citation", {
    documentId: getDocumentId(),
    style: style,
    source: data.source,
  })
  if (!newSource) return false
  data.source = newSource
  field.Data = JSON.stringify(data)
  return true
}

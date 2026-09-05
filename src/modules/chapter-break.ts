import type { CitationSource } from "../typings/style"
import type { BanyanFieldData, FieldAndData } from "../utils/field"
import { createEmptyRichText, isBanyanFieldData, isUnknownRecord, readFieldData } from "../utils/field"
import { useI10n } from "../utils/i10n"
import { getPreference } from "./preference"
import type { PrefStyle } from "./preference"

const CHAPTER_BREAK_MESSAGE_ZH = {
  chapterBreak: " ==========Banyan章节分隔符（请勿编辑）========== ",
  notInMainText: "章节分隔符只能插入在主文档正文中。",
  notAtParagraphStart: "章节分隔符只能插入在段落开头。",
  inField: "章节分隔符不能插入在域代码中，请将光标移到普通正文位置。",
  noStyle: "请先在设置中选择引用样式。",
}

const CHAPTER_BREAK_MESSAGE_EN = {
  chapterBreak: " ==========Banyan chapter break (Do not edit)========== ",
  notInMainText: "Chapter breaks can only be inserted into the main text of the document.",
  notAtParagraphStart: "Chapter breaks can only be inserted at the beginning of a paragraph.",
  inField: "Chapter breaks cannot be inserted inside fields. Move the cursor to normal text.",
  noStyle: "Please set citation style in Preferences first.",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { chapterBreak: CHAPTER_BREAK_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { chapterBreak: CHAPTER_BREAK_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { chapterBreak: CHAPTER_BREAK_MESSAGE_EN },
})

export type ChapterBreak = BanyanFieldData & {
  type: "chapter-break";
  style: PrefStyle;
  bibliographyTitleStyle: string;
  bibliographyEntryStyle: string;
  extraSource?: CitationSource;
}

export function isChapterBreak(v: unknown): v is ChapterBreak {
  return (
    isBanyanFieldData(v) &&
    v.type === "chapter-break" &&
    "style" in v &&
    isPrefStyle(v.style) &&
    (!("extraSource" in v) || isCitationSource(v.extraSource))
  )
}

export function isCitationSource(v: unknown): v is CitationSource {
  return (
    isUnknownRecord(v) &&
    Array.isArray(v.cites) &&
    isUnknownRecord(v.params)
  )
}

function isPrefStyle(v: unknown): v is PrefStyle {
  const props = [
    "id",
    "title",
    "citationType",
  ]
  return (
    isUnknownRecord(v) &&
    props.every((prop) => typeof v[prop] === "string")
  )
}

export async function onInsertChapterBreakEvent(): Promise<void> {
  if (wps.Selection.StoryType !== wps.Enum.wdMainTextStory) {
    alert(t("chapterBreak.notInMainText"))
    return
  }
  if (wps.Selection.Range.Fields.Count > 0) {
    alert(t("chapterBreak.inField"))
    return
  }
  if (!isSelectionAtParagraphStart(wps.Selection.Range)) {
    alert(t("chapterBreak.notAtParagraphStart"))
    return
  }
  const pref = await getPreference()
  if (!pref) {
    alert(t("chapterBreak.noStyle"))
    return
  }
  insertChapterBreak(pref.style, pref.bibliographyTitleStyle, pref.bibliographyEntryStyle, pref.extraSource)
}

export function canInsertChapterBreakAtSelection(): boolean {
  if (wps.Selection.StoryType !== wps.Enum.wdMainTextStory) {
    return false
  }
  if (wps.Selection.Range.Fields.Count > 0) {
    return false
  }
  return isSelectionAtParagraphStart(wps.Selection.Range)
}

export function insertChapterBreak(style: PrefStyle, bibliographyTitleStyle: string, bibliographyEntryStyle: string, extraSource?: CitationSource) {
  const selection = ensureCaretInMainText()
  const breakData: ChapterBreak = {
    type: "chapter-break",
    style,
    bibliographyTitleStyle,
    bibliographyEntryStyle,
    extraSource,
    // 考虑不用 Field.Result 来显示提示文本，而是直接放在域代码里
    // 这样提示文本可以在“切换域代码”时显示或隐藏，既能提醒用户又不至于引入字符影响排版细节
    content: createEmptyRichText(),
  }
  const field = selection.Fields.Add(
    selection.Range,
    wps.Enum.wdFieldAddin,
    // 从 API 传入参数，以便获得带有 `ADDIN` 前缀的域代码，这样才能读取 Data 属性
    t("chapterBreak.chapterBreak"),
  )
  field.Data = JSON.stringify(breakData)
  // 默认显示 Code 中的提示文本
  field.ShowCodes = true
  field.Locked = true
  // 突出显示，方便用户识别和删除
  field.Code.Font.Color = wps.Enum.wdColorRed
  field.Result.InsertParagraphAfter()
  // 自动换行：将光标移到分隔符后的新段落，跳过刚插入的段落标记
  const caret = field.Result.Duplicate
  caret.SetRange(field.Result.End + 1, field.Result.End + 1)
  caret.Select()
}

function isSelectionAtParagraphStart(range: Wps.Range): boolean {
  const caret = range.Duplicate
  caret.Collapse(wps.Enum.wdCollapseStart)
  if (caret.Paragraphs.Count === 0) {
    return false
  }
  const paragraphStart = caret.Paragraphs.Item(1).Range.Start
  return caret.Start === paragraphStart
}

export function findPreviousChapterBreak(): FieldAndData<ChapterBreak> | null {
  const originalRange = wps.Selection.Range.Duplicate
  try {
    const selection = ensureCaretInMainText()
    let field = selection.PreviousField()
    while (field !== null) {
      const data = readFieldData(field)
      if (isChapterBreak(data)) {
        return {
          field,
          data,
        }
      }
      field = field.Previous
    }
    return null
  }
  finally {
    originalRange.Select()
  }
}

export function findNextChapterBreak(): FieldAndData<ChapterBreak> | null {
  const originalRange = wps.Selection.Range.Duplicate
  try {
    const selection = ensureCaretInMainText()
    let field = selection.NextField()
    while (field !== null) {
      const data = readFieldData(field)
      if (isChapterBreak(data)) {
        return {
          field,
          data,
        }
      }
      field = field.Next
    }
    return null
  }
  finally {
    originalRange.Select()
  }
}

export function ensureCaretInMainText() {
  // 确保在主文档正文中：若不在，跳回当前页末尾
  if (wps.Selection.StoryType !== wps.Enum.wdMainTextStory) {
    const currentPage = wps.Selection.Information(
      wps.Enum.wdActiveEndPageNumber,
    )
    const pageRange = wps.Selection.GoTo(wps.Enum.wdGoToPage, wps.Enum.wdGoToAbsolute, currentPage)
    pageRange.Collapse(wps.Enum.wdCollapseEnd)
    pageRange.Select()
  }
  if (wps.Selection.Fields.Count > 0) {
    const field = wps.Selection.Fields.Item(1)
    if (isChapterBreak(readFieldData(field))) {
      const after = field.Result.Duplicate
      after.Collapse(wps.Enum.wdCollapseEnd)
      after.Select()
      return wps.Selection
    }
  }
  const selection = wps.Selection
  selection.Collapse(wps.Enum.wdCollapseEnd)
  return selection
}

export function getUpdateRange(): Wps.Range {
  const prevChapterBreak = findPreviousChapterBreak()
  // 范围起点：上一个分章符的开头（定义本章），或文档开头（第一章）
  const start = prevChapterBreak
    ? prevChapterBreak.field.Result.Start
    : wps.ActiveDocument.Content.Start
  const nextChapterBreak = findNextChapterBreak()
  // 范围终点：下一个分章符的开头，或文档结尾
  const end = nextChapterBreak
    ? nextChapterBreak.field.Result.Start
    : wps.ActiveDocument.Content.End
  const range = wps.ActiveDocument.Range().Duplicate
  range.SetRange(start, end)
  return range
}

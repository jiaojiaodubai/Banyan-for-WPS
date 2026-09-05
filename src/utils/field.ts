
import type {
  BibliographyEntry,
  BibliographyTitle,
  Citation,
  CitationSource,
  IntextCitation,
  NoteCitation,
} from "../typings/style"
import type { StyleIdentifier } from "../typings/http"
import type { InlineMark, RichText } from "../typings/unit"
import { logError, logWarn } from "./log"

export const FIELD_PLACEHOLDER_COLOR = "#ff0000"

const PLACEHOLDER_INTEXT_CITATION_CONTENT = createPlaceholderRichText("{ INTEXT_CITATION }")
const PLACEHOLDER_NOTE_CITATION_CONTENT = createPlaceholderRichText("{ NOTE_CITATION }")

const STYLE_NAMES = {
  intextCitation: {
    [wps.Enum.msoLanguageIDSimplifiedChinese]: "Banyan 引注",
    [wps.Enum.msoLanguageIDChineseSingapore]: "Banyan 引注",
    [wps.Enum.msoLanguageIDEnglishUS]: "Banyan Citation",
  }
}

function getIntextCitationStyleName(): string {
  const langId = Application.Language
  return STYLE_NAMES.intextCitation[langId] ?? STYLE_NAMES.intextCitation[wps.Enum.msoLanguageIDEnglishUS]
}

export type BanyanFieldData = {
  id?: string
  type: string
  content: RichText
}

export type FieldAndData<T extends BanyanFieldData> = {
  field: Wps.Field
  data: T
}

export type IntextCitationFieldAndData = {
  field: Wps.Field
  data: IntextCitation
}

export function asStyleIdentifier(style: { id: string; title: string }): StyleIdentifier {
  return {
    id: style.id,
    title: style.title,
  }
}

export type NoteCitationFootnoteAndData = {
  note: Wps.Footnote
  field: Wps.Field
  data: NoteCitation
}

export type RebuiltNoteCitation = {
  note: Wps.Footnote
  field: Wps.Field
}

type RichTextLinkSegment = {
  start: number
  end: number
  link: string
}

export type unknownRecord = Record<string, unknown>
export function isUnknownRecord(v: unknown): v is unknownRecord {
  return v !== null && typeof v === "object"
}

function optionalTypeof(v: unknown, type: string): boolean {
  return v === undefined || typeof v === type
}

function isInlineMark(v: unknown, textLength: number): v is InlineMark {
  if (!isUnknownRecord(v)) return false
  if (typeof v.type !== "string") return false
  const start = v.start
  const end = v.end
  if (typeof start !== "number" || !Number.isInteger(start)) return false
  if (typeof end !== "number" || !Number.isInteger(end)) return false
  if (start < 0 || end <= start || end > textLength) return false

  switch (v.type) {
    case "bold":
    case "italic":
      return typeof v.value === "boolean"
    case "script":
      return v.value === "superscript" || v.value === "subscript"
    case "color":
    case "backgroundColor":
      return typeof v.value === "string" && HexColorIndex(v.value) !== null
    case "link":
      return typeof v.value === "string" && v.value.length > 0
    default:
      return false
  }
}

export function isRichText(v: unknown): v is RichText {
  if (!isUnknownRecord(v)) return false
  if (typeof v.text !== "string") return false
  if (!Array.isArray(v.marks)) return false
  const text = v.text
  return v.marks.every((mark) => isInlineMark(mark, text.length))
}

export function isBanyanFieldData(v: unknown): v is BanyanFieldData {
  return (
    isUnknownRecord(v) &&
    optionalTypeof(v.id, "string") &&
    typeof v.type === "string" &&
    isRichText(v.content)
  )
}

export function isCitation(v: unknown): v is Citation {
  return isBanyanFieldData(v) &&
    typeof v.id === "string" &&
    "source" in v &&
    // 性能考虑，暂不验证 source 的具体结构
    isUnknownRecord(v.source)

}

export function isIntextCitation(v: unknown): v is IntextCitation {
  return isCitation(v) && v.type === "intext-citation"

}

export function isNoteCitation(v: unknown): v is NoteCitation {
  return isCitation(v) &&
    v.type === "note-citation" &&
    "reference" in v && isRichText(v.reference)
}

export function isBibliographyTitle(v: unknown): v is BibliographyTitle {
  return isBanyanFieldData(v) &&
    v.type === "bibliography-title"
}

export function isBibliographyEntry(v: unknown): v is BibliographyEntry {
  return isBanyanFieldData(v) &&
    v.type === "bibliography-entry"
}

export function readFieldData<T extends BanyanFieldData>(field: Wps.Field): T | null {
  const data = field.Data
  try {
    return JSON.parse(data) as T
  }
  catch (e) {
    logError("Field", "Failed to parse field data.", e)
    return null
  }
}

export function renderField(field: Wps.Field, content?: RichText) : boolean {
  const resolvedContent = resolveFieldContent(field, content)
  if (!resolvedContent) return false
  const resultRange = field.Result
  renderRange(resultRange, resolvedContent)
  return true
}

export function renderStyledField(
  field: Wps.Field,
  applyFieldStyle: (field: Wps.Field) => void,
  content?: RichText,
): boolean {
  const resolvedContent = resolveFieldContent(field, content)
  if (!resolvedContent) return false

  const resultRange = field.Result
  writeRangeText(resultRange, resolvedContent)
  applyFieldStyle(field)
  applyRichTextStylesToRange(resultRange, resolvedContent)
  return true
}

function resolveFieldContent(field: Wps.Field, content?: RichText): RichText | null {
  if (isRichText(content)) {
    return content
  }
  const data = readFieldData(field)
  if (!data) return null
  if (!isRichText(data.content)) return null
  return data.content
}

export function renderRange(range: Wps.Range, content: RichText): void {
  writeRangeText(range, content)
  applyRichTextStylesToRange(range, content)
}

function writeRangeText(range: Wps.Range, content: RichText): void {
  range.Text = content.text
}

export function applyRichTextStylesToRange(range: Wps.Range, content: RichText): void {
  applyRichTextStyleToRange(range, content)
  applyRichTextLinksToRange(range, content)
}

function applyRichTextStyleToRange(range: Wps.Range, content: RichText): void {
  for (const mark of content.marks) {
    if (mark.type === "link") continue
    const segment = range.Duplicate
    segment.SetRange(range.Start + mark.start, range.Start + mark.end)
    applyInlineMarkStyle(segment, mark)
  }
}

export function applyRichTextLinksToRange(range: Wps.Range, content: RichText): void {
  const segments: RichTextLinkSegment[] = []
  const baseStart = range.Start
  for (const mark of content.marks) {
    if (mark.type !== "link") continue
    segments.push({
      start: baseStart + mark.start,
      end: baseStart + mark.end,
      link: mark.value,
    })
  }

  for (let i = segments.length - 1; i >= 0; i -= 1) {
    const linkSegment = segments[i]
    const segment = range.Duplicate
    segment.SetRange(linkSegment.start, linkSegment.end)
    applyRichTextLink(segment, linkSegment.link)
  }
}

export function applyInlineMarkStyle(range: Wps.Range, mark: Exclude<InlineMark, { type: "link" }>) {
  try {
    switch (mark.type) {
      case "bold":
        range.Font.Bold = mark.value ? wps.Enum.msoCTrue : wps.Enum.msoFalse
        break
      case "italic":
        range.Font.Italic = mark.value ? wps.Enum.msoCTrue : wps.Enum.msoFalse
        break
      case "script":
        if (mark.value === "superscript") {
          range.Font.Superscript = wps.Enum.msoCTrue
          range.Font.Subscript = wps.Enum.msoFalse
        }
        else {
          range.Font.Subscript = wps.Enum.msoCTrue
          range.Font.Superscript = wps.Enum.msoFalse
        }
        break
      case "color": {
        const color = HexColorIndex(mark.value)
        if (color !== null) range.Font.Color = color
        break
      }
      case "backgroundColor": {
        const background = HexColorIndex(mark.value)
        if (background !== null) range.Shading.ForegroundPatternColor = background
        break
      }
    }
  }
  catch (error) {
    logWarn("Field", "Failed to apply inline mark style.", error)
  }
}

function HexColorIndex(hex: string): number | null {
  if (!hex) return null
  const normalized = hex.trim().replace(/^#/, "")
  const expanded =
    normalized.length === 3
      ? normalized
          .split("")
          .map((ch) => ch + ch)
          .join("")
      : normalized
  if (!/^[0-9a-fA-F]{6}$/.test(expanded)) return null

  const r = expanded.slice(0, 2)
  const g = expanded.slice(2, 4)
  const b = expanded.slice(4, 6)
  // 由于历史遗留问题，VBA颜色值的格式是BRG，即0xBBGGRR
  // 参考https://club.excelhome.net/thread-1624509-1-1.html?_dsign=b803a3d6
  return parseInt(`0x${b}${g}${r}`, 16)
}

export function getBibliographyBookmarkName(entryId: string): string {
  return `Banyan_Entry_${entryId}`
}

function parseLinkTarget(link: string): { type: "bookmark" | "url"; target: string } | null {
  if (!link) return null

  const banyanProtocolMatch = link.match(/^banyan:\/\/entry\/(.+)$/)
  if (banyanProtocolMatch) {
    const entryId = banyanProtocolMatch[1]
    return { type: "bookmark", target: getBibliographyBookmarkName(entryId) }
  }

  if (link.match(/^https?:\/\//)) {
    return { type: "url", target: link }
  }

  return null
}

export function applyRichTextLink(range: Wps.Range, link: string): void {
  try {
    const linkTarget = parseLinkTarget(link)
    if (!linkTarget) return

    const doc = wps.ActiveDocument
    if (linkTarget.type === "bookmark") {
      doc.Hyperlinks.Add(
        range,
        undefined,
        linkTarget.target,
        undefined,
        undefined
      )
    }
    else if (linkTarget.type === "url") {
      doc.Hyperlinks.Add(
        range,
        linkTarget.target,
        undefined,
        undefined,
        undefined
      )
    }
  }
  catch (error) {
    logWarn("Field", "Failed to apply rich text link.", error)
  }
}

export function addBookmarkToField(field: Wps.Field, bookmarkName: string): void {
  try {
    const doc = wps.ActiveDocument
    const bookmarks = doc.Bookmarks

    // 如果书签已存在，先删除
    for (let i = 1; i <= bookmarks.Count; i++) {
      if (bookmarks.Item(i).Name === bookmarkName) {
        bookmarks.Item(i).Delete()
        break
      }
    }

    // 在 field 的 Result 范围添加书签
    bookmarks.Add(bookmarkName, field.Result)
  }
  catch (error) {
    logWarn("Field", `Failed to add bookmark \"${bookmarkName}\" to field.`, error)
  }
}

export function getCaretStart(): Wps.Range {
  const range = wps.Selection.Range.Duplicate
  range.Collapse(wps.Enum.wdCollapseStart)
  return range
}

export function getCaretEnd(): Wps.Range {
  const range = wps.Selection.Range.Duplicate
  range.Collapse(wps.Enum.wdCollapseEnd)
  return range
}

export type WordStyleType = "character" | "paragraph"

function resolveWordStyleType(styleType: WordStyleType): number {
  return styleType === "paragraph"
    ? wps.Enum.wdStyleTypeParagraph
    : wps.Enum.wdStyleTypeCharacter
}

export function applyStyleToField(
  field: Wps.Field,
  styleName: string,
  styleType: WordStyleType = "character",
): void {
  if (!styleName || styleName.trim() === "") return
  try {
    const styles = wps.ActiveDocument.Styles
    let style: Wps.Style | null = null

    for (let i = 1; i <= styles.Count; i++) {
      if (styles.Item(i).NameLocal === styleName) {
        style = styles.Item(i)
        break
      }
    }

    if (!style) {
      try {
        style = styles.Add(styleName, resolveWordStyleType(styleType))
        if (styleType === "character") {
          style.BaseStyle = wps.Enum.wdStyleDefaultParagraphFont
        }
        style.UnhideWhenUsed = true
        style.QuickStyle = true
      }
      catch (error) {
        logWarn("Field", `Failed to create style \"${styleName}\".`, error)
        return
      }
    }
    else {
      style.UnhideWhenUsed = true
      style.QuickStyle = true
    }

    const resultRange = field.Result
    resultRange.Style = styleName
  }
  catch (error) {
    logWarn("Field", `Failed to apply style \"${styleName}\" to field.`, error)
  }
}

export function applyBuiltInStyleToField(field: Wps.Field, builtInStyle: number): void {
  try {
    field.Result.Style = builtInStyle
  }
  catch (error) {
    logWarn("Field", "Failed to apply built-in style to field.", error)
  }
}

export function applyBuiltInStyleToRange(range: Wps.Range, builtInStyle: number): void {
  try {
    range.Style = builtInStyle
  }
  catch (error) {
    logWarn("Field", "Failed to apply built-in style to range.", error)
  }
}

export function ensureBuiltInStyleInQuickStyleGallery(builtInStyleConstant: number): void {
  try {
    const styles = wps.ActiveDocument.Styles
    const style = styles.Item(builtInStyleConstant)
    if (style) {
      style.UnhideWhenUsed = true
      style.QuickStyle = true
    }
  }
  catch (error) {
    logWarn("Field", "Failed to add built-in style to quick style gallery.", error)
  }
}

function createCollapsedRange(position: number): Wps.Range {
  const range = wps.ActiveDocument.Range().Duplicate
  range.SetRange(position, position)
  return range
}

export function removeFieldSafely(field: Wps.Field | null | undefined): void {
  if (!field) {
    return
  }

  try {
    if (field.Locked) {
      field.Locked = false
    }
    field.Delete()
  }
  catch {
    // Best effort cleanup for callers in rollback paths.
  }
}

export function removeFootnoteSafely(footnote: Wps.Footnote | null | undefined): void {
  if (!footnote) {
    return
  }

  try {
    footnote.Delete()
  }
  catch {
    // Best effort cleanup for callers in rollback paths.
  }
}

function deleteFieldOrThrow(field: Wps.Field): void {
  if (field.Locked) {
    field.Locked = false
  }
  field.Delete()
}

function deleteFootnoteOrThrow(footnote: Wps.Footnote): void {
  footnote.Delete()
}

export function createEmptyCitationSource(): CitationSource {
  return {
    cites: [],
    params: {},
  }
}

export function createEmptyRichText(): RichText {
  return { text: "", marks: [] }
}

function createPlaceholderRichText(text: string): RichText {
  return {
    text,
    marks: [{ type: "color", start: 0, end: text.length, value: FIELD_PLACEHOLDER_COLOR }],
  }
}

export function createPlaceholderIntextCitationData(
  id: string,
  source: CitationSource = createEmptyCitationSource(),
): IntextCitation {
  return {
    id,
    type: "intext-citation",
    source,
    content: PLACEHOLDER_INTEXT_CITATION_CONTENT,
  }
}

export function createPlaceholderNoteCitationData(
  id: string,
  source: CitationSource = createEmptyCitationSource(),
): NoteCitation {
  return {
    id,
    type: "note-citation",
    source,
    content: PLACEHOLDER_NOTE_CITATION_CONTENT,
    reference: createEmptyRichText(),
  }
}

function createNoteCitationFromIntext(data: IntextCitation): NoteCitation {
  return createPlaceholderNoteCitationData(data.id, data.source)
}

function createIntextCitationFromNote(data: NoteCitation): IntextCitation {
  return createPlaceholderIntextCitationData(data.id, data.source)
}

export function collectIntextCitationFieldsInRange(range: Wps.Range): IntextCitationFieldAndData[] {
  const pairs: IntextCitationFieldAndData[] = []
  for (let i = 1; i <= range.Fields.Count; i += 1) {
    const field = range.Fields.Item(i)
    if (field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }

    const data = readFieldData(field)
    if (!isIntextCitation(data)) {
      continue
    }

    pairs.push({ field, data })
  }
  return pairs
}

export function collectNoteCitationFootnotesInRange(range: Wps.Range): NoteCitationFootnoteAndData[] {
  const pairs: NoteCitationFootnoteAndData[] = []
  const footnotes = range.Footnotes
  for (let i = 1; i <= footnotes.Count; i += 1) {
    const note = footnotes.Item(i)
    if (note.Range.Fields.Count === 0) {
      continue
    }

    const field = note.Range.Fields.Item(1)
    if (!field || field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }

    const data = readFieldData(field)
    if (!isNoteCitation(data)) {
      continue
    }

    pairs.push({ note, field, data })
  }
  return pairs
}

export function createNoteCitationAtRange(range: Wps.Range, data: NoteCitation): {
  note: Wps.Footnote,
  field: Wps.Field,
} {
  range.Collapse(wps.Enum.wdCollapseEnd)
  const referenceText = data.reference.text.length > 0
    ? data.reference.text
    : undefined
  const note = wps.ActiveDocument.Footnotes.Add(
    range,
    referenceText,
  )
  if (data.reference.text.length > 0) {
    const referenceRange = note.Reference
    applyRichTextStylesToRange(referenceRange, data.reference)
  }

  const textRange = note.Range.Duplicate
  textRange.Collapse(wps.Enum.wdCollapseStart)
  const field = wps.ActiveDocument.Fields.Add(
    note.Range,
    wps.Enum.wdFieldAddin,
    `BANYAN_CITATION ${data.id}`,
  )
  field.Data = JSON.stringify(data)
  renderStyledField(field, applyNoteCitationStyle)

  return {
    note,
    field,
  }
}

export function rebuildNoteCitationAtRange(
  note: Wps.Footnote,
  field: Wps.Field,
  data: NoteCitation,
): RebuiltNoteCitation | null {
  try {
    if (!note || !field) return null

    const oldData = readFieldData<NoteCitation>(field)
    const oldReferenceText = isNoteCitation(oldData) ? oldData.reference.text : ""
    const newReferenceText = data.reference.text

    // A stable reference (including Word-managed automatic numbering) does
    // not require rebuilding the footnote. Replacing only the field leaves all
    // surrounding rich text and paragraph formatting untouched.
    if (oldReferenceText === newReferenceText) {
      const newField = replaceCitationFieldInNote(note, field, data)
      return newField ? { note, field: newField } : null
    }

    // A changed custom reference requires a new footnote. Keep both notes alive
    // while copying FormattedText so direct formatting and embedded fields are
    // transferred before the old note is removed.
    const insertPosition = note.Reference.End
    const insertRange = wps.ActiveDocument.Range().Duplicate
    insertRange.SetRange(insertPosition, insertPosition)
    const referenceText = newReferenceText.length > 0 ? newReferenceText : undefined
    const newNote = wps.ActiveDocument.Footnotes.Add(insertRange, referenceText)
    if (newReferenceText.length > 0) {
      applyRichTextStylesToRange(newNote.Reference, data.reference)
    }
    newNote.Range.FormattedText = note.Range.FormattedText
    note.Delete()

    const copiedField = findCitationFieldInNote(newNote)
    const newField = copiedField
      ? replaceCitationFieldInNote(newNote, copiedField, data)
      : createCitationFieldAtNoteStart(newNote, data)
    return newField ? { note: newNote, field: newField } : null
  }
  catch (error) {
    logWarn("Field", "Failed to rebuild note citation.", error)
    return null
  }
}

function replaceCitationFieldInNote(
  note: Wps.Footnote,
  field: Wps.Field,
  data: NoteCitation,
): Wps.Field | null {
  try {
    const beforeRange = note.Range.Duplicate
    beforeRange.End = field.Result.Start
    removeFieldSafely(field)

    // Footnote bodies are a separate story; derive insertion coordinates from
    // the live note range rather than ActiveDocument.Range().
    const insertRange = note.Range.Duplicate
    insertRange.SetRange(beforeRange.End, beforeRange.End)
    const newField = insertRange.Fields.Add(
      insertRange,
      wps.Enum.wdFieldAddin,
      `BANYAN_CITATION ${data.id}`,
      false,
    )
    newField.Data = JSON.stringify(data)
    renderStyledField(newField, applyNoteCitationStyle)
    return newField
  }
  catch (error) {
    logWarn("Field", "Failed to replace note citation field.", error)
    return null
  }
}

function findCitationFieldInNote(note: Wps.Footnote): Wps.Field | null {
  try {
    const fields = note.Range.Fields
    for (let i = 1; i <= fields.Count; i += 1) {
      const field = fields.Item(i)
      if (field.Type !== wps.Enum.wdFieldAddin) continue
      if (String(field.Code.Text ?? "").toUpperCase().includes("BANYAN_CITATION")) {
        return field
      }
    }
  }
  catch (error) {
    logWarn("Field", "Failed to find copied note citation field.", error)
  }
  return null
}

function createCitationFieldAtNoteStart(note: Wps.Footnote, data: NoteCitation): Wps.Field | null {
  try {
    const insertRange = note.Range.Duplicate
    insertRange.Collapse(wps.Enum.wdCollapseStart)
    const field = insertRange.Fields.Add(
      insertRange,
      wps.Enum.wdFieldAddin,
      `BANYAN_CITATION ${data.id}`,
      false,
    )
    field.Data = JSON.stringify(data)
    renderStyledField(field, applyNoteCitationStyle)
    return field
  }
  catch (error) {
    logWarn("Field", "Failed to create note citation field.", error)
    return null
  }
}

export function createIntextCitationAtRange(range: Wps.Range, data: IntextCitation): Wps.Field {
  range.Collapse(wps.Enum.wdCollapseEnd)
  const field = range.Fields.Add(
    range,
    wps.Enum.wdFieldAddin,
    `BANYAN_CITATION ${data.id}`,
    false,
  )
  field.Data = JSON.stringify(data)
  renderStyledField(field, applyIntextCitationStyle)
  return field
}

export function migrateIntextCitationsToNotes(range: Wps.Range): void {
  const citations = collectIntextCitationFieldsInRange(range)
  for (let i = citations.length - 1; i >= 0; i -= 1) {
    const { field, data } = citations[i]

    const insertPosition = field.Result.Start
    const convertedData = createNoteCitationFromIntext(data)
    deleteFieldOrThrow(field)

    const insertRange = createCollapsedRange(insertPosition)
    createNoteCitationAtRange(insertRange, convertedData)
  }
}

export function migrateNoteCitationsToIntext(range: Wps.Range): void {
  const citations = collectNoteCitationFootnotesInRange(range)
  for (let i = citations.length - 1; i >= 0; i -= 1) {
    const { note: footnote, data } = citations[i]

    const insertPosition = footnote.Reference.Start
    const convertedData = createIntextCitationFromNote(data)
    deleteFootnoteOrThrow(footnote)

    const insertRange = createCollapsedRange(insertPosition)
    createIntextCitationAtRange(insertRange, convertedData)
  }
}

export function applyIntextCitationStyle(field: Wps.Field): void {
  applyStyleToField(field, getIntextCitationStyleName())
}

export function applyNoteCitationStyle(field: Wps.Field): void {
  ensureBuiltInStyleInQuickStyleGallery(wps.Enum.wdStyleFootnoteReference)
  ensureBuiltInStyleInQuickStyleGallery(wps.Enum.wdStyleFootnoteText)
  applyBuiltInStyleToField(field, wps.Enum.wdStyleFootnoteText)
}

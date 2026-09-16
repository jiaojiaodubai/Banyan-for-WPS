
import type {
  BibliographyEntry,
  BibliographyLine,
  BibliographyTitle,
  Citation,
  CitationSource,
  IntextCitation,
  NoteCitation,
} from "../typings/style"
import type { StyleIdentifier } from "../typings/http"
import type { InlineMark, RichText } from "../typings/unit"
import { logError, logInfo, logWarn } from "./log"

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

/**
 * 域代码契约（与 VBA 对齐）：
 *
 * ```text
 * BANYAN_CITATION <id>       正文引注 / 脚注引注
 * BANYAN_BIBLIOGRAPHY <id>   书目标题行 / 书目题录行
 * <章节分隔符提示文本>        章节分隔符
 * ```
 *
 * 每个代码都携带数据 id（含占位域）；能解析出来的代码即被信任，因此收集阶段
 * 只按域代码分类，不读 `Field.Data`。
 */
export const CITATION_CODE_PREFIX = "BANYAN_CITATION"
export const BIBLIOGRAPHY_CODE_PREFIX = "BANYAN_BIBLIOGRAPHY"
// 章节分隔符保留可读提示文本作为域代码，这里只匹配其中的固定片段。
export const CHAPTER_BREAK_CODE_ZH = "Banyan章节分隔符"
export const CHAPTER_BREAK_CODE_EN = "Banyan chapter break"

export type BanyanFieldKind = "citation" | "bibliography" | "chapter"

export type BanyanFieldCode = {
  kind: BanyanFieldKind
  id: string
}

export function citationFieldCode(id: string): string {
  return `${CITATION_CODE_PREFIX} ${id}`
}

export function bibliographyFieldCode(id: string): string {
  return `${BIBLIOGRAPHY_CODE_PREFIX} ${id}`
}

/**
 * 把域代码解析为 kind + id；不是 Banyan 域时返回 null。
 * 契约之前写下的代码可能没有 id（此时 id 为空串），判定仍然成立。
 */
export function parseBanyanFieldCode(codeText: string): BanyanFieldCode | null {
  let code = String(codeText ?? "").trim()
  if (code.length === 0) return null
  if (/^ADDIN\s/i.test(code)) code = code.slice(6).trim()

  if (code.includes(CHAPTER_BREAK_CODE_ZH) || code.includes(CHAPTER_BREAK_CODE_EN)) {
    return { kind: "chapter", id: "" }
  }

  // 只对前缀部分做大小写归一：为了比较而复制整个代码字符串不划算（VBA 用
  // `Left$` 截取后比较，没有这个开销，因此那边不需要这层处理）。
  const head = code.slice(0, BIBLIOGRAPHY_CODE_PREFIX.length).toUpperCase()
  let kind: BanyanFieldKind
  if (head === BIBLIOGRAPHY_CODE_PREFIX) kind = "bibliography"
  else if (head.slice(0, CITATION_CODE_PREFIX.length) === CITATION_CODE_PREFIX) kind = "citation"
  else return null

  const spaceIndex = code.indexOf(" ")
  const id = spaceIndex > 0 ? code.slice(spaceIndex + 1).trim() : ""
  return { kind, id }
}

/** 读取并解析域代码；宿主读取失败或不是 Banyan 域时返回 null。 */
export function getBanyanFieldCode(field: Wps.Field): BanyanFieldCode | null {
  try {
    return parseBanyanFieldCode(String(field.Code.Text ?? ""))
  }
  catch {
    return null
  }
}

export function fieldHasCodeKind(field: Wps.Field, kind: BanyanFieldKind): boolean {
  const parsed = getBanyanFieldCode(field)
  return parsed !== null && parsed.kind === kind
}

/**
 * 一次读取 `Field.Data`，同时给出原始文本与解析结果：刷新时先用文本判定无变化
 * （懒解析），只有不同时才需要已解析的对象。
 */
export function readFieldDataWithText<T extends BanyanFieldData>(
  field: Wps.Field,
): { text: string; data: T | null } {
  const text = String(field.Data ?? "")
  if (text.length === 0) return { text, data: null }
  try {
    return { text, data: JSON.parse(text) as T }
  }
  catch {
    return { text, data: null }
  }
}

/** 从书目行数据读取 id；书目行（标题与题录）统一用 id 标识。 */
export function getBibliographyLineId(line: BibliographyLine): string {
  const id: unknown = line.id
  return typeof id === "string" ? id : ""
}

export type FieldAndData<T extends BanyanFieldData> = {
  field: Wps.Field
  data: T
}

export function asStyleIdentifier(style: { id: string; title: string }): StyleIdentifier {
  return {
    id: style.id,
    title: style.title,
  }
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

/**
 * 比较写入域结果的呈现结构。
 */
export function richTextEquals(current: RichText | null | undefined, next: RichText | null | undefined): boolean {
  if (!current || !next || !isRichText(current) || !isRichText(next)) return false
  if (current.text !== next.text || current.marks.length !== next.marks.length) return false

  for (let i = 0; i < current.marks.length; i += 1) {
    const currentMark = current.marks[i]
    const nextMark = next.marks[i]
    if (
      currentMark.type !== nextMark.type
      || currentMark.start !== nextMark.start
      || currentMark.end !== nextMark.end
      || !Object.is(currentMark.value, nextMark.value)
    ) {
      return false
    }
  }
  return true
}

export function fieldContentEquals(
  current: BanyanFieldData | null | undefined,
  next: BanyanFieldData | null | undefined,
): boolean {
  return richTextEquals(current?.content, next?.content)
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
  // 与 `renderStyledFieldWithData` 同理：决定清除直接字符格式的方式。
  styleType: WordStyleType = "paragraph",
): boolean {
  const resolvedContent = resolveFieldContent(field, content)
  if (!resolvedContent) return false

  return renderStyledFieldCore(field, applyFieldStyle, resolvedContent, styleType)
}

/**
 * 用于调用方已经持有已解析数据的场景：刷新与插入路径手里本来就有响应对象，
 * 避免再解析一次 `Field.Data` 的单次开销很小，但对大书目与大批引注就不容忽视。
 *
 * `styleType` 决定清除直接字符格式的方式：字符样式赋值自带清除，段落样式赋值不会（WPS
 * 实测），后者需要在渲染前显式复位。
 */
export function renderStyledFieldWithData<T extends BanyanFieldData>(
  field: Wps.Field,
  applyFieldStyle: (field: Wps.Field) => void,
  data: T,
  content?: RichText,
  styleType: WordStyleType = "paragraph",
): boolean {
  const resolvedContent = isRichText(content) ? content : data.content
  if (!isRichText(resolvedContent)) return false

  return renderStyledFieldCore(field, applyFieldStyle, resolvedContent, styleType)
}

function renderStyledFieldCore(
  field: Wps.Field,
  applyFieldStyle: (field: Wps.Field) => void,
  content: RichText,
  styleType: WordStyleType,
): boolean {
  const resultRange = field.Result
  writeRangeText(resultRange, content)
  // 同 VBA 的 RenderStyledFieldCore：渲染结果后隐藏域代码。
  field.ShowCodes = false
  // 写文本会继承结果上的旧直接格式；字符样式路径由 `applyFieldStyle` 的赋值清掉，
  // 段落样式赋值不清除，先显式复位。
  if (styleType === "paragraph") clearDirectCharacterFormatting(resultRange)
  applyFieldStyle(field)
  applyRichTextStylesToRange(resultRange, content)
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

/**
 * 清除域结果上继承的直接字符格式，供段落样式路径在渲染前复位（字符样式赋值自带清除）。
 */
function clearDirectCharacterFormatting(range: Wps.Range): void {
  const font = range.Font
  font.Color = wps.Enum.wdColorAutomatic
  font.Bold = 0
  font.Italic = 0
  font.Subscript = 0
  font.Superscript = 0
  font.SmallCaps = 0
  font.AllCaps = 0
  range.Shading.BackgroundPatternColor = wps.Enum.wdColorAutomatic
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

const BIBLIOGRAPHY_BOOKMARK_PREFIX = "Banyan_Entry_"
const MAX_BOOKMARK_NAME_LENGTH = 40

/**
 * 归一化 Word/WPS 书签名。
 *
 * 书签名必须以字母开头，只能包含字母、数字与下划线，且不超过 40 个字符。这里
 * 刻意与 VBA 实现保持一致，使 WPS 生成的书签名与 Word 加载项生成的一致。
 * 约束依据见 WordRibbon T008718 与 Microsoft Q&A 5161621。
 */
export function normalizeBookmarkName(name: string): string {
  const source = String(name ?? "").trim()
  if (!source) return ""

  let normalized = ""
  for (let i = 0; i < source.length; i += 1) {
    const character = source[i]
    normalized += /^[A-Za-z0-9_]$/.test(character) ? character : "_"
  }

  if (!/^[A-Za-z]/.test(normalized)) {
    normalized = `B${normalized}`
  }
  return normalized.slice(0, MAX_BOOKMARK_NAME_LENGTH)
}

export function getBibliographyBookmarkName(entryId: string): string {
  return normalizeBookmarkName(`${BIBLIOGRAPHY_BOOKMARK_PREFIX}${String(entryId ?? "")}`)
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
      // 不检查书签是否存在（同 VBA）：引注常先于书目渲染，宿主会先接受 SubAddress、
      // 稍后解析；提前跳过会让该引注永久失去链接。
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
    const normalizedName = normalizeBookmarkName(bookmarkName)
    if (!normalizedName) return

    const resultRange = field.Result

    // 如果书签已存在，先删除
    if (bookmarks.Exists(normalizedName)) {
      bookmarks.Item(normalizedName).Delete()
    }

    // 在 field 的 Result 范围添加书签
    bookmarks.Add(normalizedName, resultRange)
  }
  catch (error) {
    logWarn("Field", `Failed to add bookmark \"${bookmarkName}\" to field.`, error)
  }
}

/** 把光标折叠到 `position` 处。 */
export function moveCaretToPosition(position: number): void {
  const caret = wps.ActiveDocument.Range().Duplicate
  caret.SetRange(position, position)
  caret.Select()
}

/**
 * 把光标移回正文中脚注引用之后。
 *
 * WPS 的 `Footnotes.Add` 会把插入点留在脚注里（Word 不会），创建/重建脚注后需要回移。
 */
export function restoreMainTextAfterNote(note: Wps.Footnote): void {
  moveCaretToPosition(note.Reference.End)
}

/**
 * 取“整个域”的范围：`Range(Code.Start - 1, Result.End + 1)`，对应 VBA 的
 * `BibliographyWholeFieldRange`。
 *
 * `Field.Result` 默认落在域内部（`Result.End` 就是域结束标记的位置），要在域后插入内容
 * 必须用这个范围。
 */
export function getWholeFieldRange(field: Wps.Field): Wps.Range {
  const range = wps.ActiveDocument.Range().Duplicate
  range.SetRange(field.Code.Start - 1, field.Result.End + 1)
  return range
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

/**
 * 样式类别。`renderStyledField*` 用它决定清除直接字符格式的方式（拼写与 VBA 的
 * `FieldRenderStyledFieldWithStyle` 同名参数相同，但那边只用于查找/创建样式，
 * 不决定清除策略，移植时不要混淆）。
 */
export type WordStyleType = "character" | "paragraph"

function resolveWordStyleType(styleType: WordStyleType): number {
  return styleType === "paragraph"
    ? wps.Enum.wdStyleTypeParagraph
    : wps.Enum.wdStyleTypeCharacter
}

/**
 * 按名称取段落/字符样式，不存在则创建。
 */
function getOrCreateStyle(styleName: string, styleType: WordStyleType): Wps.Style | null {
  const styles = wps.ActiveDocument.Styles
  let style: Wps.Style | null = null
  try {
    const byName = (styles.Item as unknown as (key: string) => Wps.Style)(styleName)
    if (byName && byName.NameLocal === styleName) style = byName
  }
  catch {
    // 宿主可能只接受数字索引。
  }

  if (!style) {
    for (let i = 1; i <= styles.Count; i += 1) {
      if (styles.Item(i).NameLocal === styleName) {
        style = styles.Item(i)
        break
      }
    }
  }

  if (!style) {
    style = styles.Add(styleName, resolveWordStyleType(styleType))
    if (styleType === "character") {
      style.BaseStyle = wps.Enum.wdStyleDefaultParagraphFont
    }
  }

  style.UnhideWhenUsed = true
  style.QuickStyle = true
  return style
}

function getAppliedStyleName(field: Wps.Field): string {
  try {
    const style = field.Result.Style as unknown
    if (typeof style === "string") return style
    if (style && typeof style === "object" && "NameLocal" in style) {
      return String((style as { NameLocal?: unknown }).NameLocal ?? "")
    }
  }
  catch {
    // 读不到样式名时不能阻止后面的纠正性赋值。
  }
  return ""
}

/**
 * 套用段落/字符样式，返回是否真正写入了 `Range.Style`。
 *
 * 字符样式每轮都要写：赋值本身会清掉域结果上的直接字符格式（同名重复赋值也清，WPS
 * 实测），这一步就是渲染前的清除。段落样式赋值不清直接字符格式，因此样式名相同时跳过
 * 写入（`Range.Style` 是 WPS 里最贵的一类调用），清除交给渲染路径显式复位。
 */
export function applyStyleToField(
  field: Wps.Field,
  styleName: string,
  styleType: WordStyleType = "character",
): boolean {
  if (!styleName || styleName.trim() === "") return false
  try {
    // 读不出来样式名时按“不匹配”处理，保持纠正性赋值。
    if (styleType === "paragraph" && getAppliedStyleName(field) === styleName) return false

    const style = getOrCreateStyle(styleName, styleType)
    if (!style) return false
    field.Result.Style = styleName
    return true
  }
  catch (error) {
    logWarn("Field", `Failed to apply style "${styleName}" to field.`, error)
    return false
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
    // 回滚路径下的尽力清理。
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
    // 回滚路径下的尽力清理。
  }
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

export type CitationFieldPair = {
  id: string
  field: Wps.Field
}

export type NoteCitationFieldPair = {
  id: string
  note: Wps.Footnote
  field: Wps.Field
}

export function collectIntextCitationFieldsInRange(range: Wps.Range): CitationFieldPair[] {
  const pairs: CitationFieldPair[] = []
  const seen = new Set<string>()
  const broken: Wps.Field[] = []
  // 集合与计数只取一次：重键会改写代码，但不会增删域。
  const fields = range.Fields
  const count = fields.Count
  for (let i = 1; i <= count; i += 1) {
    const field = fields.Item(i)
    if (field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }

    // 只按域代码分类：不读 Field.Data。
    const parsed = getBanyanFieldCode(field)
    if (!parsed || parsed.kind !== "citation") {
      continue
    }

    if (parsed.id.length === 0) {
      // 代码里的 id 丢了：这不是一个引注，删掉整个域。
      broken.push(field)
      continue
    }

    pairs.push({ id: uniqueCitationId(field, parsed.id, seen), field })
  }

  // 遍历期间删除会打乱域集合，因此收集完再倒序删。
  for (let i = broken.length - 1; i >= 0; i -= 1) {
    removeFieldSafely(broken[i])
  }
  return pairs
}

export function collectNoteCitationFootnotesInRange(range: Wps.Range): NoteCitationFieldPair[] {
  const pairs: NoteCitationFieldPair[] = []
  const seen = new Set<string>()
  const broken: Wps.Footnote[] = []
  const footnotes = range.Footnotes
  const count = footnotes.Count
  for (let i = 1; i <= count; i += 1) {
    const note = footnotes.Item(i)
    if (note.Range.Fields.Count === 0) {
      continue
    }

    const field = note.Range.Fields.Item(1)
    if (!field || field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }

    const parsed = getBanyanFieldCode(field)
    if (!parsed || parsed.kind !== "citation") {
      continue
    }

    if (parsed.id.length === 0) {
      // 脚注引注损坏时连同整个脚注一起删除。
      broken.push(note)
      continue
    }

    pairs.push({ id: uniqueCitationId(field, parsed.id, seen), note, field })
  }

  for (let i = broken.length - 1; i >= 0; i -= 1) {
    removeFootnoteSafely(broken[i])
  }
  return pairs
}

/**
 * 把新的引注 id 写回域代码。
 *
 * 改写 ADDIN 域代码时宿主会丢弃 `Field.Data`，因此先把已存数据读成文本，换完
 * 代码再写回去（只对确实重键的副本做一次字符串拷贝）。
 */
export function writeCitationCodeId(field: Wps.Field, id: string): boolean {
  try {
    const storedText = String(field.Data ?? "")
    field.Code.Text = ` ADDIN ${citationFieldCode(id)} `
    if (storedText.length > 0) field.Data = storedText
    return true
  }
  catch (error) {
    logWarn("Field", "Failed to rewrite citation field code id.", error)
    return false
  }
}

/**
 * 保证引注 id 唯一：首次出现保留，重复的（用户复制粘贴）在域代码里换成新 id。
 * 代码改不动时退回原 id，让后续按 id 匹配的行为保持原样。
 */
function uniqueCitationId(field: Wps.Field, id: string, seen: Set<string>): string {
  if (!seen.has(id)) {
    seen.add(id)
    return id
  }

  let candidate = crypto.randomUUID()
  while (seen.has(candidate)) candidate = crypto.randomUUID()

  if (!writeCitationCodeId(field, candidate)) return id

  seen.add(candidate)
  return candidate
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

  // 域插进折叠到起点的脚注正文范围（同 VBA）：非折叠范围会把域插到段落标记之后，
  // 脚注正文就是空的。
  const textRange = note.Range.Duplicate
  textRange.Collapse(wps.Enum.wdCollapseStart)
  const field = textRange.Fields.Add(
    textRange,
    wps.Enum.wdFieldAddin,
    citationFieldCode(data.id),
    false,
  )
  field.Data = JSON.stringify(data)
  renderStyledFieldWithData(field, applyNoteCitationStyle, data, data.content)

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
    const oldReference = isNoteCitation(oldData) ? oldData.reference : null

    // 源信息等元数据不影响脚注呈现：仅落入数据即可，不必替换域或脚注。
    if (
      isNoteCitation(oldData)
      && fieldContentEquals(oldData, data)
      && richTextEquals(oldReference, data.reference)
    ) {
      field.Data = JSON.stringify(data)
      return { note, field }
    }

    const newReferenceText = data.reference.text

    // 引用标记稳定（含 Word 托管的自动编号）时无需重建脚注：只替换域即可，
    // 周围的富文本与段落格式都不受影响。
    if (oldReference && richTextEquals(oldReference, data.reference)) {
      const newField = replaceCitationFieldInNote(note, field, data)
      return newField ? { note, field: newField } : null
    }

    // 自定义引用标记变了就必须新建脚注：复制 FormattedText 时保持新旧脚注同时
    // 存活，让直接格式与内嵌域先完成转移，再删掉旧脚注。
    // 钳制同 VBA：越界位置在 WPS 会静默落到 `0-0`，新脚注会插到文档开头。
    const insertPosition = clampInsertPosition(note.Reference.End)
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

    // 脚注正文属于另一个 story：插入坐标要从活动脚注范围推算，而不是用
    // ActiveDocument.Range()。
    const insertRange = note.Range.Duplicate
    insertRange.SetRange(beforeRange.End, beforeRange.End)
    const newField = insertRange.Fields.Add(
      insertRange,
      wps.Enum.wdFieldAddin,
      citationFieldCode(data.id),
      false,
    )
    newField.Data = JSON.stringify(data)
    renderStyledFieldWithData(newField, applyNoteCitationStyle, data, data.content)
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
      if (fieldHasCodeKind(field, "citation")) {
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
      citationFieldCode(data.id),
      false,
    )
    field.Data = JSON.stringify(data)
    renderStyledFieldWithData(field, applyNoteCitationStyle, data, data.content)
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
    citationFieldCode(data.id),
    false,
  )
  field.Data = JSON.stringify(data)
  renderStyledFieldWithData(field, applyIntextCitationStyle, data, data.content, "character")
  return field
}

export function migrateIntextCitationsToNotes(range: Wps.Range): void {
  const citations = collectIntextCitationFieldsInRange(range)
  logInfo("Field", `Migrating ${citations.length} intext citation(s) to note citations.`)
  let converted = 0
  for (let i = citations.length - 1; i >= 0; i -= 1) {
    const { field } = citations[i]
    // 收集阶段不再读数据；这里确实需要数据才读取。
    const data = readFieldData<IntextCitation>(field)
    if (!isIntextCitation(data)) {
      logWarn("Field", "Skipped an intext citation without readable data during migration.")
      continue
    }

    // 插入位置取域起点 `Code.Start - 1`（域代码从起始标记之后一个字符开始）；删除域
    // 会连代码一起删掉、域内位置会偏移，域起点不会。反向遍历使前面的引注不受影响。
    const insertPosition = field.Code.Start - 1
    const convertedData = createNoteCitationFromIntext(data)
    // 与 VBA 一致，用安全版本删除：单个引注失败不影响其余引注。
    removeFieldSafely(field)

    // 插入点在待删域之前取得，删除不会使文档在该点之前变短，无需钳制。
    const insertRange = createCollapsedRange(insertPosition)
    try {
      const created = createNoteCitationAtRange(insertRange, convertedData)
      converted += 1
      // 创建脚注会把插入点留在脚注里；立即回移到引用之后。
      restoreMainTextAfterNote(created.note)
    }
    catch (error) {
      logWarn("Field", `Failed to create a note citation for ${data.id} during migration.`, error)
    }
  }
  logInfo("Field", `Migrated ${converted} of ${citations.length} intext citation(s) to note citations.`)
}

export function migrateNoteCitationsToIntext(range: Wps.Range): void {
  const citations = collectNoteCitationFootnotesInRange(range)
  logInfo("Field", `Migrating ${citations.length} note citation(s) to intext citations.`)
  let converted = 0
  for (let i = citations.length - 1; i >= 0; i -= 1) {
    const { note: footnote, field } = citations[i]
    const data = readFieldData<NoteCitation>(field)
    if (!isNoteCitation(data)) {
      logWarn("Field", "Skipped a note citation without readable data during migration.")
      continue
    }

    // 引用标记只占一个字符，删除脚注不会移动它的起点，无需活动锚点。
    const insertPosition = footnote.Reference.Start
    const convertedData = createIntextCitationFromNote(data)
    removeFootnoteSafely(footnote)

    // 插入点取自在脚注引用处，删除脚注会让文档变短，需要钳制（同 VBA）。
    const insertRange = createCollapsedRange(clampInsertPosition(insertPosition))
    try {
      createIntextCitationAtRange(insertRange, convertedData)
      converted += 1
    }
    catch (error) {
      logWarn("Field", `Failed to create an intext citation for ${data.id} during migration.`, error)
    }
  }
  logInfo("Field", `Migrated ${converted} of ${citations.length} note citation(s) to intext citations.`)
}

/**
 * 把位置限制在 `[Content.Start, Content.End - 1]` 内（同 VBA 的 ClampInsertPosition）：
 * 删除文档末尾的域/脚注会让文档变短，越界位置在 WPS 会静默落到 `0-0`（插到开头）。
 */
function clampInsertPosition(position: number): number {
  const content = wps.ActiveDocument.Content
  let clamped = position
  if (content.End > content.Start && clamped > content.End - 1) {
    clamped = content.End - 1
  }
  if (clamped < content.Start) {
    clamped = content.Start
  }
  return clamped
}

export function applyIntextCitationStyle(field: Wps.Field): void {
  applyStyleToField(field, getIntextCitationStyleName())
}

export function applyNoteCitationStyle(field: Wps.Field): void {
  ensureBuiltInStyleInQuickStyleGallery(wps.Enum.wdStyleFootnoteReference)
  ensureBuiltInStyleInQuickStyleGallery(wps.Enum.wdStyleFootnoteText)
  applyBuiltInStyleToField(field, wps.Enum.wdStyleFootnoteText)
}

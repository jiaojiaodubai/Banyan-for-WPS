import { request, getDocumentId } from "../utils/http"
import { useI10n } from "../utils/i10n"
import { withProgress } from "../utils/progress"
import { withBatchUpdate } from "../utils/batch-update"
import {
  FIELD_PLACEHOLDER_COLOR,
  addBookmarkToField,
  applyStyleToField,
  asStyleIdentifier,
  bibliographyFieldCode,
  fieldContentEquals,
  fieldHasCodeKind,
  getBanyanFieldCode,
  getBibliographyBookmarkName,
  getBibliographyLineId,
  getWholeFieldRange,
  isBibliographyEntry,
  isBibliographyTitle,
  isIntextCitation,
  isNoteCitation,
  readFieldData,
  readFieldDataWithText,
  removeFieldSafely,
  renderStyledFieldWithData,
} from "../utils/field"
import type { BibliographyTitle, CitationContext, BibliographyLine, IntextCitation, NoteCitation } from "../typings/style"
import { getPreference, savePreference } from "./preference"
import { getUpdateRange } from "./chapter-break"
import { logError, logWarn } from "../utils/log"

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
  styleProgressReason: "正在应用书目样式...",
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
  styleProgressReason: "Applying bibliography styles...",
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

function createPendingBibliographyLine(): BibliographyTitle {
  const text = "{ BIBLIOGRAPHY }"
  return {
    id: crypto.randomUUID(),
    type: "bibliography-title",
    content: {
      text,
      marks: [
        {
          type: "color",
          start: 0,
          end: text.length,
          value: FIELD_PLACEHOLDER_COLOR,
        },
      ],
    },
  }
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
          const inPlaceResult = updateBibliographyInPlace(updateRange, lines, pref)
          if (inPlaceResult === null) {
            deleteExistingBibliography(updateRange)
            insertBibliography(insertRange, lines, pref)
          }
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
    const contentChanged = !fieldContentEquals(currentLine, response.line)
    mode.field.Data = JSON.stringify(response.line)
    if (contentChanged) {
      renderStyledFieldWithData(
        mode.field,
        (field) => applyStyleToField(field, pref.bibliographyEntryStyle, "paragraph"),
        response.line,
        response.line.content,
      )
    }
    addBookmarkToField(mode.field, getBibliographyBookmarkName(response.line.id))
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

  // 与收集器一致：只按域代码分类，命中后才读数据；集合与计数只取一次。
  const fields = range.Fields
  const fieldCount = fields.Count
  for (let i = 1; i <= fieldCount; i += 1) {
    const field = fields.Item(i)
    if (field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }
    const parsed = getBanyanFieldCode(field)
    if (!parsed || parsed.kind !== "citation" || parsed.id.length === 0) {
      continue
    }
    const data = readFieldData<IntextCitation>(field)
    if (!data || !isIntextCitation(data)) {
      continue
    }
    const page = field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
    contexts.set(parsed.id, {
      id: parsed.id,
      page,
      ...data.source,
    })
  }

  const footnotes = range.Footnotes
  const noteCount = footnotes.Count
  for (let i = 1; i <= noteCount; i += 1) {
    const note = footnotes.Item(i)
    if (note.Range.Fields.Count === 0) {
      continue
    }
    const field = note.Range.Fields.Item(1)
    if (!field || field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }
    const parsed = getBanyanFieldCode(field)
    if (!parsed || parsed.kind !== "citation" || parsed.id.length === 0) {
      continue
    }
    const data = readFieldData<NoteCitation>(field)
    if (!data || !isNoteCitation(data)) {
      continue
    }
    const page = field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
    contexts.set(parsed.id, {
      id: parsed.id,
      page,
      ...data.source,
    })
  }
  return [...contexts.values()]
}

export function collectBibliographyFieldsInRange(range: Wps.Range): Wps.Field[] {
  return collectBibliographyFieldsCore(range, true)
}

/**
 * 书目行由样式生成、不由用户直接编辑：id 缺失（代码被改坏）或已被占用（用户
 * 粘贴了一行）的行根本不算行，整行删除后继续处理其余行。
 *
 * 删除会移动后面所有内容，因此删完重新收集一次；第二遍不再删非法行，让它们
 * 交给下游的 id 校验处理，而不是产出一个错误的 diff。
 */
function collectBibliographyFieldsCore(range: Wps.Range, dropIllegal: boolean): Wps.Field[] {
  const result: Wps.Field[] = []
  const seen = new Set<string>()
  const illegal: Wps.Field[] = []

  // 集合与计数只取一次；删除在循环之后进行，遍历期间集合不变。
  const fields = range.Fields
  const count = fields.Count
  for (let i = 1; i <= count; i += 1) {
    const field = fields.Item(i)
    if (field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }
    // 只按域代码分类：不读 Field.Data（与 VBA 的收集器一致）。
    const parsed = getBanyanFieldCode(field)
    if (!parsed || parsed.kind !== "bibliography") {
      continue
    }

    if (parsed.id.length === 0 || seen.has(parsed.id)) {
      if (dropIllegal) illegal.push(field)
      else result.push(field)
      continue
    }

    seen.add(parsed.id)
    result.push(field)
  }

  if (illegal.length > 0) {
    deleteBibliographyLines(illegal)
    return collectBibliographyFieldsCore(range, false)
  }
  return result
}

/**
 * 删除一整行：域（域代码 + 域结果）加上紧随其后的段落分隔符。
 *
 * 只删这一行本身，而不是整个段落：段落里若混入了别的内容，把它们留在原地即可
 * ——书目行按设计自成一段，即使真的被误删，后续刷新也会重建。尾随分隔符不存在
 * 或不是段落标记时只删域本身。倒序删除避免前面的删除使后面的域位置引用失效。
 */
function deleteBibliographyLines(fields: Wps.Field[]): void {
  for (let i = fields.length - 1; i >= 0; i -= 1) {
    const field = fields[i]
    try {
      if (field.Locked) field.Locked = false
      const start = Math.min(field.Code.Start, field.Result.Start)
      const resultEnd = field.Result.End

      const line = wps.ActiveDocument.Range().Duplicate
      line.SetRange(start, resultEnd + 1)
      if (!String(line.Text ?? "").endsWith("\r")) {
        line.SetRange(start, resultEnd)
      }
      line.Delete()
    }
    catch (error) {
      logWarn("Bibliography", "Failed to delete an illegal bibliography line.", error)
    }
  }
}

/** 书目行 id 来自域代码（契约）。 */
export function getBibliographyFieldId(field: Wps.Field): string {
  const parsed = getBanyanFieldCode(field)
  return parsed && parsed.kind === "bibliography" ? parsed.id : ""
}

type BibliographyMatch = { oldIndex: number; nextIndex: number }

type BibliographyBlockRange = { start: number; end: number }

/**
 * 以最小结构差异更新已有书目：未变化的域（含用户手工格式）保留，插入或删除的
 * 片段就地替换。返回 `null` 表示现有块无法安全建立索引，调用方应改走全量重建。
 *
 * 响应决定书目的最终性状：文档中未被 LIS 匹配到的现有行（旧版前端残留的空
 * id 域、用户复制出的重复 id 行）会随空隙修补一并删除，必要时整块替换，最终
 * 形状与全量刷新一致。
 */
export function updateBibliographyInPlace(
  range: Wps.Range,
  lines: BibliographyLine[],
  pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string },
): boolean | null {
  const fields = collectBibliographyFieldsInRange(range)
  if (fields.length === 0 || lines.length === 0) return null

  const currentIds = fields.map((field) => getBibliographyFieldId(field))
  const nextIds = lines.map((line) => getBibliographyLineId(line))
  const nextPositions = new Map<string, number>()
  for (let i = 0; i < nextIds.length; i += 1) {
    nextPositions.set(nextIds[i], i)
  }

  const structureChanged = currentIds.length !== nextIds.length
    || currentIds.some((id, index) => id !== nextIds[index])
  const matches = findBibliographyMatches(currentIds, nextPositions)

  if (structureChanged) {
    if (matches.length === 0) {
      return replaceBibliographyBlock(fields, lines, pref)
    }
    if (!patchBibliographyGaps(fields, lines, pref, matches)) {
      return null
    }
  }

  let changed = structureChanged
  for (const match of matches) {
    const field = fields[match.oldIndex]
    const next = lines[match.nextIndex]
    if (updateBibliographyField(field, next, pref)) {
      changed = true
    }
  }
  return changed
}

function findBibliographyMatches(
  currentIds: string[],
  nextPositions: Map<string, number>,
): BibliographyMatch[] {
  const sequence: number[] = []
  const oldIndexes: number[] = []
  for (let oldIndex = 0; oldIndex < currentIds.length; oldIndex += 1) {
    const nextIndex = nextPositions.get(currentIds[oldIndex])
    if (nextIndex === undefined) continue
    sequence.push(nextIndex)
    oldIndexes.push(oldIndex)
  }
  if (sequence.length === 0) return []

  // 最长递增子序列：最大化保留现有域，只修补它们之间的空隙。
  const tails: number[] = []
  const tailSequenceIndexes: number[] = []
  const previous: number[] = new Array(sequence.length).fill(-1)
  for (let sequenceIndex = 0; sequenceIndex < sequence.length; sequenceIndex += 1) {
    let low = 0
    let high = tails.length
    while (low < high) {
      const middle = Math.floor((low + high) / 2)
      if (tails[middle] < sequence[sequenceIndex]) low = middle + 1
      else high = middle
    }
    if (low > 0) previous[sequenceIndex] = tailSequenceIndexes[low - 1]
    tails[low] = sequence[sequenceIndex]
    tailSequenceIndexes[low] = sequenceIndex
  }

  const result: BibliographyMatch[] = []
  let sequenceIndex = tailSequenceIndexes[tails.length - 1]
  while (sequenceIndex !== undefined && sequenceIndex >= 0) {
    result.push({ oldIndex: oldIndexes[sequenceIndex], nextIndex: sequence[sequenceIndex] })
    sequenceIndex = previous[sequenceIndex]
  }
  result.reverse()
  return result
}

function getBibliographyBlockRange(field: Wps.Field): BibliographyBlockRange | null {
  try {
    const documentStart = wps.ActiveDocument.Content.Start
    const documentEnd = wps.ActiveDocument.Content.End
    const start = Math.max(
      documentStart,
      Math.min(field.Code.Start, field.Result.Start) - 1,
    )
    // 连同书目域之间的段落分隔符一起纳入范围。
    const end = Math.min(documentEnd, field.Result.End + 1)
    return end >= start ? { start, end } : null
  }
  catch {
    return null
  }
}

function patchBibliographyGaps(
  fields: Wps.Field[],
  lines: BibliographyLine[],
  pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string },
  matches: BibliographyMatch[],
): boolean {
  for (let gap = matches.length; gap >= 0; gap -= 1) {
    const previousOld = gap === 0 ? -1 : matches[gap - 1].oldIndex
    const previousNext = gap === 0 ? -1 : matches[gap - 1].nextIndex
    const followingOld = gap === matches.length ? fields.length : matches[gap].oldIndex
    const followingNext = gap === matches.length ? lines.length : matches[gap].nextIndex
    const oldGapCount = followingOld - previousOld - 1
    const nextGapCount = followingNext - previousNext - 1
    if (oldGapCount === 0 && nextGapCount === 0) continue

    const firstOld = previousOld < 0 ? fields[0] : fields[previousOld]
    const lastOld = followingOld >= fields.length ? fields[fields.length - 1] : fields[followingOld]
    const firstRange = getBibliographyBlockRange(firstOld)
    const lastRange = getBibliographyBlockRange(lastOld)
    if (!firstRange || !lastRange) return false

    const start = previousOld < 0 ? firstRange.start : getBibliographyBlockRange(fields[previousOld])?.end
    const end = followingOld >= fields.length ? lastRange.end : getBibliographyBlockRange(fields[followingOld])?.start
    if (start === undefined || end === undefined || end < start) return false

    const gapRange = wps.ActiveDocument.Range().Duplicate
    gapRange.SetRange(start, end)
    if (end > start) gapRange.Delete()

    const insertRange = wps.ActiveDocument.Range().Duplicate
    insertRange.SetRange(start, start)
    if (!insertBibliographyLines(
      insertRange,
      lines,
      previousNext + 1,
      followingNext - 1,
      pref,
      previousOld > 0,
      followingOld < fields.length && nextGapCount > 0,
    )) {
      return false
    }
  }
  return true
}

function replaceBibliographyBlock(
  fields: Wps.Field[],
  lines: BibliographyLine[],
  pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string },
): boolean {
  const firstRange = getBibliographyBlockRange(fields[0])
  const lastRange = getBibliographyBlockRange(fields[fields.length - 1])
  if (!firstRange || !lastRange) return false

  const block = wps.ActiveDocument.Range().Duplicate
  block.SetRange(firstRange.start, lastRange.end)
  block.Delete()
  const insertRange = wps.ActiveDocument.Range().Duplicate
  insertRange.SetRange(firstRange.start, firstRange.start)
  return insertBibliographyLines(insertRange, lines, 0, lines.length - 1, pref, false, false)
}

function updateBibliographyField(
  field: Wps.Field,
  next: BibliographyLine,
  pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string },
): boolean {
  const nextJson = JSON.stringify(next)
  if (nextJson.length === 0) return false

  // 一次读取：先用文本判定无变化（懒解析），不同才需要解析后的对象。
  const { text, data: current } = readFieldDataWithText<BibliographyLine>(field)
  // 没有数据（残留域）或数据无法解析时保持原样：不去猜它本来应该是什么。
  if (text.length === 0 || !current) return false
  // 直接用 `===`：JS 的字符串比较自带长度短路，不需要 VBA 那样的 FieldTextEquals 封装。
  if (text === nextJson) return false

  const contentChanged = !fieldContentEquals(current, next)
  const expectedStyle = isBibliographyTitle(next)
    ? pref.bibliographyTitleStyle
    : pref.bibliographyEntryStyle

  field.Data = nextJson
  if (contentChanged) {
    renderStyledFieldWithData(
      field,
      (target) => applyStyleToField(target, expectedStyle, "paragraph"),
      next,
      next.content,
    )
  }
  if (isBibliographyEntry(next)) {
    addBookmarkToField(field, getBibliographyBookmarkName(next.id))
  }
  return true
}

/**
 * 把章节设置里的 Word 段落样式名套用到目标范围内的书目域。
 *
 * 样式名是章节级设置，与后端响应内容无关：这里只写域的段落样式，不改写
 * `Field.Data`、也不重渲染富文本内容，用户对内容做的局部格式因此得以保留。
 *
 * 返回实际发生样式写入的域数量。
 */
export function applyBibliographyStyles(
  range: Wps.Range,
  pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string },
): number {
  const fields = collectBibliographyFieldsInRange(range)
  let applied = 0
  for (const field of fields) {
    const data = readFieldData(field)
    const styleName = isBibliographyTitle(data)
      ? pref.bibliographyTitleStyle
      : isBibliographyEntry(data)
        ? pref.bibliographyEntryStyle
        : ""
    if (applyStyleToField(field, styleName, "paragraph")) {
      applied += 1
    }
  }
  return applied
}

/**
 * 设置对话框保存路径：把刚保存的章节样式名套用到当前章节的书目域。
 *
 * 只套用样式名，不改写 `Field.Data`、也不重渲染内容——样式名是章节级设置，
 * 与后端响应无关。
 *
 * 返回实际发生样式写入的域数量。
 */
export async function applyChapterBibliographyStyles(
  pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string },
): Promise<number> {
  const range = getUpdateRange()
  return withProgress(t("bibliography.styleProgressReason"), () =>
    withBatchUpdate("Banyan Styles", () => applyBibliographyStyles(range, pref)),
  )
}

function insertBibliographyLines(
  range: Wps.Range,
  lines: BibliographyLine[],
  firstIndex: number,
  lastIndex: number,
  pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string },
  separatorBeforeFirst: boolean,
  separatorAfterLast: boolean,
): boolean {
  if (firstIndex > lastIndex) return true
  const caret = range.Duplicate
  caret.Collapse(wps.Enum.wdCollapseStart)

  for (let i = firstIndex; i <= lastIndex; i += 1) {
    // 行间用段落分隔符隔开；插入点不在行首时首行也要补（VBA 的
    // separatorBeforeFirst），否则新块会并进上一行所在的段落。
    if ((i === firstIndex && separatorBeforeFirst) || i > firstIndex) {
      caret.InsertAfter("\r")
      caret.Collapse(wps.Enum.wdCollapseEnd)
    }

    const line = lines[i]
    const field = caret.Fields.Add(
      caret,
      wps.Enum.wdFieldAddin,
      bibliographyFieldCode(getBibliographyLineId(line)),
      false,
    )
    field.Data = JSON.stringify(line)
    const styleName = isBibliographyTitle(line)
      ? pref.bibliographyTitleStyle
      : pref.bibliographyEntryStyle
    renderStyledFieldWithData(
      field,
      (target) => applyStyleToField(target, styleName, "paragraph"),
      line,
      line.content,
    )
    if (isBibliographyEntry(line)) {
      addBookmarkToField(field, getBibliographyBookmarkName(line.id))
    }

    // 光标移到域之后（见 getWholeFieldRange）：`Field.Result` 在域内部，直接用它会把
    // 下一行卷进当前域的结果范围。
    const after = getWholeFieldRange(field)
    after.Collapse(wps.Enum.wdCollapseEnd)
    caret.SetRange(after.Start, after.End)
  }

  if (separatorAfterLast) {
    caret.InsertAfter("\r")
  }
  return true
}

export function deleteExistingBibliography(range: Wps.Range): boolean {
  const bibliographyFields = collectBibliographyFieldsInRange(range)

  if (bibliographyFields.length === 0) return false

  // Banyan 生成的书目是一段连续的 ADDIN 域，安全时一次宿主调用删掉整段；
  // 混有其它内容或边界不确定时退回较慢的逐域清理。
  const first = bibliographyFields[0]
  const last = bibliographyFields[bibliographyFields.length - 1]
  try {
    // 按“整个域”的边界删除（同 VBA 的 ReplaceBibliography）：只删到 `Result.End` 会
    // 残留域标记。
    const start = getWholeFieldRange(first).Start
    const end = getWholeFieldRange(last).End
    const block = wps.ActiveDocument.Range().Duplicate
    block.SetRange(start, end)
    let onlyBibliographyFields = block.Fields.Count === bibliographyFields.length
    for (let i = 1; onlyBibliographyFields && i <= block.Fields.Count; i += 1) {
      onlyBibliographyFields = fieldHasCodeKind(block.Fields.Item(i), "bibliography")
    }
    if (onlyBibliographyFields && end > start) {
      block.Delete()
      if (collectBibliographyFieldsInRange(range).length === 0) return true
    }
  }
  catch {
    // WPS 拿不到稳定的连续范围时，退回下面的逐域删除。
  }

  // 必须删除整个域：只删除 Result 会残留域代码，导致 collectBibliographyFieldsInRange
  // 仍能检测到这些域，旧的参考文献表无法被完全清除。
  // 同时倒序删除，避免前面域的删除使后面域的位置引用失效；
  // removeFieldSafely 保证单个域失败不影响其余域的清理。
  for (let i = bibliographyFields.length - 1; i >= 0; i--) {
    removeFieldSafely(bibliographyFields[i])
  }
  return true
}

export function insertBibliography(range: Wps.Range, lines: BibliographyLine[], pref: { bibliographyTitleStyle: string; bibliographyEntryStyle: string }) {
  const caret = range.Duplicate
  caret.Collapse(wps.Enum.wdCollapseEnd)
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    const field = caret.Fields.Add(
      caret,
      wps.Enum.wdFieldAddin,
      bibliographyFieldCode(getBibliographyLineId(line)),
      false,
    )
    if (isBibliographyTitle(line)) {
      field.Data = JSON.stringify(line)
      renderStyledFieldWithData(
        field,
        (field) => applyStyleToField(field, pref.bibliographyTitleStyle, "paragraph"),
        line,
        line.content,
      )
    }
    else if (isBibliographyEntry(line)) {
      field.Data = JSON.stringify(line)
      renderStyledFieldWithData(
        field,
        (field) => applyStyleToField(field, pref.bibliographyEntryStyle, "paragraph"),
        line,
        line.content,
      )
      addBookmarkToField(field, getBibliographyBookmarkName(line.id))
    }

    if (i < lines.length - 1) {
      // 从域之后插入分隔符（见 getWholeFieldRange）：`Field.Result` 在域内部，直接用它
      // 会把分隔符卷进结果范围。
      const after = getWholeFieldRange(field)
      after.Collapse(wps.Enum.wdCollapseEnd)
      after.InsertAfter("\r")
      after.Collapse(wps.Enum.wdCollapseEnd)
      caret.SetRange(after.Start, after.End)
    }
  }
}

function createPendingBibliographyField(
  range: Wps.Range,
  pref: { bibliographyTitleStyle: string }
): Wps.Field {
  const cursor = range.Duplicate
  cursor.Collapse(wps.Enum.wdCollapseEnd)
  const line = createPendingBibliographyLine()
  const field = cursor.Fields.Add(
    cursor,
    wps.Enum.wdFieldAddin,
    bibliographyFieldCode(line.id),
    false,
  )
  field.Data = JSON.stringify(line)
  renderStyledFieldWithData(
    field,
    (field) => applyStyleToField(field, pref.bibliographyTitleStyle, "paragraph"),
    line,
    line.content,
  )
  return field
}

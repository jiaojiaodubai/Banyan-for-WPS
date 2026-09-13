import {
  applyBibliographyStyles,
  collectBibliographyFieldsInRange,
  deleteExistingBibliography,
  getBibliographyFieldId,
  insertBibliography,
  updateBibliographyInPlace,
} from "../src/modules/bibliography"
import {
  getBibliographyBookmarkName,
  getBibliographyLineId,
  isBibliographyEntry,
} from "../src/utils/field"
import type { BibliographyEntry, BibliographyLine } from "../src/typings/style"
import { assert, type TestContext } from "./framework"
import {
  getDocumentEnd,
  makeBibliographyLines,
  resetTestDocument,
} from "./fixtures"

const MODULE = "bibliography"
const SIZES = [10, 50, 100]
const TEST_STYLES = {
  bibliographyTitleStyle: "Banyan Perf Bibliography Title",
  bibliographyEntryStyle: "Banyan Perf Bibliography Entry",
}

function bibliographyFieldIds(fields: Wps.Field[]): string[] {
  return fields.map((field) => getBibliographyFieldId(field))
}

/**
 * 插入路径把行尾段落标记放进域结果（`InsertParagraphAfter` 之后再渲染下一行），
 * 因此保留行的 `Result.Text` 会带上该行段落的段落标记；比较渲染文本时忽略它。
 */
function stripParagraphMarks(text: string): string {
  return text.replace(/\r/g, "")
}

/**
 * 镜像 `updateBibliographyField` 里的 `fieldStyleName`：段落样式名漂移是
 * 无变化刷新里唯一需要在门禁之后保留的宿主调用。
 */
function readStyleName(field: Wps.Field): string {
  try {
    const style = field.Result.Style as unknown
    if (typeof style === "string") return style
    if (style && typeof style === "object" && "NameLocal" in style) {
      return String((style as { NameLocal?: unknown }).NameLocal ?? "")
    }
  }
  catch {
    // 读取失败按样式不匹配处理，与生产代码一致。
  }
  return ""
}

/**
 * 追加一个域代码可控的 ADDIN 域，模拟被用户粘贴或改坏的一整行书目。
 * 按设计书目行自成一段，因此给域补上前后段落标记。
 */
function appendRawBibliographyField(code: string, data: BibliographyLine): Wps.Field {
  getDocumentEnd().InsertParagraphAfter()
  const range = getDocumentEnd()
  const field = range.Fields.Add(range, wps.Enum.wdFieldAddin, code, false)
  field.Data = JSON.stringify(data)
  field.Result.Text = data.content.text
  field.Result.InsertParagraphAfter()
  return field
}

export async function runBibliographyTests(context: TestContext): Promise<void> {
  await context.test(MODULE, "insert, collect and delete", () => {
    try {
      resetTestDocument()
      const lines = makeBibliographyLines(2)
      insertBibliography(getDocumentEnd(), lines, TEST_STYLES)
      assert.equal(collectBibliographyFieldsInRange(wps.ActiveDocument.Content).length, lines.length)
      assert.ok(wps.ActiveDocument.Bookmarks.Exists(getBibliographyBookmarkName("bib-1")))
      assert.ok(wps.ActiveDocument.Bookmarks.Exists(getBibliographyBookmarkName("bib-2")))
      assert.equal(deleteExistingBibliography(wps.ActiveDocument.Content), true)
    }
    finally {
      resetTestDocument()
    }
  })

  for (const size of SIZES) {
    await context.test(MODULE, `rebuild size=${size}`, async () => {
      try {
        resetTestDocument()
        const lines = makeBibliographyLines(size)
        await context.measure(MODULE, "insert bibliography", () => {
          insertBibliography(getDocumentEnd(), lines, TEST_STYLES)
        }, { size, mode: "rebuild" })

        let fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
        await context.measure(MODULE, "collect bibliography", () => {
          fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
        }, { size, mode: "rebuild" })
        assert.equal(fields.length, lines.length)

        await context.measure(MODULE, "delete bibliography", () => {
          assert.equal(deleteExistingBibliography(wps.ActiveDocument.Content), true)
        }, { size, mode: "rebuild" })
      }
      finally {
        resetTestDocument()
      }
    })
  }

  await context.test(MODULE, "incremental diff retains existing fields", async () => {
    try {
      resetTestDocument()
      const initial = makeBibliographyLines(4)
      insertBibliography(getDocumentEnd(), initial, TEST_STYLES)
      let fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(fields.length, initial.length)

      // 模拟用户手工格式/内容：在已有首条题录前插入新题录后，它必须仍然存在。
      fields[1].Result.Text = "LOCAL A"
      const firstEntry = initial.find(isBibliographyEntry)
      assert.ok(firstEntry)
      const insertedEntry: BibliographyEntry = { ...firstEntry, id: "bib-extra" }
      const updated: BibliographyLine[] = [initial[0], insertedEntry, ...initial.slice(1)]

      await context.measure(MODULE, "incremental bibliography diff", () => {
        assert.equal(updateBibliographyInPlace(wps.ActiveDocument.Content, updated, TEST_STYLES), true)
      }, { size: updated.length, mode: "incremental" })

      fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.deepEqual(bibliographyFieldIds(fields), ["bib-title", "bib-extra", "bib-1", "bib-2", "bib-3", "bib-4"])
      assert.equal(fields[2].Result.Text, "LOCAL A")
      assert.ok(wps.ActiveDocument.Bookmarks.Exists(getBibliographyBookmarkName("bib-extra")))
    }
    finally {
      resetTestDocument()
    }
  })

  // VBA 一致性：宿主回读数据与响应逐字节一致时，整个域都不触碰，应用返回 false。
  await context.test(MODULE, "incremental refresh keeps unchanged bibliography fields", async () => {
    resetTestDocument()
    const lines = makeBibliographyLines(10)
    insertBibliography(getDocumentEnd(), lines, TEST_STYLES)
    const dataBefore = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      .map((field) => field.Data)
    await context.measure(MODULE, "bibliography unchanged refresh", () => {
      assert.equal(updateBibliographyInPlace(wps.ActiveDocument.Content, lines, TEST_STYLES), false)
    }, { size: lines.length, mode: "no-op" })
    const fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
    assert.equal(fields.length, lines.length)
    assert.deepEqual(fields.map((field) => field.Data), dataBefore)
    resetTestDocument()
  })

  await context.test(MODULE, "bibliography deletion strategy comparison", async () => {
    const size = 50
    const prepare = () => {
      resetTestDocument()
      insertBibliography(getDocumentEnd(), makeBibliographyLines(size), TEST_STYLES)
      const fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(fields.length, size + 1)
    }
    await context.measureMedian(MODULE, "delete bibliography strategy", () => {
      const fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      const first = fields[0]
      const last = fields[fields.length - 1]
      const block = wps.ActiveDocument.Range().Duplicate
      block.SetRange(
        Math.max(wps.ActiveDocument.Content.Start, first.Code.Start - 1),
        Math.min(wps.ActiveDocument.Content.End, last.Result.End + 1),
      )
      block.Delete()
      assert.equal(collectBibliographyFieldsInRange(wps.ActiveDocument.Content).length, 0)
    }, { size, mode: "contiguous-range" }, 3, prepare)

    await context.measureMedian(MODULE, "delete bibliography strategy", () => {
      assert.equal(deleteExistingBibliography(wps.ActiveDocument.Content), true)
      assert.equal(collectBibliographyFieldsInRange(wps.ActiveDocument.Content).length, 0)
    }, { size, mode: "field-by-field" }, 3, prepare)
  })

  await context.test(MODULE, "bibliography deletion preserves surrounding content", () => {
    resetTestDocument()
    const document = wps.ActiveDocument
    document.Content.Text = "BEFORE "
    const insertRange = document.Content.Duplicate
    insertRange.Collapse(wps.Enum.wdCollapseEnd)
    insertBibliography(insertRange, makeBibliographyLines(3), TEST_STYLES)
    const after = document.Content.Duplicate
    after.Collapse(wps.Enum.wdCollapseEnd)
    after.InsertAfter(" AFTER")
    assert.equal(deleteExistingBibliography(document.Content), true)
    const text = document.Content.Text
    assert.ok(text.includes("BEFORE"))
    assert.ok(text.includes("AFTER"))
    assert.equal(collectBibliographyFieldsInRange(document.Content).length, 0)
    for (let i = 1; i <= 3; i += 1) {
      assert.equal(document.Bookmarks.Exists(getBibliographyBookmarkName(`bib-${i}`)), false)
    }
    resetTestDocument()
  })

  await context.test(MODULE, "bibliography deletion handles multiple sizes and empty ranges", () => {
    resetTestDocument()
    assert.equal(deleteExistingBibliography(wps.ActiveDocument.Content), false)
    for (const size of [1, 10, 100]) {
      resetTestDocument()
      insertBibliography(getDocumentEnd(), makeBibliographyLines(size), TEST_STYLES)
      assert.equal(deleteExistingBibliography(wps.ActiveDocument.Content), true)
      assert.equal(collectBibliographyFieldsInRange(wps.ActiveDocument.Content).length, 0)
      assert.equal(wps.ActiveDocument.Bookmarks.Exists(getBibliographyBookmarkName(`bib-${size}`)), false)
    }
    resetTestDocument()
  })

  // A1 结论（51 域实测）：宿主回读的 Field.Data 与序列化响应逐字节一致
  // （skipped=51/51），因此 updateBibliographyInPlace 现在按 VBA 语义在逐域
  // 门禁处短路。这里保留「门禁后的实现」与「无条件回写」的对照，并单独测量
  // 样式套用的成本——样式套用已从内容渲染中解耦。
  await context.test(MODULE, "no-op refresh data write cost", async () => {
    const size = 50
    try {
      resetTestDocument()
      const lines = makeBibliographyLines(size)
      insertBibliography(getDocumentEnd(), lines, TEST_STYLES)
      const fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(fields.length, lines.length)

      let changed: unknown = null
      await context.measureMedian(MODULE, "no-op in-place diff", () => {
        changed = updateBibliographyInPlace(wps.ActiveDocument.Content, lines, TEST_STYLES)
      }, { size: fields.length, mode: "gated" }, 3)
      assert.equal(changed, false)

      await context.measureMedian(MODULE, "no-op field data rewrite", () => {
        for (let i = 0; i < fields.length; i += 1) {
          fields[i].Data = JSON.stringify(lines[i])
        }
      }, { size: fields.length, mode: "writeOnly" }, 3)

      let applied = -1
      await context.measureMedian(MODULE, "no-op style apply", () => {
        applied = applyBibliographyStyles(wps.ActiveDocument.Content, TEST_STYLES)
      }, { size: fields.length, mode: "styleApply" }, 3)

      // 诊断：枚举一次 ActiveDocument.Styles 的成本。`applyStyleToField` 只在样式
      // 名不一致时才会走到样式查找，这里量化那次查找的下限。
      await context.measureMedian(MODULE, "styles enumeration", () => {
        const styles = wps.ActiveDocument.Styles
        let found = 0
        for (let i = 1; i <= styles.Count; i += 1) {
          if (styles.Item(i).NameLocal) found += 1
        }
        assert.ok(found > 0)
      }, { size: 1, mode: "diagnostic" }, 3)

      // WPS 专属门禁（VBA 侧按名取样式只能枚举 Styles）：量化「按名直取」相对
      // 「枚举查找」的收益。
      let byNameWorks = false
      await context.measureMedian(MODULE, "style lookup by name", () => {
        try {
          const styles = wps.ActiveDocument.Styles
          const style = (styles.Item as unknown as (key: string) => Wps.Style)(
            TEST_STYLES.bibliographyEntryStyle,
          )
          byNameWorks = !!style && style.NameLocal === TEST_STYLES.bibliographyEntryStyle
        }
        catch {
          byNameWorks = false
        }
      }, { size: 1, mode: "by-name" }, 3)

      let byEnumerationWorks = false
      await context.measureMedian(MODULE, "style lookup by enumeration", () => {
        const styles = wps.ActiveDocument.Styles
        const count = styles.Count
        byEnumerationWorks = false
        for (let i = 1; i <= count; i += 1) {
          if (styles.Item(i).NameLocal === TEST_STYLES.bibliographyEntryStyle) {
            byEnumerationWorks = true
            break
          }
        }
      }, { size: 1, mode: "by-enumeration" }, 3)

      // 为什么不做样式对象缓存：`wps.ActiveDocument` 每次访问都返回新的包装对象，
      // 基于文档身份的缓存必然每次失效（因此该缓存已移除）。
      let identityStable = false
      await context.measureMedian(MODULE, "document identity check", () => {
        identityStable = wps.ActiveDocument === wps.ActiveDocument
      }, { size: 1, mode: "diagnostic" }, 3)

      context.note(`bibliography no-op diff changed=${String(changed)}; style apply writes=${applied}/${fields.length} fields; lookup by name=${String(byNameWorks)} vs by enumeration=${String(byEnumerationWorks)}; ActiveDocument identity stable=${String(identityStable)}`)
    }
    finally {
      resetTestDocument()
    }
  })

  // 样式名是章节级设置：改名后只需要把新样式名套到目标 range，不涉及内容重渲染。
  await context.test(MODULE, "style name change applies without re-render", () => {
    try {
      resetTestDocument()
      insertBibliography(getDocumentEnd(), makeBibliographyLines(2), TEST_STYLES)
      let fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      const dataBefore = fields.map((field) => field.Data)
      const textBefore = fields.map((field) => field.Result.Text)

      const renamed = {
        bibliographyTitleStyle: "Banyan Renamed Title",
        bibliographyEntryStyle: "Banyan Renamed Entry",
      }
      assert.equal(applyBibliographyStyles(wps.ActiveDocument.Content, renamed), fields.length)

      fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(readStyleName(fields[0]), renamed.bibliographyTitleStyle)
      assert.equal(readStyleName(fields[1]), renamed.bibliographyEntryStyle)
      assert.equal(readStyleName(fields[2]), renamed.bibliographyEntryStyle)
      // 只写样式名：Field.Data 与渲染结果都不变。
      assert.deepEqual(fields.map((field) => field.Data), dataBefore)
      assert.deepEqual(fields.map((field) => field.Result.Text), textBefore)
      // 已套用后再次调用不应产生写入。
      assert.equal(applyBibliographyStyles(wps.ActiveDocument.Content, renamed), 0)
    }
    finally {
      resetTestDocument()
    }
  })

  // 契约：非法行（id 缺失或已被占用）在收集阶段就整行删除，剩下的行继续参与
  // ID 匹配；因此收集结果直接就是合法行，diff 不再需要为此退回全量重建。
  await context.test(MODULE, "illegal document lines are dropped", async () => {
    try {
      resetTestDocument()
      const lines = makeBibliographyLines(2)
      insertBibliography(getDocumentEnd(), lines, TEST_STYLES)

      appendRawBibliographyField("BANYAN_BIBLIOGRAPHY", lines[0])
      appendRawBibliographyField(
        `BANYAN_BIBLIOGRAPHY ${getBibliographyLineId(lines[2])}`,
        lines[2],
      )

      // 收集阶段就把两行非法行删掉（含行尾段落标记）。
      const fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.deepEqual(bibliographyFieldIds(fields), lines.map((line) => getBibliographyLineId(line)))
      assert.equal(stripParagraphMarks(fields[1].Result.Text), lines[1].content.text)

      // 剩下的行与响应完全一致，无需任何写入。
      assert.equal(updateBibliographyInPlace(wps.ActiveDocument.Content, lines, TEST_STYLES), false)
    }
    finally {
      resetTestDocument()
    }
  })

  // 契约规则（与 VBA 一致）：没有数据（旧前端残留的空数据域）或数据无法解析的
  // 行一律保持原样——刷新不去猜它本来应该是什么。
  await context.test(MODULE, "unreadable field data is left alone", async () => {
    try {
      resetTestDocument()
      const lines = makeBibliographyLines(2)
      insertBibliography(getDocumentEnd(), lines, TEST_STYLES)

      const fields = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(fields.length, lines.length)
      fields[1].Data = "{ not json"

      assert.equal(updateBibliographyInPlace(wps.ActiveDocument.Content, lines, TEST_STYLES), false)
      let updated = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(updated[1].Data, "{ not json")

      // 完全没有数据的域同样不被触碰（域代码里带着 id，因此仍然参与 ID 匹配）。
      const emptyLine = { ...(lines[2] as BibliographyEntry), id: "bib-empty" }
      const emptyField = appendRawBibliographyField(
        `BANYAN_BIBLIOGRAPHY ${emptyLine.id}`,
        emptyLine,
      )
      emptyField.Data = ""

      assert.equal(
        updateBibliographyInPlace(wps.ActiveDocument.Content, [...lines, emptyLine], TEST_STYLES),
        false,
      )
      updated = collectBibliographyFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(updated[3].Data, "")
    }
    finally {
      resetTestDocument()
    }
  })

  resetTestDocument()
}

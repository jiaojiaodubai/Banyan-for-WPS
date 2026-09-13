import {
  applyIntextCitationStyle,
  collectIntextCitationFieldsInRange,
  getBanyanFieldCode,
  readFieldData,
  renderStyledField,
} from "../src/utils/field"
import { assert, type TestContext } from "./framework"
import {
  insertRawIntextField,
  insertSeparator,
  makeIntextCitation,
  resetTestDocument,
} from "./fixtures"

const MODULE = "refresh"
const SIZES = [10, 50, 100]

export async function runRefreshTests(context: TestContext): Promise<void> {
  for (const size of SIZES) {
    for (const screenUpdating of [true, false]) {
      const mode = screenUpdating ? "screenOn" : "screenOff"
      await context.test(MODULE, `refresh-like path size=${size} ${mode}`, async () => {
        const previousScreenUpdating = Application.ScreenUpdating
        try {
          resetTestDocument()
          // WPS can drop fields while a large fixture is being created with
          // repainting enabled. Setup is excluded from the measured stages.
          Application.ScreenUpdating = false
          for (let i = 0; i < size; i += 1) {
            insertRawIntextField(makeIntextCitation(`refresh-${i}`, "plain"))
            if (i < size - 1) insertSeparator()
          }

          let pairs = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
          assert.equal(pairs.length, size)
          Application.ScreenUpdating = screenUpdating
          await context.measure(MODULE, "collect+parse", () => {
            pairs = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
          }, { size, mode })
          assert.equal(pairs.length, size)

          await context.measure(MODULE, "page lookup", () => {
            for (const pair of pairs) {
              pair.field.Result.Information(wps.Enum.wdActiveEndPageNumber)
            }
          }, { size, mode })

          const updates = pairs.map((_, index) => makeIntextCitation(`refresh-${index}`, "rich"))
          await context.measure(MODULE, "JSON stringify", () => {
            JSON.stringify(updates)
          }, { size, mode: "memory" })
          const payload = JSON.stringify(updates)
          await context.measure(MODULE, "JSON parse", () => {
            JSON.parse(payload)
          }, { size, mode: "memory" })

          await context.measure(MODULE, "linear response matching", () => {
            for (let i = size - 1; i >= 0; i -= 1) {
              const found = updates.find(item => item.id === `refresh-${i}`)
              assert.ok(found)
            }
          }, { size, mode: "worstOrder" })

          await context.measure(MODULE, "indexed response matching", () => {
            const index = new Map(updates.map(item => [item.id, item]))
            for (let i = size - 1; i >= 0; i -= 1) {
              assert.ok(index.get(`refresh-${i}`))
            }
          }, { size, mode: "build+lookup" })

          await context.measure(MODULE, "write field data", () => {
            for (let i = 0; i < pairs.length; i += 1) {
              pairs[i].field.Data = JSON.stringify(updates[i])
            }
          }, { size, mode })

          await context.measure(MODULE, "render rich", () => {
            for (let i = 0; i < pairs.length; i += 1) {
              renderStyledField(pairs[i].field, applyIntextCitationStyle, updates[i].content)
            }
          }, { size, mode })
        }
        finally {
          Application.ScreenUpdating = previousScreenUpdating
          resetTestDocument()
        }
      })
    }
  }

  // WPS may spend substantial time repaginating and checking spelling while
  // fields are rewritten. Measure the same render workload with each option
  // disabled, restoring every host setting even when a JSAPI call fails.
  await context.test(MODULE, "render with host options disabled", async () => {
    // Keep this diagnostic probe small: it protects against blindly adding a
    // global host-option toggle without dominating the whole test run.
    const size = 10
    let options: Wps.Options
    try {
      options = Application.Options
      if (!options || typeof options.Pagination !== "boolean") {
        context.note("Application.Options pagination is unavailable; skipped host-option benchmark.")
        return
      }
    }
    catch {
      context.note("Application.Options is unavailable; skipped host-option benchmark.")
      return
    }
    const originalScreenUpdating = Application.ScreenUpdating
    const original = {
      pagination: options.Pagination,
      spelling: options.CheckSpellingAsYouType,
      grammar: options.CheckGrammarAsYouType,
    }
    try {
      resetTestDocument()
      const modes: Array<{ name: string; pagination: boolean; spelling: boolean; grammar: boolean }> = [
        { name: "baseline", pagination: original.pagination, spelling: original.spelling, grammar: original.grammar },
        { name: "no-pagination", pagination: false, spelling: original.spelling, grammar: original.grammar },
        { name: "no-spelling", pagination: original.pagination, spelling: false, grammar: original.grammar },
        { name: "no-grammar", pagination: original.pagination, spelling: original.spelling, grammar: false },
        { name: "all-disabled", pagination: false, spelling: false, grammar: false },
      ]
      for (const mode of modes) {
        resetTestDocument()
        Application.ScreenUpdating = false
        for (let i = 0; i < size; i += 1) {
          insertRawIntextField(makeIntextCitation(`options-${i}`, "rich"))
          if (i < size - 1) insertSeparator()
        }
        const pairs = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
        assert.equal(pairs.length, size)
        const updates = pairs.map((_, index) => makeIntextCitation(`options-${index}`, "rich"))
        options.Pagination = mode.pagination
        options.CheckSpellingAsYouType = mode.spelling
        options.CheckGrammarAsYouType = mode.grammar
        await context.measure(MODULE, "render host option mode", () => {
          for (let i = 0; i < pairs.length; i += 1) {
            renderStyledField(pairs[i].field, applyIntextCitationStyle, updates[i].content)
          }
        }, { size, mode: mode.name })
      }
    }
    finally {
      options.Pagination = original.pagination
      options.CheckSpellingAsYouType = original.spelling
      options.CheckGrammarAsYouType = original.grammar
      Application.ScreenUpdating = originalScreenUpdating
      resetTestDocument()
    }
  })
  // B1：复制粘贴引注会产生重复 id；收集阶段必须就地重键——把新 id 写进域代码，
  // 已存数据保持不动（响应会带着新 id 回来，在下一次写入时整体替换旧文本）。
  await context.test(MODULE, "duplicate citation ids are re-keyed", async () => {
    try {
      resetTestDocument()
      insertRawIntextField(makeIntextCitation("duplicate-id", "plain"))
      insertSeparator()
      const copy = insertRawIntextField(makeIntextCitation("duplicate-id", "plain"))
      const storedBefore = copy.Data

      const pairs = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(pairs.length, 2)
      assert.equal(pairs[0].id, "duplicate-id")
      assert.ok(pairs[1].id !== "duplicate-id")

      // 新 id 必须写进了域代码，否则下一次刷新又会撞上重复。
      assert.equal(getBanyanFieldCode(pairs[1].field)?.id, pairs[1].id)
      // 已存数据保持原样（仍带旧 id），等响应回来时整体替换。
      assert.equal(pairs[1].field.Data, storedBefore)

      // 第二次收集必须稳定：不再重键。
      const second = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
      assert.equal(second.length, 2)
      assert.equal(second[0].id, "duplicate-id")
      assert.equal(second[1].id, pairs[1].id)
    }
    finally {
      resetTestDocument()
    }
  })

  // 契约：代码里的 id 丢了就不是一个引注，收集阶段直接删除该域。
  await context.test(MODULE, "citations without an id are dropped", async () => {
    try {
      resetTestDocument()
      insertRawIntextField(makeIntextCitation("keep-id", "plain"))
      insertSeparator()
      const end = wps.ActiveDocument.Content.Duplicate
      end.Collapse(wps.Enum.wdCollapseEnd)
      end.Fields.Add(end, wps.Enum.wdFieldAddin, "BANYAN_CITATION ", false)

      assert.equal(collectIntextCitationFieldsInRange(wps.ActiveDocument.Content).length, 1)
      assert.equal(wps.ActiveDocument.Fields.Count, 1)
    }
    finally {
      resetTestDocument()
    }
  })

  // 懒解析收益：收集阶段只解析域代码（不读 Field.Data），更新阶段用请求前缓存的
  // 文本跳过无变化的引注。对照指标是「为每个引注读一次 Field.Data 并 JSON 解析」。
  for (const size of SIZES) {
    await context.test(MODULE, `lazy parsing stages size=${size}`, async () => {
      const previousScreenUpdating = Application.ScreenUpdating
      try {
        resetTestDocument()
        // 大夹具在开启重绘时可能掉域，构造阶段不计入测量。
        Application.ScreenUpdating = false
        for (let i = 0; i < size; i += 1) {
          insertRawIntextField(makeIntextCitation(`lazy-${i}`, "rich"))
          if (i < size - 1) insertSeparator()
        }
        Application.ScreenUpdating = previousScreenUpdating

        let pairs = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
        await context.measure(MODULE, "classify by field code", () => {
          pairs = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
        }, { size, mode: "code-only" })
        assert.equal(pairs.length, size)

        // 公平对照：同样走一遍域集合，只是分类方式换回「读 Field.Data + 类型守卫」。
        await context.measure(MODULE, "classify by field data", () => {
          const range = wps.ActiveDocument.Content
          let found = 0
          for (let i = 1; i <= range.Fields.Count; i += 1) {
            const field = range.Fields.Item(i)
            if (field.Type !== wps.Enum.wdFieldAddin) continue
            const data = readFieldData(field)
            if (data && (data.type === "intext-citation" || data.type === "note-citation")) found += 1
          }
          assert.equal(found, size)
        }, { size, mode: "data-read" })

        // 公平拆分「取值」阶段：两边的脚手架一致，只差读 Code.Text 还是 Data。
        await context.measure(MODULE, "read field code text", () => {
          const range = wps.ActiveDocument.Content
          const fields = range.Fields
          const count = fields.Count
          let chars = 0
          for (let i = 1; i <= count; i += 1) {
            const field = fields.Item(i)
            if (field.Type !== wps.Enum.wdFieldAddin) continue
            chars += String(field.Code.Text ?? "").length
          }
          assert.ok(chars > 0)
        }, { size, mode: "code-text" })

        await context.measure(MODULE, "read field data text", () => {
          const range = wps.ActiveDocument.Content
          const fields = range.Fields
          const count = fields.Count
          let chars = 0
          for (let i = 1; i <= count; i += 1) {
            const field = fields.Item(i)
            if (field.Type !== wps.Enum.wdFieldAddin) continue
            chars += String(field.Data ?? "").length
          }
          assert.ok(chars > 0)
        }, { size, mode: "data-text" })

        await context.measure(MODULE, "read cached field data", () => {
          for (const pair of pairs) {
            readFieldData(pair.field)
          }
        }, { size, mode: "data-only" })

        const cached = pairs.map((pair) => JSON.stringify(readFieldData(pair.field)))
        const response = pairs.map((pair) => JSON.stringify(readFieldData(pair.field)))
        await context.measure(MODULE, "skip unchanged by cached text", () => {
          for (let i = 0; i < cached.length; i += 1) {
            assert.equal(cached[i] === response[i], true)
          }
        }, { size, mode: "text-compare" })
      }
      finally {
        Application.ScreenUpdating = previousScreenUpdating
        resetTestDocument()
      }
    })
  }
  resetTestDocument()
}

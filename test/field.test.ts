import {
  applyIntextCitationStyle,
  addBookmarkToField,
  applyRichTextLinksToRange,
  applyRichTextStylesToRange,
  collectIntextCitationFieldsInRange,
  createIntextCitationAtRange,
  createNoteCitationAtRange,
  createPlaceholderIntextCitationData,
  getBibliographyBookmarkName,
  isIntextCitation,
  isNoteCitation,
  isRichText,
  normalizeBookmarkName,
  readFieldData,
  richTextEquals,
  renderStyledField,
} from "../src/utils/field"
import { assert, type TestContext } from "./framework"
import {
  getDocumentEnd,
  insertRawIntextField,
  insertSeparator,
  makeIntextCitation,
  makeNoteCitation,
  resetTestDocument,
  type RichTextVariant,
} from "./fixtures"

const MODULE = "field"
const REPETITIONS = 4

export async function runFieldTests(context: TestContext): Promise<void> {
  await context.test(MODULE, "runtime validators", () => {
    const intext = makeIntextCitation("validator")
    const note = makeNoteCitation("validator-note")
    assert.ok(isRichText(intext.content))
    assert.ok(isIntextCitation(intext))
    assert.ok(isNoteCitation(note))
    assert.equal(isRichText({ text: "invalid", marks: [{ type: "bold", start: 0, end: 99, value: true }] }), false)
  })

  await context.test(MODULE, "rich-text structural comparison", () => {
    const current = makeIntextCitation("compare", "rich").content
    const equivalent = { text: current.text, marks: current.marks.map(mark => ({ ...mark })) }
    assert.equal(richTextEquals(current, equivalent), true)
    assert.equal(richTextEquals(current, { ...equivalent, text: `${equivalent.text}!` }), false)
    assert.equal(richTextEquals(current, { ...equivalent, marks: equivalent.marks.slice(0, -1) }), false)
  })

  await context.test(MODULE, "Word bookmark name normalization", () => {
    assert.equal(normalizeBookmarkName("Banyan_Entry_JZGANGMD"), "Banyan_Entry_JZGANGMD")
    assert.equal(normalizeBookmarkName("supp fig 1.2"), "supp_fig_1_2")
    assert.equal(normalizeBookmarkName("1abc"), "B1abc")
    assert.equal(normalizeBookmarkName("_hidden"), "B_hidden")
    assert.equal(normalizeBookmarkName(""), "")
    assert.equal(normalizeBookmarkName("  padded  "), "padded")
    assert.equal(normalizeBookmarkName("a".repeat(80)).length, 40)
    assert.equal(getBibliographyBookmarkName("JZGANGMD"), "Banyan_Entry_JZGANGMD")
    assert.equal(getBibliographyBookmarkName("a".repeat(60)).length, 40)
    const unicodeName = normalizeBookmarkName("中文 entry")
    assert.ok(/^[A-Za-z][A-Za-z0-9_]*$/.test(unicodeName))
    assert.ok(unicodeName.length <= 40)
    const bibliographyName = getBibliographyBookmarkName("a b.c-d")
    assert.ok(/^[A-Za-z][A-Za-z0-9_]*$/.test(bibliographyName))
    assert.ok(bibliographyName.length <= 40)
    assert.equal(normalizeBookmarkName("Banyan_Entry_item_1"), "Banyan_Entry_item_1")
    assert.equal(
      getBibliographyBookmarkName("item-1"),
      normalizeBookmarkName("Banyan_Entry_item-1"),
    )

    resetTestDocument()
    const field = insertRawIntextField(makeIntextCitation("bookmark-gate", "plain"))
    addBookmarkToField(field, "Banyan Entry bad.name")
    assert.ok(wps.ActiveDocument.Bookmarks.Exists("Banyan_Entry_bad_name"))
  })

  await context.test(MODULE, "in-text citation create and collect", () => {
    resetTestDocument()
    const data = makeIntextCitation("create-1")
    const field = createIntextCitationAtRange(getDocumentEnd(), data)
    const readBack = readFieldData(field)
    assert.ok(isIntextCitation(readBack))
    assert.equal(readBack.id, data.id)
    assert.equal(field.Result.Text, data.content.text)
    const collected = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
    assert.equal(collected.length, 1)
  })

  await context.test(MODULE, "note citation create", () => {
    resetTestDocument()
    const data = makeNoteCitation("note-1")
    const created = createNoteCitationAtRange(getDocumentEnd(), data)
    const readBack = readFieldData(created.field)
    assert.ok(isNoteCitation(readBack))
    assert.equal(readBack.id, data.id)
    assert.equal(wps.ActiveDocument.Footnotes.Count, 1)
  })

  await context.test(MODULE, "placeholder insertion cost", async () => {
    await context.measure(MODULE, "placeholder create+write+render", () => {
      for (let i = 0; i < REPETITIONS; i += 1) {
        resetTestDocument()
        const placeholder = createPlaceholderIntextCitationData(`placeholder-${i}`)
        createIntextCitationAtRange(getDocumentEnd(), placeholder)
      }
    }, { iterations: REPETITIONS })
  })

  await context.test(MODULE, "field creation stages", async () => {
    const data = makeIntextCitation("field-create", "plain")
    await context.measure(MODULE, "field add", () => {
      resetTestDocument()
      for (let i = 0; i < REPETITIONS; i += 1) {
        const range = getDocumentEnd()
        range.Fields.Add(range, wps.Enum.wdFieldAddin, `BANYAN_CITATION ${data.id}-${i}`, false)
      }
    }, { iterations: REPETITIONS, mode: "host-create" })
    resetTestDocument()
    const dataFields: Wps.Field[] = []
    for (let i = 0; i < REPETITIONS; i += 1) {
      const range = getDocumentEnd()
      dataFields.push(range.Fields.Add(
        range,
        wps.Enum.wdFieldAddin,
        `BANYAN_CITATION ${data.id}-data-${i}`,
        false,
      ))
    }
    await context.measure(MODULE, "field data write", () => {
      for (const field of dataFields) {
        field.Data = JSON.stringify(data)
      }
    }, { iterations: REPETITIONS, mode: "host-data" })
  })

  const variants: RichTextVariant[] = ["plain", "format", "link", "rich"]
  for (const variant of variants) {
    await context.test(MODULE, `render ${variant}`, async () => {
      await context.measure(MODULE, `render ${variant}`, () => {
        for (let i = 0; i < REPETITIONS; i += 1) {
          resetTestDocument()
          const data = makeIntextCitation(`${variant}-${i}`, variant)
          const field = insertRawIntextField(data)
          assert.ok(renderStyledField(field, applyIntextCitationStyle, data.content, "character"))
        }
      }, { iterations: REPETITIONS })
    })
  }

  await context.test(MODULE, "explicit versus implicit content", async () => {
    const data = makeIntextCitation("content-argument")
    await context.measure(MODULE, "render explicit content", () => {
      for (let i = 0; i < REPETITIONS; i += 1) {
        resetTestDocument()
        const field = insertRawIntextField(data)
        renderStyledField(field, applyIntextCitationStyle, data.content, "character")
      }
    }, { iterations: REPETITIONS })
    await context.measure(MODULE, "render implicit content", () => {
      for (let i = 0; i < REPETITIONS; i += 1) {
        resetTestDocument()
        const field = insertRawIntextField(data)
        renderStyledField(field, applyIntextCitationStyle, undefined, "character")
      }
    }, { iterations: REPETITIONS })
  })

  await context.test(MODULE, "citation style application cost", async () => {
    resetTestDocument()
    const field = insertRawIntextField(makeIntextCitation("style-lookup", "plain"))
    const iterations = REPETITIONS * 5
    await context.measure(MODULE, "apply citation style", () => {
      for (let i = 0; i < iterations; i += 1) {
        applyIntextCitationStyle(field)
      }
    }, { iterations, mode: "current-style-scan" })
  })

  await context.test(MODULE, "style assignment optimization comparison", async () => {
    const data = makeIntextCitation("style-compare", "plain")
    const iterations = 4
    const measureMode = async (mode: "always-assign" | "skip-if-matching") => {
      resetTestDocument()
      const fields: Wps.Field[] = []
      for (let i = 0; i < iterations; i += 1) {
        const field = insertRawIntextField(data)
        applyIntextCitationStyle(field) // warm style and establish matching state
        fields.push(field)
        if (i < iterations - 1) insertSeparator()
      }
      let expectedStyle = ""
      try {
        const value = fields[0].Result.Style as unknown
        expectedStyle = typeof value === "string"
          ? value
          : String((value as { NameLocal?: unknown })?.NameLocal ?? value ?? "")
      }
      catch {
        expectedStyle = ""
      }
      await context.measureMedian(MODULE, "style assignment strategy", () => {
        for (const field of fields) {
          if (mode === "always-assign") {
            applyIntextCitationStyle(field)
          }
          else {
            let currentStyle = ""
            try {
              const value = field.Result.Style as unknown
              currentStyle = typeof value === "string"
                ? value
                : String((value as { NameLocal?: unknown })?.NameLocal ?? value ?? "")
            }
            catch {
              // 宿主不支持读样式时，为保证正确性直接赋值。
              applyIntextCitationStyle(field)
              continue
            }
            if (!expectedStyle || currentStyle !== expectedStyle) applyIntextCitationStyle(field)
          }
        }
      }, { iterations, mode }, 3)
    }
    await measureMode("always-assign")
    await measureMode("skip-if-matching")
  })

  // 把昂贵的渲染路径拆成逐个宿主调用。上面的汇总渲染指标用于回归跟踪，
  // 这些细分阶段则用于看出是哪个 JSAPI 操作导致总时长变化。
  await context.test(MODULE, "render stage breakdown", async () => {
    const data = makeIntextCitation("stage-breakdown", "rich")
    const measureStage = async (
      name: string,
      callback: (field: Wps.Field) => void,
      prepare?: (field: Wps.Field) => void,
    ) => {
      resetTestDocument()
      const fields: Wps.Field[] = []
      for (let i = 0; i < REPETITIONS; i += 1) {
        const field = insertRawIntextField(data)
        prepare?.(field)
        fields.push(field)
        if (i < REPETITIONS - 1) insertSeparator()
      }
      // 先热一次样式应用，使该阶段指标反映的是赋值成本，而不是首次的样式查找/
      // 创建开销。
      if (name === "apply field style") callback(fields[0])
      await context.measure(MODULE, name, () => {
        for (const field of fields) callback(field)
      }, { iterations: REPETITIONS, mode: "rich" })
    }

    await measureStage("write result text", (field) => {
      field.Result.Text = data.content.text
    })
    await measureStage("apply field style", (field) => {
      applyIntextCitationStyle(field)
    })
    await measureStage("apply rich text styles", (field) => {
      applyRichTextStylesToRange(field.Result, data.content)
    })
    await measureStage("apply rich text links", (field) => {
      const linkContent = {
        text: data.content.text,
        marks: [{
          type: "link" as const,
          start: 0,
          end: data.content.text.length,
          value: "https://example.com/banyan-performance",
        }],
      }
      applyRichTextLinksToRange(field.Result, linkContent)
    })
  })

  resetTestDocument()
}

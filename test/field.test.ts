import {
  applyIntextCitationStyle,
  collectIntextCitationFieldsInRange,
  createIntextCitationAtRange,
  createNoteCitationAtRange,
  createPlaceholderIntextCitationData,
  isIntextCitation,
  isNoteCitation,
  isRichText,
  readFieldData,
  renderStyledField,
} from "../src/utils/field"
import { assert, type TestContext } from "./framework"
import {
  getDocumentEnd,
  insertRawIntextField,
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

  const variants: RichTextVariant[] = ["plain", "format", "link", "rich"]
  for (const variant of variants) {
    await context.test(MODULE, `render ${variant}`, async () => {
      await context.measure(MODULE, `render ${variant}`, () => {
        for (let i = 0; i < REPETITIONS; i += 1) {
          resetTestDocument()
          const data = makeIntextCitation(`${variant}-${i}`, variant)
          const field = insertRawIntextField(data)
          assert.ok(renderStyledField(field, applyIntextCitationStyle, data.content))
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
        renderStyledField(field, applyIntextCitationStyle, data.content)
      }
    }, { iterations: REPETITIONS })
    await context.measure(MODULE, "render implicit content", () => {
      for (let i = 0; i < REPETITIONS; i += 1) {
        resetTestDocument()
        const field = insertRawIntextField(data)
        renderStyledField(field, applyIntextCitationStyle)
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

  resetTestDocument()
}

import {
  applyIntextCitationStyle,
  collectIntextCitationFieldsInRange,
  fieldContentEquals,
  readFieldData,
  renderStyledFieldWithData,
  isIntextCitation,
  isNoteCitation,
} from "../src/utils/field"
import { assert, type TestContext } from "./framework"
import {
  insertSeparator,
  makeIntextCitation,
  makeNoteCitation,
  resetTestDocument,
} from "./fixtures"
import {
  insertAutomationPendingIntextCitation,
  insertAutomationPendingNoteCitation,
  makeAutomationIntextData,
  makeAutomationNoteData,
} from "./automation"

const MODULE = "citation"
const SIZES = [10, 50, 100]

export async function runCitationTests(context: TestContext): Promise<void> {
  await context.test(MODULE, "intext insertion and readback", () => {
    resetTestDocument()
    const data = makeAutomationIntextData("citation-readback", makeIntextCitation("citation-readback", "plain").content)
    const field = insertAutomationPendingIntextCitation(data.content, data.id)
    const readBack = readFieldData(field)
    assert.ok(isIntextCitation(readBack))
    assert.equal(readBack.id, data.id)
    assert.equal(readBack.source.cites.length, 1)
    assert.equal(readBack.source.cites[0].item.key, data.source.cites[0].item.key)
    assert.equal(field.Result.Text, data.content.text)
    assert.equal(collectIntextCitationFieldsInRange(wps.ActiveDocument.Content).length, 1)
  })

  await context.test(MODULE, "note insertion and readback", () => {
    resetTestDocument()
    const note = makeNoteCitation("citation-note")
    const data = makeAutomationNoteData("citation-note", note.content, note.reference)
    const created = insertAutomationPendingNoteCitation(data.content, data.reference, data.id)
    const readBack = readFieldData(created.field)
    assert.ok(isNoteCitation(readBack))
    assert.equal(readBack.id, data.id)
    assert.equal(readBack.source.cites.length, 1)
    assert.equal(wps.ActiveDocument.Footnotes.Count, 1)
  })

  for (const size of SIZES) {
    await context.test(MODULE, `batch intext insertion size=${size}`, async () => {
      resetTestDocument()
      const fields: Wps.Field[] = []
      const data = makeAutomationIntextData("batch-citation", makeIntextCitation("batch-citation", "plain").content)
      await context.measure(MODULE, "insert intext citations", () => {
        for (let i = 0; i < size; i += 1) {
          const item = makeAutomationIntextData(`batch-${i}`, data.content)
          fields.push(insertAutomationPendingIntextCitation(item.content, item.id, item.id))
          if (i < size - 1) insertSeparator()
        }
      }, { size, mode: "insert" })
      assert.equal(fields.length, size)
      assert.equal(collectIntextCitationFieldsInRange(wps.ActiveDocument.Content).length, size)
    })

    await context.test(MODULE, `batch citation data update size=${size}`, async () => {
      resetTestDocument()
      const fields: Wps.Field[] = []
      const initial = makeAutomationIntextData("update-citation", makeIntextCitation("update-citation", "plain").content)
      for (let i = 0; i < size; i += 1) {
        const item = makeAutomationIntextData(`update-${i}`, initial.content)
        fields.push(insertAutomationPendingIntextCitation(item.content, item.id, item.id))
        if (i < size - 1) insertSeparator()
      }
      const updates = fields.map((_, i) => makeIntextCitation(`update-${i}`, "rich"))
      await context.measure(MODULE, "update citation data", () => {
        for (let i = 0; i < fields.length; i += 1) {
          fields[i].Data = JSON.stringify(updates[i])
        }
      }, { size, mode: "data-only" })
      await context.measure(MODULE, "rerender updated citations", () => {
        for (let i = 0; i < fields.length; i += 1) {
          renderStyledFieldWithData(fields[i], applyIntextCitationStyle, updates[i], updates[i].content, "character")
        }
      }, { size, mode: "content-changed" })
      assert.equal(collectIntextCitationFieldsInRange(wps.ActiveDocument.Content).length, size)
    })
  }

  await context.test(MODULE, "citation refresh no-op and change decisions", async () => {
    resetTestDocument()
    const size = 50
    const fields: Wps.Field[] = []
    const data = makeAutomationIntextData("repeat-citation", makeIntextCitation("repeat-citation", "rich").content)
    for (let i = 0; i < size; i += 1) {
      const item = makeAutomationIntextData(`repeat-${i}`, data.content)
      fields.push(insertAutomationPendingIntextCitation(item.content, item.id, item.id))
      if (i < size - 1) insertSeparator()
    }

    const current = fields.map(field => readFieldData(field))
    assert.ok(current.every(isIntextCitation))
    const unchanged = current.map(value => ({ ...value }))
    const sourceOnly = current.map(value => ({
      ...value,
      source: { ...value.source, params: { changed: true } },
    }))
    const contentChanged = current.map(value => ({
      ...value,
      content: makeIntextCitation(value.id, "plain").content,
    }))

    let noOpRenders = 0
    let noOpDataWrites = 0
    await context.measure(MODULE, "unchanged citation decision", () => {
      for (let i = 0; i < fields.length; i += 1) {
        fields[i].Data = JSON.stringify(unchanged[i])
        noOpDataWrites += 1
        if (!fieldContentEquals(current[i], unchanged[i])) {
          noOpRenders += 1
          renderStyledFieldWithData(fields[i], applyIntextCitationStyle, unchanged[i], unchanged[i].content, "character")
        }
      }
    }, { size, mode: "skip-render" })
    assert.equal(noOpDataWrites, size)
    assert.equal(noOpRenders, 0)

    let sourceOnlyRenders = 0
    let sourceOnlyDataWrites = 0
    await context.measure(MODULE, "source-only citation decision", () => {
      for (let i = 0; i < fields.length; i += 1) {
        fields[i].Data = JSON.stringify(sourceOnly[i])
        sourceOnlyDataWrites += 1
        if (!fieldContentEquals(current[i], sourceOnly[i])) {
          sourceOnlyRenders += 1
          renderStyledFieldWithData(fields[i], applyIntextCitationStyle, sourceOnly[i], sourceOnly[i].content, "character")
        }
      }
    }, { size, mode: "data-only" })
    assert.equal(sourceOnlyDataWrites, size)
    assert.equal(sourceOnlyRenders, 0)

    let contentRenders = 0
    await context.measure(MODULE, "content-changed citation decision", () => {
      for (let i = 0; i < fields.length; i += 1) {
        fields[i].Data = JSON.stringify(contentChanged[i])
        if (!fieldContentEquals(sourceOnly[i], contentChanged[i])) {
          contentRenders += 1
          renderStyledFieldWithData(fields[i], applyIntextCitationStyle, contentChanged[i], contentChanged[i].content, "character")
        }
      }
    }, { size, mode: "content-changed" })
    assert.equal(contentRenders, size)
    assert.equal(collectIntextCitationFieldsInRange(wps.ActiveDocument.Content).length, size)
  })

  resetTestDocument()
}

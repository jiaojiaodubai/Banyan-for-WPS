import {
  applyIntextCitationStyle,
  collectIntextCitationFieldsInRange,
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
        resetTestDocument()
        const previousScreenUpdating = Application.ScreenUpdating
        try {
          Application.ScreenUpdating = screenUpdating
          for (let i = 0; i < size; i += 1) {
            insertRawIntextField(makeIntextCitation(`refresh-${i}`, "plain"))
            if (i < size - 1) insertSeparator()
          }

          let pairs = collectIntextCitationFieldsInRange(wps.ActiveDocument.Content)
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
        }
      })
    }
  }
  resetTestDocument()
}

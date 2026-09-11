import {
  collectBibliographyFieldsInRange,
  deleteExistingBibliography,
  insertBibliography,
} from "../src/modules/bibliography"
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

export async function runBibliographyTests(context: TestContext): Promise<void> {
  await context.test(MODULE, "insert, collect and delete", () => {
    resetTestDocument()
    const lines = makeBibliographyLines(2)
    insertBibliography(getDocumentEnd(), lines, TEST_STYLES)
    assert.equal(collectBibliographyFieldsInRange(wps.ActiveDocument.Content).length, lines.length)
    assert.ok(wps.ActiveDocument.Bookmarks.Count >= 2)
    assert.equal(deleteExistingBibliography(wps.ActiveDocument.Content), true)
  })

  for (const size of SIZES) {
    await context.test(MODULE, `rebuild size=${size}`, async () => {
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
    })
  }
  resetTestDocument()
}

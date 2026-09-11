import type {
  BibliographyEntry,
  BibliographyLine,
  IntextCitation,
  NoteCitation,
} from "../src/typings/style"
import type { RichText } from "../src/typings/unit"

const CITATION_TEXT = "(Zhang, 2020, p. 37)"

export type RichTextVariant = "plain" | "format" | "link" | "rich"

export function makeRichText(
  variant: RichTextVariant = "rich",
  linkId = "item-1",
): RichText {
  const marks: RichText["marks"] = []
  if (variant === "format" || variant === "rich") {
    marks.push(
      { type: "bold", start: 1, end: 6, value: true },
      { type: "italic", start: 8, end: 12, value: true },
      { type: "color", start: 0, end: CITATION_TEXT.length, value: "#c00000" },
    )
  }
  if (variant === "link" || variant === "rich") {
    marks.push({
      type: "link",
      start: 0,
      end: CITATION_TEXT.length,
      value: `banyan://entry/${linkId}`,
    })
  }
  return { text: CITATION_TEXT, marks }
}

export function makeIntextCitation(
  id: string,
  variant: RichTextVariant = "rich",
): IntextCitation {
  return {
    id,
    type: "intext-citation",
    source: { cites: [], params: {} },
    content: makeRichText(variant, id),
  }
}

export function makeNoteCitation(id: string): NoteCitation {
  return {
    id,
    type: "note-citation",
    source: { cites: [], params: {} },
    content: {
      text: "Zhou, J. Self-Efficacy, 2003, pp. 37-38.",
      marks: [{ type: "italic", start: 9, end: 22, value: true }],
    },
    reference: { text: "1", marks: [] },
  }
}

export function makeBibliographyLines(count: number): BibliographyLine[] {
  const lines: BibliographyLine[] = [{
    type: "bibliography-title",
    content: { text: "References", marks: [] },
  }]
  for (let i = 1; i <= count; i += 1) {
    const text = `${i}. Zhang. A representative bibliography entry for performance testing.`
    const entry: BibliographyEntry = {
      id: `bib-${i}`,
      type: "bibliography-entry",
      content: {
        text,
        marks: [{ type: "italic", start: text.indexOf("Zhang"), end: text.indexOf("Zhang") + 5, value: true }],
      },
    }
    lines.push(entry)
  }
  return lines
}

export function resetTestDocument(): void {
  wps.ActiveDocument.Content.Text = "Banyan performance test. "
}

export function getDocumentEnd(): Wps.Range {
  const range = wps.ActiveDocument.Content.Duplicate
  range.Collapse(wps.Enum.wdCollapseEnd)
  return range
}

export function insertRawIntextField(data: IntextCitation): Wps.Field {
  const range = getDocumentEnd()
  const field = range.Fields.Add(
    range,
    wps.Enum.wdFieldAddin,
    `BANYAN_CITATION ${data.id}`,
    false,
  )
  field.Data = JSON.stringify(data)
  field.Result.Text = data.content.text
  return field
}

export function insertSeparator(): void {
  getDocumentEnd().InsertAfter(" ")
}

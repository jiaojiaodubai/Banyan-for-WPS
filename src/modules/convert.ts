import type { ConvertFieldInput, ConvertResponseData } from "../typings/http"
import type { IntextCitation, NoteCitation } from "../typings/style"
import {
  createIntextCitationAtRange,
  createNoteCitationAtRange,
  isIntextCitation,
  isNoteCitation,
} from "../utils/field"
import { getDocumentId, request } from "../utils/http"
import { useI10n } from "../utils/i10n"
import { logError, logWarn } from "../utils/log"
import { withProgress } from "../utils/progress"
import { notifyTaskpaneCitationsRefreshed } from "./taskpane"

type CitationType = "intext-citation" | "note-citation"

type ZoteroFieldTarget = {
  field: Wps.Field
  fieldCode: string
  fieldId: string
}

type ZoteroNoteFieldTarget = ZoteroFieldTarget & {
  note: Wps.Footnote | Wps.Endnote
}

type ZoteroNoteCollection = {
  readonly Count: number
  Item: (index: number) => Wps.Footnote | Wps.Endnote
}

type CustomDocumentProperty = {
  Name?: unknown
  Value?: unknown
}

type CustomDocumentProperties = {
  Count?: unknown
  Item: (indexOrName: number | string) => CustomDocumentProperty
}

type DocumentVariables = {
  Count?: unknown
  Item: (indexOrName: number | string) => {
    Name?: unknown
    Value?: unknown
  }
}

const ZOTERO_PREF_PROPERTY = "ZOTERO_PREF"
const ZOTERO_CITATION_CODE_PREFIX_PATTERNS = [
  /^(?:ADDIN\s+)?(?:ZOTERO_)?ITEM\s+CSL_CITATION\s+/i,
  /^(?:ADDIN\s+)?CSL_CITATION\s+/i,
]
const ZOTERO_NOTE_FOOTNOTE = 1
const ZOTERO_NOTE_ENDNOTE = 2

const CONVERT_MESSAGE_ZH = {
  progressReason: "正在转换 Zotero 域...",
  noDocumentPreference:
    "未检测到 Zotero 文档偏好，无法判断当前文档是正文引注还是脚注引注。",
  noFieldsFound: "未检测到可转换的 Zotero 引注域。",
  invalidResponse: "后端没有返回可用的转换结果。",
  error: "转换 Zotero 域时发生错误：{message}",
}

const CONVERT_MESSAGE_EN = {
  progressReason: "Converting Zotero fields...",
  noDocumentPreference:
    "Zotero document preferences were not found, so the document citation type could not be determined.",
  noFieldsFound: "No Zotero citation fields were found to convert.",
  invalidResponse: "The backend did not return any usable conversion results.",
  error: "An error occurred while converting Zotero fields: {message}",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { convert: CONVERT_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { convert: CONVERT_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { convert: CONVERT_MESSAGE_EN },
})

export async function onConvertEvent(): Promise<void> {
  try {
    const noteType = resolveZoteroNoteType()
    if (noteType === null) {
      alert(t("convert.noDocumentPreference"))
      return
    }

    const citationType: CitationType = noteType === ZOTERO_NOTE_FOOTNOTE
      ? "note-citation"
      : "intext-citation"
    const converted = await withProgress(t("convert.progressReason"), async () => {
      if (noteType === ZOTERO_NOTE_ENDNOTE) {
        return convertEndnoteCitationsToIntext()
      }
      return citationType === "note-citation"
        ? convertNoteCitations()
        : convertIntextCitations()
    })

    if (converted) {
      notifyTaskpaneCitationsRefreshed()
    }
  }
  catch (error) {
    logError("Convert", "Failed to convert Zotero fields.", error)
    alert(t("convert.error", { message: getErrorMessage(error) }))
  }
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message
  }
  return String(error)
}

function resolveZoteroNoteType(): number | null {
  const data = readZoteroDocumentPreferenceData()
  if (!data) {
    return null
  }

  return parseZoteroDocumentPreferenceXml(data)
    ?? parseLegacyZoteroDocumentPreference(data)
}

function readZoteroDocumentPreferenceData(): string | null {
  const activeDocument = wps.ActiveDocument as unknown as {
    CustomDocumentProperties?: CustomDocumentProperties
    Variables?: DocumentVariables
  }

  const properties = activeDocument.CustomDocumentProperties
  const fragmentedPropertyValue = readCustomPropertyByFragments(properties, ZOTERO_PREF_PROPERTY)
  if (fragmentedPropertyValue !== null) {
    return fragmentedPropertyValue
  }

  const directPropertyValue = readCustomPropertyByName(properties, ZOTERO_PREF_PROPERTY)
  if (directPropertyValue !== null) {
    return directPropertyValue
  }

  const iteratedPropertyValue = readCustomPropertyByIteration(properties, ZOTERO_PREF_PROPERTY)
  if (iteratedPropertyValue !== null) {
    return iteratedPropertyValue
  }

  const fragmentedVariableValue = readVariableByFragments(activeDocument.Variables, ZOTERO_PREF_PROPERTY)
  if (fragmentedVariableValue !== null) {
    return fragmentedVariableValue
  }

  const variableValue = readVariableByName(activeDocument.Variables, ZOTERO_PREF_PROPERTY)
  if (variableValue !== null) {
    return variableValue
  }

  logWarn("Convert", "ZOTERO_PREF not found in custom properties or variables.")
  return null
}

function readCustomPropertyByFragments(
  properties: CustomDocumentProperties | undefined,
  name: string,
): string | null {
  if (!properties) {
    return null
  }

  const count = Number(properties.Count)
  if (!Number.isFinite(count) || count <= 0) {
    return null
  }

  const normalizedPrefix = `${name.toUpperCase()}_`
  const fragments = new Map<number, string>()
  for (let i = 1; i <= count; i += 1) {
    let property: CustomDocumentProperty
    try {
      property = properties.Item(i)
    }
    catch {
      continue
    }

    const propertyName = String(property?.Name ?? "").toUpperCase()
    if (!propertyName.startsWith(normalizedPrefix)) {
      continue
    }

    const suffix = propertyName.slice(normalizedPrefix.length)
    if (!/^\d+$/.test(suffix)) {
      continue
    }

    const partIndex = Number.parseInt(suffix, 10)
    if (partIndex <= 0 || fragments.has(partIndex)) {
      continue
    }

    fragments.set(
      partIndex,
      property.Value === undefined || property.Value === null
        ? ""
        : String(property.Value),
    )
  }

  return joinSequentialFragments(fragments)
}

function joinSequentialFragments(fragments: Map<number, string>): string | null {
  if (!fragments.has(1)) {
    return null
  }

  const combined: string[] = []
  for (let i = 1; ; i += 1) {
    const part = fragments.get(i)
    if (part === undefined) {
      break
    }
    combined.push(part)
  }

  return combined.join("")
}


function readCustomPropertyByName(
  properties: CustomDocumentProperties | undefined,
  name: string,
): string | null {
  if (!properties) {
    return null
  }

  try {
    const property = properties.Item(name)
    if (property?.Value === undefined || property?.Value === null) {
      return null
    }
    if (property.Name !== undefined && String(property.Name).toUpperCase() !== name) {
      return null
    }
    return String(property.Value)
  }
  catch {
    return null
  }
}

function readCustomPropertyByIteration(
  properties: CustomDocumentProperties | undefined,
  name: string,
): string | null {
  if (!properties) {
    return null
  }

  const count = Number(properties.Count)
  if (!Number.isFinite(count) || count <= 0) {
    return null
  }

  const expectedName = name.toUpperCase()
  for (let i = 1; i <= count; i += 1) {
    let property: CustomDocumentProperty
    try {
      property = properties.Item(i)
    }
    catch {
      continue
    }

    if (String(property?.Name ?? "").toUpperCase() !== expectedName) {
      continue
    }
    if (property.Value === undefined || property.Value === null) {
      return null
    }
    return String(property.Value)
  }

  return null
}

function readVariableByFragments(
  variables: DocumentVariables | undefined,
  name: string,
): string | null {
  if (!variables) {
    return null
  }

  const count = Number(variables.Count)
  if (!Number.isFinite(count) || count <= 0) {
    return null
  }

  const normalizedPrefix = `${name.toUpperCase()}_`
  const fragments = new Map<number, string>()
  for (let i = 1; i <= count; i += 1) {
    let variable: { Name?: unknown; Value?: unknown }
    try {
      variable = variables.Item(i)
    }
    catch {
      continue
    }

    const variableName = String(variable?.Name ?? "").toUpperCase()
    if (!variableName.startsWith(normalizedPrefix)) {
      continue
    }

    const suffix = variableName.slice(normalizedPrefix.length)
    if (!/^\d+$/.test(suffix)) {
      continue
    }

    const partIndex = Number.parseInt(suffix, 10)
    if (partIndex <= 0 || fragments.has(partIndex)) {
      continue
    }

    fragments.set(
      partIndex,
      variable.Value === undefined || variable.Value === null
        ? ""
        : String(variable.Value),
    )
  }

  return joinSequentialFragments(fragments)
}

function readVariableByName(
  variables: DocumentVariables | undefined,
  name: string,
): string | null {
  if (!variables) {
    return null
  }

  try {
    const variable = variables.Item(name)
    if (variable?.Value === undefined || variable?.Value === null) {
      return null
    }
    return String(variable.Value)
  }
  catch {
    // Fall back to iteration for hosts that do not support name-based access.
  }

  const count = Number(variables.Count)
  if (!Number.isFinite(count) || count <= 0) {
    return null
  }

  const expectedName = name.toUpperCase()
  for (let i = 1; i <= count; i += 1) {
    let variable: { Name?: unknown; Value?: unknown }
    try {
      variable = variables.Item(i)
    }
    catch {
      continue
    }

    if (String(variable?.Name ?? "").toUpperCase() !== expectedName) {
      continue
    }
    if (variable.Value === undefined || variable.Value === null) {
      return null
    }
    return String(variable.Value)
  }

  return null
}

function parseZoteroDocumentPreferenceXml(data: string): number | null {
  const trimmed = data.trim()
  if (!trimmed.startsWith("<")) {
    return null
  }

  const doc = new DOMParser().parseFromString(trimmed, "application/xml")
  if (doc.getElementsByTagName("parsererror").length > 0) {
    return null
  }

  const prefs = doc.getElementsByTagName("pref")
  for (let i = 0; i < prefs.length; i += 1) {
    const pref = prefs.item(i)
    if (!pref || pref.getAttribute("name") !== "noteType") {
      continue
    }
    const value = pref.getAttribute("value")
    if (value === null || value.trim() === "") {
      return 0
    }
    const noteType = Number.parseInt(value, 10)
    return Number.isNaN(noteType) ? null : noteType
  }

  return 0
}

function parseLegacyZoteroDocumentPreference(data: string): number | null {
  const parts = data.split("::")
  if (parts.length < 5) {
    return null
  }

  if (parts[2] !== "note") {
    return 0
  }

  return parts[4] === "1" || parts[4] === "True"
    ? ZOTERO_NOTE_ENDNOTE
    : ZOTERO_NOTE_FOOTNOTE
}

async function convertIntextCitations(): Promise<boolean> {
  const targets = collectZoteroIntextFields()
  if (targets.length === 0) {
    alert(t("convert.noFieldsFound"))
    return false
  }

  const response = await requestConvert("intext-citation", targets)
  if (!response) {
    return false
  }

  let convertedCount = 0
  for (let i = targets.length - 1; i >= 0; i -= 1) {
    const target = targets[i]
    const converted = getConvertedIntextCitation(response, target.fieldId)
    if (!converted) {
      continue
    }

    const insertRange = createCollapsedRange(target.field.Result.Start)
    target.field.Delete()
    createIntextCitationAtRange(insertRange, converted)
    convertedCount += 1
  }

  if (convertedCount === 0) {
    alert(t("convert.invalidResponse"))
    return false
  }
  return true
}

async function convertNoteCitations(): Promise<boolean> {
  const targets = collectZoteroNoteFields(wps.ActiveDocument.Footnotes)
  if (targets.length === 0) {
    alert(t("convert.noFieldsFound"))
    return false
  }

  const response = await requestConvert("note-citation", targets)
  if (!response) {
    return false
  }

  let convertedCount = 0
  for (let i = targets.length - 1; i >= 0; i -= 1) {
    const target = targets[i]
    const converted = getConvertedNoteCitation(response, target.fieldId)
    if (!converted) {
      continue
    }

    const insertRange = createCollapsedRange(target.note.Reference.Start)
    target.note.Delete()
    createNoteCitationAtRange(insertRange, converted)
    convertedCount += 1
  }

  if (convertedCount === 0) {
    alert(t("convert.invalidResponse"))
    return false
  }
  return true
}

async function convertEndnoteCitationsToIntext(): Promise<boolean> {
  const targets = collectZoteroNoteFields(wps.ActiveDocument.Endnotes)
  if (targets.length === 0) {
    alert(t("convert.noFieldsFound"))
    return false
  }

  const response = await requestConvert("intext-citation", targets)
  if (!response) {
    return false
  }

  let convertedCount = 0
  for (let i = targets.length - 1; i >= 0; i -= 1) {
    const target = targets[i]
    const converted = getConvertedIntextCitation(response, target.fieldId)
    if (!converted) {
      continue
    }

    const insertRange = createCollapsedRange(target.note.Reference.Start)
    target.note.Delete()
    createIntextCitationAtRange(insertRange, converted)
    convertedCount += 1
  }

  if (convertedCount === 0) {
    alert(t("convert.invalidResponse"))
    return false
  }
  return true
}

async function requestConvert(
  citationType: CitationType,
  targets: ZoteroFieldTarget[] | ZoteroNoteFieldTarget[],
): Promise<ConvertResponseData | null> {
  const fieldCodesById = new Map<string, string>()
  for (const target of targets) {
    const previous = fieldCodesById.get(target.fieldId)
    if (previous !== undefined && previous !== target.fieldCode) {
      logWarn(
        "Convert",
        `Found duplicated Zotero citation id with different field codes: ${target.fieldId}. Using the first occurrence.`
      )
      continue
    }
    fieldCodesById.set(target.fieldId, target.fieldCode)
  }

  const fields: ConvertFieldInput[] = [...fieldCodesById.entries()].map(([fieldId, fieldCode]) => ({
    fieldId,
    fieldCode,
  }))

  return request("convert", {
    documentId: getDocumentId(),
    citationType,
    fields,
  })
}

function collectZoteroIntextFields(): ZoteroFieldTarget[] {
  const targets: ZoteroFieldTarget[] = []
  const fields = wps.ActiveDocument.Fields

  for (let i = 1; i <= fields.Count; i += 1) {
    const field = fields.Item(i)
    const target = parseZoteroCitationField(field)
    if (target) {
      targets.push({
        field,
        fieldCode: target.fieldCode,
        fieldId: target.fieldId,
      })
    }
  }

  return targets
}

function collectZoteroNoteFields(notes: ZoteroNoteCollection): ZoteroNoteFieldTarget[] {
  const targets: ZoteroNoteFieldTarget[] = []

  for (let i = 1; i <= notes.Count; i += 1) {
    const note = notes.Item(i)
    for (let j = 1; j <= note.Range.Fields.Count; j += 1) {
      const field = note.Range.Fields.Item(j)
      const target = parseZoteroCitationField(field)
      if (!target) {
        continue
      }

      targets.push({
        note,
        field,
        fieldCode: target.fieldCode,
        fieldId: target.fieldId,
      })
      break
    }
  }

  return targets
}

function parseZoteroCitationField(field: Wps.Field): {
  fieldCode: string
  fieldId: string
} | null {
  const fieldCode = readZoteroCitationCode(field)
  if (!fieldCode) {
    return null
  }

  const fieldId = readZoteroCitationId(fieldCode)
  if (!fieldId) {
    logWarn("Convert", "Skipped Zotero field without citationID.")
    return null
  }

  return {
    fieldCode,
    fieldId,
  }
}

function readZoteroCitationCode(field: Wps.Field): string | null {
  try {
    const codeText = String(field.Code.Text ?? "")
      .replace(/[\r\n]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()

    if (!ZOTERO_CITATION_CODE_PREFIX_PATTERNS.some((pattern) => pattern.test(codeText))) {
      return null
    }

    const jsonStart = codeText.indexOf("{")
    const jsonEnd = codeText.lastIndexOf("}")
    if (jsonStart === -1 || jsonEnd < jsonStart) {
      return null
    }

    return codeText.slice(jsonStart, jsonEnd + 1)
  }
  catch (error) {
    logWarn("Convert", "Failed to read Zotero field code.", error)
    return null
  }
}

function readZoteroCitationId(fieldCode: string): string | null {
  try {
    const parsed = JSON.parse(fieldCode) as {
      citationID?: unknown
      CITATIONID?: unknown
    }

    if (typeof parsed.citationID === "string" && parsed.citationID) {
      return parsed.citationID
    }
    if (typeof parsed.CITATIONID === "string" && parsed.CITATIONID) {
      return parsed.CITATIONID
    }
    return null
  }
  catch (error) {
    logWarn("Convert", "Failed to parse Zotero citation code JSON.", error)
    return null
  }
}

function getConvertedIntextCitation(
  response: ConvertResponseData,
  fieldId: string,
): IntextCitation | null {
  const data = (response as Record<string, unknown>)[fieldId]
  if (!isIntextCitation(data)) {
    return null
  }
  return data
}

function getConvertedNoteCitation(
  response: ConvertResponseData,
  fieldId: string,
): NoteCitation | null {
  const data = (response as Record<string, unknown>)[fieldId]
  if (!isNoteCitation(data)) {
    return null
  }
  return data
}

function createCollapsedRange(position: number): Wps.Range {
  const range = wps.ActiveDocument.Range().Duplicate
  range.SetRange(position, position)
  return range
}

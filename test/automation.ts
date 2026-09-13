import type { CitationSource, IntextCitation, NoteCitation } from "../src/typings/style"
import {
  createNoteCitationAtRange,
  createIntextCitationAtRange,
  createPlaceholderIntextCitationData,
  createPlaceholderNoteCitationData,
} from "../src/utils/field"
import type { RichText } from "../src/typings/unit"
import { getDocumentEnd } from "./fixtures"

/**
 * 无对话框的测试自动化，参照 Word 侧的 modAutomation.bas。
 *
 * 这些辅助函数直接通过域工厂注入域数据，有意不调用 GUI 的引注/样式事件及其
 * 路由，以便性能测试无人值守且结果稳定。
 */
export function makeAutomationCitationSource(itemKey: string): CitationSource {
  const item = {
    id: 0,
    key: itemKey,
    itemType: "journalArticle",
    uri: `http://zotero.org/users/0/items/${itemKey}`,
    title: `Automation item ${itemKey}`,
    date: "2020",
    year: "2020",
    firstCreator: "Automation Author",
    language: "en-US",
    creators: [{ creatorType: "author", firstName: "Automation", lastName: "Author" }],
    tags: [],
    extra: {},
  }
  return {
    cites: [{ item: item as unknown as import("../src/typings/item").Item, params: {} }],
    params: {},
  }
}

export function makeAutomationIntextData(
  id: string,
  content: RichText,
  itemKey = id,
): IntextCitation {
  return {
    id,
    type: "intext-citation",
    source: makeAutomationCitationSource(itemKey),
    content,
  }
}

export function makeAutomationNoteData(
  id: string,
  content: RichText,
  reference: RichText,
  itemKey = id,
): NoteCitation {
  return {
    id,
    type: "note-citation",
    source: makeAutomationCitationSource(itemKey),
    content,
    reference,
  }
}

export function insertAutomationPendingIntextCitation(
  content: RichText,
  id: string,
  itemKey = id,
) {
  const data = createPlaceholderIntextCitationData(id, makeAutomationCitationSource(itemKey))
  data.content = content
  return createIntextCitationAtRange(getDocumentEnd(), data)
}

export function insertAutomationPendingNoteCitation(
  content: RichText,
  reference: RichText,
  id: string,
  itemKey = id,
) {
  const data = createPlaceholderNoteCitationData(id, makeAutomationCitationSource(itemKey))
  data.content = content
  data.reference = reference
  return createNoteCitationAtRange(getDocumentEnd(), data)
}

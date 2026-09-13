import { request, getDocumentId } from "../utils/http"
import type { CitationSource, StyleInfo } from "../typings/style"
import { useI10n } from "../utils/i10n"
import { openDialog } from "../utils/window"
import { findPreviousChapterBreak } from "./chapter-break"

export type PrefStyle = Pick<StyleInfo, "id" | "title" | "citationType">

export type GlobalPreference = {
  // 全局：刷新时从 Zotero 客户端同步条目元数据，或仅在插入/编辑引注时同步
  syncItems: boolean
  // 全局：刷新时处理全部章节的域，或只处理当前章节（前后两个章节分隔符之间）
  refreshAll: boolean
}

export type ChapterPreference = {
  // 章节级
  style: PrefStyle
  // 章节级：在书目编辑器中手工添加的未被引用的文献
  extraSource?: CitationSource
  // 章节级：书目标题的 Word 样式名，首章必填，后续章节继承
  bibliographyTitleStyle: string
  // 章节级：书目题录的 Word 样式名，首章必填，后续章节继承
  bibliographyEntryStyle: string
}

export type Preference =  GlobalPreference & ChapterPreference

const PREFERENCE_MESSAGE_ZH = {
  dialogTitle: "Banyan 设置",
  defaultBibliographyTitleStyle: "文献列表标题",
  defaultBibliographyEntryStyle: "文献列表题录",
}
const PREFERENCE_MESSAGE_EN = {
  dialogTitle: "Banyan Preferences",
  defaultBibliographyTitleStyle: "Bibliography Title",
  defaultBibliographyEntryStyle: "Bibliography Entry",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { preference: PREFERENCE_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { preference: PREFERENCE_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { preference: PREFERENCE_MESSAGE_EN },
})

const PREFERENCE_PROPERTY = "BANYAN_PREF"
const MSO_PROPERTY_TYPE_STRING = 4

const preferenceUrl = import.meta.env.DEV
  ? new URL("/src/ui/preference.html", window.location.href).toString()
  : new URL("ui/preference.html", window.location.href).toString()

export async function onPreferenceEvent(): Promise<void> {
  const pref = await getPreference()
  if (!pref) {
    return
  }
  openDialog(preferenceUrl, t("preference.dialogTitle"), 520, 340, true)
}

export async function getPreference(): Promise<Preference | null> {
  const pref = getLocalPreference()
  if (pref) {
    return pref
  }
  return createPreference()
}

export function getDocumentPreference(): Preference | null {
  return getPrefFromDocumentProperty()
}

export function getLocalPreference(): Preference | null {
  const partPref = getPrefFromDocumentProperty()
  if (!partPref) {
    return null
  }
  const prevChapterBreak = findPreviousChapterBreak()
  if (!prevChapterBreak) {
    return partPref
  }
  return {
    syncItems: partPref.syncItems,
    refreshAll: partPref.refreshAll,
    style: prevChapterBreak.data.style,
    extraSource: prevChapterBreak.data.extraSource,
    bibliographyTitleStyle: prevChapterBreak.data.bibliographyTitleStyle,
    bibliographyEntryStyle: prevChapterBreak.data.bibliographyEntryStyle,
  }
}

function getPrefFromDocumentProperty(): Preference | null {
  const props = (wps.ActiveDocument as unknown as { CustomDocumentProperties?: { Item: (n: string) => { Value?: unknown } } }).CustomDocumentProperties
  if (!props) return null
  try {
    const raw = props.Item(PREFERENCE_PROPERTY).Value
    if (typeof raw !== "string") return null
    const parsed = JSON.parse(raw) as Preference
    return parsed && parsed.style ? parsed : null
  }
  catch { return null }
}

function saveDocumentProperty(pref: Preference) {
  const doc = wps.ActiveDocument as unknown as {
    CustomDocumentProperties?: { Item: (n: string) => { Value: unknown }; Add: (n: string, link: boolean, type: number, value: string) => unknown }
  }
  const props = doc.CustomDocumentProperties
  if (!props) return
  const value = JSON.stringify(pref)
  try { props.Item(PREFERENCE_PROPERTY).Value = value; return }
  catch { /* property does not exist */ }
  try { props.Add(PREFERENCE_PROPERTY, false, MSO_PROPERTY_TYPE_STRING, value) }
  catch { /* host may expose a restricted property collection */ }
}

export async function savePreference(pref: Preference) {
  const previousChapterBreak = findPreviousChapterBreak()
  const documentPref = getDocumentPreference()
  if (previousChapterBreak) {
    saveDocumentProperty({
      ...(documentPref ?? pref),
      syncItems: pref.syncItems,
      refreshAll: pref.refreshAll,
    })
    previousChapterBreak.field.Data = JSON.stringify({
      ...previousChapterBreak.data,
      style: pref.style,
      // extraSource 未显式提供（例如设置对话框）时保留已有值，避免误清除手动添加的文献
      extraSource: pref.extraSource ?? previousChapterBreak.data.extraSource,
      bibliographyTitleStyle: pref.bibliographyTitleStyle,
      bibliographyEntryStyle: pref.bibliographyEntryStyle,
    })
  }
  else {
    saveDocumentProperty({
      ...(documentPref ?? pref),
      syncItems: pref.syncItems,
      refreshAll: pref.refreshAll,
      style: pref.style,
      extraSource: pref.extraSource ?? documentPref?.extraSource,
      bibliographyTitleStyle: pref.bibliographyTitleStyle,
      bibliographyEntryStyle: pref.bibliographyEntryStyle,
    })
  }
}

export function removePreferencePart() {
  const props = (wps.ActiveDocument as unknown as { CustomDocumentProperties?: { Item: (n: string) => { Delete?: () => void } } }).CustomDocumentProperties
  try {
    props?.Item(PREFERENCE_PROPERTY).Delete?.()
  }
  catch {
    /* absent */
  }
}

async function createPreference(): Promise<Preference | null> {
  const style = await request("style", { documentId: getDocumentId() })
  if (!style) {
    return null
  }
  const pref: Preference = {
    syncItems: true,
    refreshAll: true,
    style: {
      id: style.id,
      title: style.title,
      citationType: style.citationType,
    },
    bibliographyTitleStyle: t("preference.defaultBibliographyTitleStyle"),
    bibliographyEntryStyle: t("preference.defaultBibliographyEntryStyle"),
  }
  saveDocumentProperty(pref)
  return pref
}

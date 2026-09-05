import { getDocumentId, request } from "../utils/http"
import { getConfig } from "../utils/config"
import { useI10n } from "../utils/i10n"
import { logError } from "../utils/log"
import { getPreference, savePreference, type Preference } from "../modules/preference"
import { refreshForStyleChange } from "../modules/refresh"
import "./components/banyan-checkbox"
import type { BanyanCheckboxElement } from "./components/banyan-checkbox"

type ThemeMode = "light" | "dark"

const THEME_SYNC_CHANNEL_NAME = "banyan-theme-sync"

const LABELS_ZH = {
  global: "全局设置",
  section: "章节设置",
  syncItems: "刷新时同步条目元数据",
  refreshAll: "刷新全部章节（否则仅当前章节）",
  style: "引注样式（Banyan 样式）",
  styleUnset: "未选择样式",
  fetchStyle: "选择...",
  fetchingStyle: "请求中...",
  bibTitleStyle: "文献列表标题（Word样式）",
  bibTitleStylePlaceholder: "输入Word样式名称",
  bibEntryStyle: "文献列表题录（Word样式）",
  bibEntryStylePlaceholder: "输入Word样式名称",
  defaultBibliographyTitleStyle: "文献列表标题",
  defaultBibliographyEntryStyle: "文献列表题录",
  clear: "清除",
  ok: "确定",
  cancel: "取消",
  saveError: "保存设置失败：{message}",
  loadError: "读取设置失败：{message}",
  styleError: "获取样式失败：{message}",
}

const LABELS_EN = {
  global: "Global",
  section: "Section",
  syncItems: "Sync item metadata when refreshing citations",
  refreshAll: "Refresh all chapters (otherwise current chapter only)",
  style: "Banyan style",
  styleUnset: "No style selected",
  fetchStyle: "Choose...",
  fetchingStyle: "Loading...",
  bibTitleStyle: "Bibliography title (Word style)",
  bibTitleStylePlaceholder: "Enter Word style name",
  bibEntryStyle: "Bibliography entry (Word style)",
  bibEntryStylePlaceholder: "Enter Word style name",
  defaultBibliographyTitleStyle: "Bibliography Title",
  defaultBibliographyEntryStyle: "Bibliography Entry",
  clear: "Clear",
  ok: "OK",
  cancel: "Cancel",
  saveError: "Failed to save preference: {message}",
  loadError: "Failed to load preference: {message}",
  styleError: "Failed to fetch style: {message}",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { preferenceUi: LABELS_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { preferenceUi: LABELS_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { preferenceUi: LABELS_EN },
})

function createDefaultState(): Preference {
  return {
    syncItems: true,
    refreshAll: true,
    style: { id: "", title: "", citationType: "intext-citation" },
    bibliographyTitleStyle: t("preferenceUi.defaultBibliographyTitleStyle"),
    bibliographyEntryStyle: t("preferenceUi.defaultBibliographyEntryStyle"),
  }
}

const state: Preference = {
  syncItems: true,
  refreshAll: true,
  style: { id: "", title: "", citationType: "intext-citation" },
  bibliographyTitleStyle: "",
  bibliographyEntryStyle: "",
}

function resetStateToDefaults(): void {
  const defaultState = createDefaultState()
  state.syncItems = defaultState.syncItems
  state.refreshAll = defaultState.refreshAll
  state.style = defaultState.style
  state.extraSource = undefined
  state.bibliographyTitleStyle = defaultState.bibliographyTitleStyle
  state.bibliographyEntryStyle = defaultState.bibliographyEntryStyle
}

type OptionKey = "syncItems" | "refreshAll"

let currentThemeMode: ThemeMode | null = null
let themeChannel: BroadcastChannel | null = null
let initialStyle: Preference["style"] | undefined

let nodes!: {
  global: HTMLHeadingElement
  section: HTMLHeadingElement
  styleLabel: HTMLDivElement
  styleTitle: HTMLInputElement
  fetchStyleBtn: HTMLButtonElement
  bibTitleStyleLabel: HTMLDivElement
  bibTitleStyle: HTMLInputElement
  clearBibTitleStyleBtn: HTMLButtonElement
  bibEntryStyleLabel: HTMLDivElement
  bibEntryStyle: HTMLInputElement
  clearBibEntryStyleBtn: HTMLButtonElement
  okBtn: HTMLButtonElement
  cancelBtn: HTMLButtonElement
}

let optionNodes!: Record<OptionKey, BanyanCheckboxElement>

function initializeNodes() {
  nodes = {
    global: document.getElementById("global-title") as HTMLHeadingElement,
    section: document.getElementById("chapter-title") as HTMLHeadingElement,
    styleLabel: document.getElementById("style-label") as HTMLDivElement,
    styleTitle: document.getElementById("style-title") as HTMLInputElement,
    fetchStyleBtn: document.getElementById("fetch-style-btn") as HTMLButtonElement,
    bibTitleStyleLabel: document.getElementById("bib-title-style-label") as HTMLDivElement,
    bibTitleStyle: document.getElementById("bib-title-style") as HTMLInputElement,
    clearBibTitleStyleBtn: document.getElementById("clear-bib-title-style-btn") as HTMLButtonElement,
    bibEntryStyleLabel: document.getElementById("bib-entry-style-label") as HTMLDivElement,
    bibEntryStyle: document.getElementById("bib-entry-style") as HTMLInputElement,
    clearBibEntryStyleBtn: document.getElementById("clear-bib-entry-style-btn") as HTMLButtonElement,
    okBtn: document.getElementById("ok-btn") as HTMLButtonElement,
    cancelBtn: document.getElementById("cancel-btn") as HTMLButtonElement,
  }

  optionNodes = {
    syncItems: document.getElementById("syncItems") as BanyanCheckboxElement,
    refreshAll: document.getElementById("refreshAll") as BanyanCheckboxElement,
  }
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message
  }
  return String(error)
}

function cloneStyle(style: Preference["style"] | undefined): Preference["style"] | undefined {
  return style ? { ...style } : undefined
}

function readThemeMode(): ThemeMode {
  const stored = getConfig("ribbon_theme_mode", "light")
  return stored === "dark" ? "dark" : "light"
}

function applyThemeMode(mode: ThemeMode): void {
  if (currentThemeMode === mode) {
    return
  }
  currentThemeMode = mode
  document.body.dataset.theme = mode
}

function syncThemeMode(): void {
  applyThemeMode(readThemeMode())
}

function parseThemeMode(payload: unknown): ThemeMode | null {
  if (typeof payload !== "string") {
    return null
  }
  try {
    const parsed = JSON.parse(payload) as { mode?: unknown }
    return parsed.mode === "dark" || parsed.mode === "light"
      ? parsed.mode
      : null
  }
  catch {
    return null
  }
}

function setupThemeSync(): void {
  syncThemeMode()

  try {
    themeChannel = new BroadcastChannel(THEME_SYNC_CHANNEL_NAME)
    themeChannel.addEventListener("message", (event) => {
      const mode = parseThemeMode(event.data)
      if (mode) {
        applyThemeMode(mode)
      }
    })
  }
  catch {
    themeChannel = null
  }
}

function disposeThemeSync(): void {
  if (themeChannel) {
    themeChannel.close()
    themeChannel = null
  }
}

function renderLabels() {
  nodes.global.textContent = t("preferenceUi.global")
  nodes.section.textContent = t("preferenceUi.section")
  optionNodes.syncItems.label = t("preferenceUi.syncItems")
  optionNodes.refreshAll.label = t("preferenceUi.refreshAll")
  nodes.styleLabel.textContent = t("preferenceUi.style")
  nodes.fetchStyleBtn.textContent = t("preferenceUi.fetchStyle")
  nodes.bibTitleStyleLabel.textContent = t("preferenceUi.bibTitleStyle")
  nodes.bibTitleStyle.placeholder = t("preferenceUi.bibTitleStylePlaceholder")
  nodes.clearBibTitleStyleBtn.textContent = t("preferenceUi.clear")
  nodes.bibEntryStyleLabel.textContent = t("preferenceUi.bibEntryStyle")
  nodes.bibEntryStyle.placeholder = t("preferenceUi.bibEntryStylePlaceholder")
  nodes.clearBibEntryStyleBtn.textContent = t("preferenceUi.clear")
  nodes.okBtn.textContent = t("preferenceUi.ok")
  nodes.cancelBtn.textContent = t("preferenceUi.cancel")
}

function renderState() {
  optionNodes.syncItems.checked = state.syncItems
  optionNodes.refreshAll.checked = state.refreshAll
  nodes.styleTitle.value = state.style?.title ?? ""
  nodes.styleTitle.placeholder = t("preferenceUi.styleUnset")
  nodes.bibTitleStyle.value = state.bibliographyTitleStyle
  nodes.bibEntryStyle.value = state.bibliographyEntryStyle
}

function bindOptionEvents() {
  optionNodes.syncItems.addEventListener("change", () => {
    state.syncItems = optionNodes.syncItems.checked
  })
  optionNodes.refreshAll.addEventListener("change", () => {
    state.refreshAll = optionNodes.refreshAll.checked
  })
}

async function loadPreference() {
  try {
    const pref = await getPreference()
    if (!pref) {
      renderState()
      return
    }

    initialStyle = cloneStyle(pref.style)

    state.syncItems = pref.syncItems
    state.refreshAll = pref.refreshAll
    state.style = pref.style
    state.bibliographyTitleStyle = pref.bibliographyTitleStyle
    state.bibliographyEntryStyle = pref.bibliographyEntryStyle
    renderState()
  }
  catch (error) {
    renderState()
    logError("PreferenceUI", "Failed to load preference.", error)
    alert(t("preferenceUi.loadError", { message: getErrorMessage(error) }))
  }
}

async function fetchStyle() {
  const originalText = nodes.fetchStyleBtn.textContent
  nodes.fetchStyleBtn.disabled = true
  nodes.fetchStyleBtn.textContent = t("preferenceUi.fetchingStyle")
  try {
    const style = await request(
      "style",
      state.style
        ? { documentId: getDocumentId(), id: state.style.id, title: state.style.title }
        : { documentId: getDocumentId() },
    )
    if (style) {
      state.style = {
        id: style.id,
        title: style.title,
        citationType: style.citationType,
      }
      renderState()
    }
  }
  catch (error) {
    alert(t("preferenceUi.styleError", { message: getErrorMessage(error) }))
  }
  finally {
    nodes.fetchStyleBtn.disabled = false
    nodes.fetchStyleBtn.textContent = originalText || t("preferenceUi.fetchStyle")
  }
}

async function saveAndClose() {
  try {
    state.syncItems = optionNodes.syncItems.checked
    state.refreshAll = optionNodes.refreshAll.checked
    state.bibliographyTitleStyle = nodes.bibTitleStyle.value.trim()
    state.bibliographyEntryStyle = nodes.bibEntryStyle.value.trim()
    const previousStyle = cloneStyle(initialStyle)
    savePreference(state)
    await refreshForStyleChange(previousStyle, state.style)
    disposeThemeSync()
    window.close()
  }
  catch (error) {
    alert(t("preferenceUi.saveError", { message: getErrorMessage(error) }))
  }
}

function bindEvents() {
  nodes.fetchStyleBtn.addEventListener("click", () => {
    void fetchStyle()
  })
  nodes.clearBibTitleStyleBtn.addEventListener("click", () => {
    nodes.bibTitleStyle.value = ""
  })
  nodes.clearBibEntryStyleBtn.addEventListener("click", () => {
    nodes.bibEntryStyle.value = ""
  })
  nodes.okBtn.addEventListener("click", () => {
    void saveAndClose()
  })
  nodes.cancelBtn.addEventListener("click", () => {
    disposeThemeSync()
    window.close()
  })
}

function initialize() {
  initializeNodes()
  setupThemeSync()
  resetStateToDefaults()
  renderLabels()
  renderState()
  bindOptionEvents()
  bindEvents()
  scheduleLoadPreference()
}

function scheduleLoadPreference(): void {
  const load = () => {
    void loadPreference()
  }

  if (typeof window.requestAnimationFrame === "function") {
    window.requestAnimationFrame(() => {
      setTimeout(load, 0)
    })
    return
  }

  setTimeout(load, 0)
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initialize)
}
else {
  initialize()
}

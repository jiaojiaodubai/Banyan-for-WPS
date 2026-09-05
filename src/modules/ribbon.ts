import { getConfig, setConfig } from "../utils/config"
import { useI10n } from "../utils/i10n"
import {
  onPreferenceEvent,
} from "./preference"
import { onInsertChapterBreakEvent } from "./chapter-break"
import { onCitationEvent } from "./citation"
import { onBibliographyEvent } from "./bibliography"
import { onRefreshEvent } from "./refresh"
import { onConvertEvent } from "./convert"
import { onFinalizeEvent } from "./finalize"
import { toggleTaskpaneVisibility as onTaskpaneEvent } from "./taskpane"
import { withOperationLock } from "../utils/operation-lock"

type RibbonControl = { Id: string }
type ThemeMode = "light" | "dark"

const THEME_SYNC_CHANNEL_NAME = "banyan-theme-sync"
const ICON_BUTTON_IDS = [
  "btnCitation",
  "btnChapterBreak",
  "btnBibliography",
  "btnCitationPane",
  "btnRefresh",
  "btnConvert",
  "btnUnlink",
  "btnDebugFieldData",
  "btnSettings",
  "btnDarkTheme",
]

// 这个函数在整个 WPS 加载项中最先执行，用于缓存 ribbonUI。
function OnAddinLoad(ribbonUI: WPS.RibbonUi) {
  const app = Application as WPSMergedRoot & { ribbonUI?: WPS.RibbonUi }
  if (!app.ribbonUI) {
    app.ribbonUI = ribbonUI
  }
  ensureThemeModeInitialized()
  return true
}

const LABELS_ZH = {
  btnCitation: "插入/编辑引注",
  btnChapterBreak: "插入分隔符",
  btnBibliography: "插入/编辑书目",
  btnCitationPane: "打开引注窗格",
  btnRefresh: "刷新",
  btnConvert: "转换 Zotero 域",
  btnUnlink: "定稿",
  btnDebugFieldData: "调试：输出字段数据",
  btnSettings: "设置",
  btnDarkTheme: "暗色主题",
}
const LABELS_EN = {
  btnCitation: "Insert/Edit Citation",
  btnChapterBreak: "Insert Break",
  btnBibliography: "Insert/Edit Bibliography",
  btnCitationPane: "Open Citation Pane",
  btnRefresh: "Refresh",
  btnConvert: "Convert Zotero Fields",
  btnUnlink: "Finalize",
  btnDebugFieldData: "Debug: Log Field Data",
  btnSettings: "Preferences",
  btnDarkTheme: "Dark Theme",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { ribbon: LABELS_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { ribbon: LABELS_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { ribbon: LABELS_EN },
})

function OnGetLabel(control: RibbonControl) {
  return t(`ribbon.${control.Id}`)
}

function OnAction(arg1: unknown, arg2?: unknown) {
  const isControl = (v: unknown): v is RibbonControl => typeof v === "object" && v !== null && "Id" in v
  // Standard Office behavior for toggleButton onAction is (control, pressed).
  // WPS currently may pass reversed arguments for toggleButton: (pressed, control).
  const control = isControl(arg2) ? arg2 : arg1 as RibbonControl

  switch (control.Id) {
    case "btnCitationPane": {
      onTaskpaneEvent()
      break
    }
    case "btnCitation":
      void withOperationLock(() => onCitationEvent(), undefined, "citation")
      break
    case "btnChapterBreak":
      void withOperationLock(() => onInsertChapterBreakEvent(), undefined, "chapter-break")
      break
    case "btnBibliography":
      void withOperationLock(() => onBibliographyEvent(), undefined, "bibliography")
      break
    case "btnRefresh":
      void withOperationLock(() => onRefreshEvent(), undefined, "refresh")
      break
    case "btnConvert":
      void withOperationLock(() => onConvertEvent(), undefined, "convert")
      break
    case "btnUnlink":
      void withOperationLock(() => onFinalizeEvent(), undefined, "finalize")
      break
    case "btnSettings":
      void withOperationLock(() => onPreferenceEvent(), undefined, "open-preference-dialog")
      break
    case "btnDebugFieldData": {
      logFirstSelectedFieldData()
      break
    }
    case "btnDarkTheme": {
      const pressed = isControl(arg2) ? arg1 : arg2
      setThemeMode(pressed ? "dark" : "light")
      refreshRibbonIcons()
      break
    }
    default:
      break
  }
  return true
}

function ensureThemeModeInitialized() {
  const stored = getConfig("ribbon_theme_mode", "light")
  if (stored === "light" || stored === "dark") {
    return
  }
  setConfig("ribbon_theme_mode", "light")
}

function getThemeMode(): ThemeMode {
  const stored = getConfig("ribbon_theme_mode", "light")
  return stored === "dark" ? "dark" : "light"
}

function setThemeMode(mode: ThemeMode) {
  setConfig("ribbon_theme_mode", mode, { persistent: true })
  broadcastThemeMode(mode)
}

function broadcastThemeMode(mode: ThemeMode) {
  const payload = JSON.stringify({
    type: "theme-changed",
    mode,
    ts: Date.now(),
  })
  try {
    const channel = new BroadcastChannel(THEME_SYNC_CHANNEL_NAME)
    channel.postMessage(payload)
    channel.close()
  }
  catch {
    // Ignore environment without BroadcastChannel support.
  }
}

function refreshRibbonIcons() {
  const app = Application
  app.ribbonUI.Invalidate()
  for (const controlId of ICON_BUTTON_IDS) {
    app.ribbonUI?.InvalidateControl(controlId)
  }
  Application.UpdateRibbon()
}

function logFirstSelectedFieldData() {
  const range = wps.Selection.Range.Duplicate
  if (range.Fields.Count === 0) {
    console.warn("[Banyan Debug] Selected range does not contain any field.")
    return
  }

  const field = range.Fields.Item(1)
  try {
    const parsed = JSON.parse(field.Data)
    console.debug("[Banyan Debug] First selected field data:", parsed)
  }
  catch (error) {
    console.error("[Banyan Debug] Failed to parse first selected field data.", {
      data: field.Data,
      error,
    })
  }
}

function OnGetImage(control: RibbonControl) {
  const isDark = getThemeMode() === "dark"

  switch (control.Id) {
    case "btnCitation":
      return isDark ? "assets/format-quote-open-dark.svg" : "assets/format-quote-open.svg"
    case "btnChapterBreak":
      return isDark ? "assets/format-page-break-dark.svg" : "assets/format-page-break.svg"
    case "btnBibliography":
      return isDark ? "assets/format-list-numbered-dark.svg" : "assets/format-list-numbered.svg"
    case "btnCitationPane":
      return isDark ? "assets/dock-right-dark.svg" : "assets/dock-right.svg"
    case "btnRefresh":
      return isDark ? "assets/autorenew-dark.svg" : "assets/autorenew.svg"
    case "btnConvert":
      return isDark ? "assets/code-json-dark.svg" : "assets/code-json.svg"
    case "btnUnlink":
      return isDark ? "assets/link-variant-off-dark.svg" : "assets/link-variant-off.svg"
    case "btnSettings":
      return isDark ? "assets/cog-outline-dark.svg" : "assets/cog-outline.svg"
    case "btnDebugFieldData":
      return isDark ? "assets/cog-outline-dark.svg" : "assets/cog-outline.svg"
    case "btnDarkTheme":
      return isDark ? "assets/theme-light-dark-dark.svg" : "assets/theme-light-dark.svg"
    default:
      return isDark ? "assets/cog-outline-dark.svg" : "assets/cog-outline.svg"
  }
}

function OnGetPressed(control: RibbonControl) {
  if (control.Id === "btnDarkTheme") {
    return getThemeMode() === "dark"
  }
  return false
}

function OnGetVisible(_control: RibbonControl) {
  if (_control.Id === "btnDebugFieldData") {
    return import.meta.env.DEV
  }
  return true
}

// Expose callbacks for WPS ribbon runtime
export type RibbonCallbacks = {
  OnAddinLoad: typeof OnAddinLoad
  OnGetLabel: typeof OnGetLabel
  OnAction: typeof OnAction
  OnGetImage: typeof OnGetImage
  OnGetPressed: typeof OnGetPressed
  OnGetVisible: typeof OnGetVisible
}

export const ribbonCallbacks = {
  OnAddinLoad,
  OnGetLabel,
  OnAction,
  OnGetImage,
  OnGetPressed,
  OnGetVisible,
} satisfies RibbonCallbacks

export function registerRibbonCallbacks(scope: Window = window) {
  Object.assign(scope as Window & Partial<RibbonCallbacks>, ribbonCallbacks)
}

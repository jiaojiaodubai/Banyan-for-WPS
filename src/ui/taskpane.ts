import {
  findNextChapterBreak,
  findPreviousChapterBreak,
  getUpdateRange,
} from "../modules/chapter-break"
import { getLocalPreference } from "../modules/preference"
import type { IntextCitation, NoteCitation } from "../typings/style"
import { getConfig } from "../utils/config"
import { isIntextCitation, isNoteCitation, readFieldData } from "../utils/field"
import { request } from "../utils/http"
import { useI10n } from "../utils/i10n"

function getAssetUrl(filename: string): string {
  const path = `assets/${filename}`
  if (import.meta.env.DEV) {
    return new URL(`/${path}`, window.location.href).toString()
  }
  return new URL(`../${path}`, window.location.href).toString()
}

const shelfIcon = getAssetUrl("library-shelves.svg")
const itemBulletIcon = getAssetUrl("file-document-outline.svg")
const chevronLeftIcon = getAssetUrl("chevron-double-left.svg")
const chevronRightIcon = getAssetUrl("chevron-double-right.svg")
const chevronCollapsedIcon = getAssetUrl("chevron-right.svg")
const chevronExpandedIcon = getAssetUrl("chevron-down.svg")

const LABELS_ZH = {
  title: "当前章节引注",
  titleWithCount: "当前章节引注（{count}）",
  reload: "刷新",
  prevChapter: "上一章",
  nextChapter: "下一章",
  openInLibrary: "在Zotero文库中查看",
  pageLabel: "第 {page} 页",
  empty: "当前章节未检测到引注。",
  noStyle: "尚未设置引注样式，请先在设置中选择样式。",
  loadError: "读取引注列表失败：{message}",
  jumpError: "跳转引注失败：{message}",
  showLibraryError: "在Zotero文库中定位失败：{message}",
  untitled: "（无文本预览）",
}

const LABELS_EN = {
  title: "Citations in Current Section",
  titleWithCount: "Citations in Current Section ({count})",
  reload: "Reload",
  prevChapter: "Previous Chapter",
  nextChapter: "Next Chapter",
  openInLibrary: "Show in Zotero Library",
  pageLabel: "Page {page}",
  empty: "No citation found in current section.",
  noStyle: "Citation style is not set. Please configure style in Preferences first.",
  loadError: "Failed to load citation list: {message}",
  jumpError: "Failed to jump to citation: {message}",
  showLibraryError: "Failed to show item in Zotero library: {message}",
  untitled: "(No preview text)",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { taskpane: LABELS_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { taskpane: LABELS_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { taskpane: LABELS_EN },
})

type CitationData = IntextCitation | NoteCitation
type ThemeMode = "light" | "dark"
type CitationItem = {
  id: number
  uri: string
  title: string
}

type LoadResult = {
  entries: CitationEntry[]
  message?: string
}

type CitationEntry = {
  field: Wps.Field
  data: CitationData
  preview: string
  page: number
  items: CitationItem[]
}

const THEME_SYNC_CHANNEL_NAME = "banyan-theme-sync"
const TASKPANE_SYNC_CHANNEL_NAME = "banyan-taskpane-sync"
const TASKPANE_SYNC_STORAGE_KEY = "banyan_taskpane_sync_event"

const nodes = {
  paneTitle: document.getElementById("pane-title") as HTMLDivElement,
  reloadBtn: document.getElementById("reload-btn") as HTMLButtonElement,
  prevChapterBtn: document.getElementById("prev-chapter-btn") as HTMLButtonElement,
  prevChapterIcon: document.getElementById("prev-chapter-icon") as HTMLImageElement,
  nextChapterBtn: document.getElementById("next-chapter-btn") as HTMLButtonElement,
  nextChapterIcon: document.getElementById("next-chapter-icon") as HTMLImageElement,
  citationList: document.getElementById("citation-list") as HTMLUListElement,
  citationMessage: document.getElementById("citation-message") as HTMLDivElement,
}

let currentThemeMode: ThemeMode | null = null
let themeChannel: BroadcastChannel | null = null
let taskpaneSyncChannel: BroadcastChannel | null = null
let activeGroupElement: HTMLLIElement | null = null
let activeItemElement: HTMLLIElement | null = null
let selectionRefreshTimer: number | null = null
let lastSelectionKey: string | null = null
let selectionPreservationDepth = 0

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message
  }
  return String(error)
}

function renderStaticText() {
  renderTitleWithCount(0)
  nodes.reloadBtn.textContent = t("taskpane.reload")
  nodes.prevChapterBtn.title = t("taskpane.prevChapter")
  nodes.prevChapterBtn.ariaLabel = t("taskpane.prevChapter")
  nodes.nextChapterBtn.title = t("taskpane.nextChapter")
  nodes.nextChapterBtn.ariaLabel = t("taskpane.nextChapter")

  nodes.prevChapterIcon.src = chevronLeftIcon
  nodes.nextChapterIcon.src = chevronRightIcon
}

function renderTitleWithCount(count: number) {
  nodes.paneTitle.textContent = t("taskpane.titleWithCount", { count })
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

function shouldRefreshTaskpane(payload: unknown): boolean {
  if (typeof payload !== "string") {
    return false
  }
  try {
    const parsed = JSON.parse(payload) as { type?: unknown }
    return parsed.type === "citations-refreshed"
  }
  catch {
    return false
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

function setupTaskpaneSync(): void {
  try {
    taskpaneSyncChannel = new BroadcastChannel(TASKPANE_SYNC_CHANNEL_NAME)
    taskpaneSyncChannel.addEventListener("message", (event) => {
      if (shouldRefreshTaskpane(event.data)) {
        void refreshPane()
      }
    })
  }
  catch {
    taskpaneSyncChannel = null
  }

  window.addEventListener("storage", (event) => {
    if (event.key !== TASKPANE_SYNC_STORAGE_KEY) {
      return
    }
    if (shouldRefreshTaskpane(event.newValue)) {
      void refreshPane()
    }
  })
}

function scheduleRefreshForSelectionChange(): void {
  if (selectionPreservationDepth > 0) {
    return
  }
  const selectionKey = getSelectionKey()
  if (selectionKey === null || selectionKey === lastSelectionKey) {
    return
  }
  lastSelectionKey = selectionKey

  if (selectionRefreshTimer !== null) {
    window.clearTimeout(selectionRefreshTimer)
  }
  selectionRefreshTimer = window.setTimeout(() => {
    selectionRefreshTimer = null
    void refreshPane()
  }, 120)
}

function setupSelectionSync(): void {
  rememberCurrentSelection()
  try {
    Application.ApiEvent.AddApiEventListener(
      "WindowSelectionChange",
      scheduleRefreshForSelectionChange,
    )
  }
  catch {
    // Ignore environments without WPS selection events.
  }
}

function disposeThemeSync(): void {
  if (themeChannel) {
    themeChannel.close()
    themeChannel = null
  }
  if (taskpaneSyncChannel) {
    taskpaneSyncChannel.close()
    taskpaneSyncChannel = null
  }
  if (selectionRefreshTimer !== null) {
    window.clearTimeout(selectionRefreshTimer)
    selectionRefreshTimer = null
  }
  try {
    Application.ApiEvent.RemoveApiEventListener(
      "WindowSelectionChange",
      scheduleRefreshForSelectionChange,
    )
  }
  catch {
    // Ignore environments without WPS selection events.
  }
}

async function withSelectionPreserved<T>(task: () => Promise<T> | T): Promise<T> {
  const originalRange = wps.Selection.Range.Duplicate
  selectionPreservationDepth += 1
  try {
    return await task()
  }
  finally {
    try {
      originalRange.Select()
      rememberCurrentSelection()
    }
    finally {
      selectionPreservationDepth -= 1
    }
  }
}

function getSelectionKey(): string | null {
  try {
    const range = wps.Selection.Range.Duplicate
    return `${wps.Selection.StoryType}:${range.Start}:${range.End}`
  }
  catch {
    return null
  }
}

function rememberCurrentSelection(): void {
  lastSelectionKey = getSelectionKey()
}

function isRangeStartInRange(container: Wps.Range, candidate: Wps.Range): boolean {
  return candidate.Start >= container.Start && candidate.Start < container.End
}

function getPreview(field: Wps.Field, data: CitationData): string {
  const text = field.Result.Text.replace(/[\r\n]+/g, " ").trim()
  if (text) {
    return text
  }
  return data.id || t("taskpane.untitled")
}

function extractCitationItems(data: CitationData): CitationItem[] {
  const items = new Map<number, CitationItem>()
  for (const cite of data.source.cites) {
    if (typeof cite.item.id === "number") {
      const existing = items.get(cite.item.id)
      if (existing) {
        continue
      }
      const title = typeof cite.item.title === "string" && cite.item.title.trim()
        ? cite.item.title
        : t("taskpane.untitled")
      items.set(cite.item.id, { id: cite.item.id, uri: cite.item.uri, title })
    }
  }
  return [...items.values()]
}

function collectIntextEntries(range: Wps.Range): CitationEntry[] {
  const entries: CitationEntry[] = []
  for (let i = 1; i <= range.Fields.Count; i += 1) {
    const field = range.Fields.Item(i)
    if (field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }
    if (!isRangeStartInRange(range, field.Result)) {
      continue
    }
    const data = readFieldData(field)
    if (!isIntextCitation(data)) {
      continue
    }
    const page = field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
    entries.push({
      field,
      data,
      preview: getPreview(field, data),
      page,
      items: extractCitationItems(data),
    })
  }
  return entries
}

function collectNoteEntries(range: Wps.Range): CitationEntry[] {
  const entries: CitationEntry[] = []
  for (let i = 1; i <= range.Footnotes.Count; i += 1) {
    const note = range.Footnotes.Item(i)
    if (!isRangeStartInRange(range, note.Reference)) {
      continue
    }
    if (note.Range.Fields.Count === 0) {
      continue
    }
    const field = note.Range.Fields.Item(1)
    if (field.Type !== wps.Enum.wdFieldAddin) {
      continue
    }
    const data = readFieldData(field)
    if (!isNoteCitation(data)) {
      continue
    }
    const page = field.Result.Information(wps.Enum.wdActiveEndPageNumber) as number
    entries.push({
      field,
      data,
      preview: getPreview(field, data),
      page,
      items: extractCitationItems(data),
    })
  }
  return entries
}

async function loadCitationEntries(): Promise<LoadResult> {
  const pref = getLocalPreference()
  if (!pref) {
    return {
      entries: [],
      message: t("taskpane.noStyle"),
    }
  }
  const style = pref.style

  const entries = await withSelectionPreserved(() => {
    const range = getUpdateRange()
    if (style.citationType === "note-citation") {
      return collectNoteEntries(range)
    }
    return collectIntextEntries(range)
  })
  return { entries }
}

function clearList() {
  activeGroupElement = null
  activeItemElement = null
  nodes.citationList.hidden = false
  nodes.citationMessage.hidden = true
  nodes.citationMessage.textContent = ""
  while (nodes.citationList.firstChild) {
    nodes.citationList.removeChild(nodes.citationList.firstChild)
  }
}

function renderListMessage(message: string) {
  nodes.citationList.hidden = true
  nodes.citationMessage.hidden = false
  nodes.citationMessage.textContent = message
}

function updateChapterNavigationAvailability() {
  void withSelectionPreserved(() => {
    nodes.prevChapterBtn.disabled = findPreviousChapterBreak() === null
    nodes.nextChapterBtn.disabled = findNextChapterBreak() === null
  })
}

function goToPreviousChapter() {
  const prevBreak = findPreviousChapterBreak()
  if (!prevBreak) {
    return
  }
  const range = wps.ActiveDocument.Range().Duplicate
  range.SetRange(prevBreak.field.Result.Start, prevBreak.field.Result.Start)
  range.Select()
}

function goToNextChapter() {
  const nextBreak = findNextChapterBreak()
  if (!nextBreak) {
    return
  }
  const range = nextBreak.field.Result.Duplicate
  range.Collapse(wps.Enum.wdCollapseEnd)
  range.Select()
}

function createLibraryButton(item: CitationItem, onSelect?: () => void): HTMLButtonElement {
  const button = document.createElement("button")
  button.type = "button"
  button.tabIndex = -1
  button.className = "icon-button library-btn icon-library icon"
  button.title = t("taskpane.openInLibrary")
  button.ariaLabel = t("taskpane.openInLibrary")

  const icon = document.createElement("img")
  icon.className = "icon"
  icon.src = shelfIcon
  icon.alt = ""
  icon.setAttribute("aria-hidden", "true")
  button.appendChild(icon)

  button.addEventListener("click", (event) => {
    event.stopPropagation()
    onSelect?.()
    void showInLibrary(item)
  })

  return button
}

function setGroupExpanded(toggle: HTMLButtonElement, itemsList: HTMLUListElement, expanded: boolean) {
  itemsList.hidden = !expanded
  toggle.classList.toggle("is-expanded", expanded)
  const icon = toggle.querySelector("img") as HTMLImageElement | null
  if (icon) {
    icon.src = expanded ? chevronExpandedIcon : chevronCollapsedIcon
  }
}

function createToggleButton(itemsList: HTMLUListElement): HTMLButtonElement {
  const button = document.createElement("button")
  button.type = "button"
  button.tabIndex = -1
  button.className = "icon-button icon-twisty icon"
  const icon = document.createElement("img")
  icon.className = "icon"
  icon.src = chevronCollapsedIcon
  icon.alt = ""
  icon.setAttribute("aria-hidden", "true")
  button.appendChild(icon)

  button.addEventListener("click", () => {
    const shouldExpand = itemsList.hidden !== false
    setGroupExpanded(button, itemsList, shouldExpand)
  })
  return button
}

function createCitationItems(entry: CitationEntry): HTMLUListElement {
  const list = document.createElement("ul")
  list.className = "citation-item-list"
  list.hidden = true
  for (const item of entry.items) {
    const row = document.createElement("li")
    row.className = "citation-item"

    const itemIcon = document.createElement("img")
    itemIcon.className = "icon-document icon"
    itemIcon.src = itemBulletIcon
    itemIcon.alt = ""
    itemIcon.setAttribute("aria-hidden", "true")

    const title = document.createElement("span")
    title.className = "citation-item-title"
    title.textContent = item.title
    title.title = item.title

    const markSelected = () => {
      if (activeItemElement === row) {
        return
      }
      activeGroupElement?.classList.remove("is-active")
      activeGroupElement = null
      activeItemElement?.classList.remove("is-selected")
      activeItemElement = row
      activeItemElement.classList.add("is-selected")
    }

    row.addEventListener("click", (event) => {
      event.stopPropagation()
      markSelected()
    })

    row.appendChild(itemIcon)
    row.appendChild(title)
    row.appendChild(createLibraryButton(item, markSelected))
    list.appendChild(row)
  }
  return list
}

function createJumpButton(entry: CitationEntry): HTMLButtonElement {
  const button = document.createElement("button")
  button.type = "button"
  button.tabIndex = -1
  button.className = "citation-preview"
  button.textContent = entry.preview
  button.title = entry.preview

  button.addEventListener("click", () => {
    try {
      entry.field.Select()
    }
    catch (error) {
      alert(t("taskpane.jumpError", { message: getErrorMessage(error) }))
    }
  })

  return button
}

function keepChildOutOfKeyboardFocus(child: HTMLElement, group: HTMLLIElement) {
  child.tabIndex = -1
  child.addEventListener("focus", () => {
    group.focus()
  })
}

function moveCitationFocus(currentGroup: HTMLLIElement, direction: "up" | "down") {
  const groups = Array.from(nodes.citationList.querySelectorAll(":scope > li")) as HTMLLIElement[]
  const currentIndex = groups.indexOf(currentGroup)
  if (currentIndex < 0) {
    return
  }
  const nextIndex = direction === "up" ? currentIndex - 1 : currentIndex + 1
  if (nextIndex < 0 || nextIndex >= groups.length) {
    return
  }
  groups[nextIndex].focus()
}

function markActiveGroup(group: HTMLLIElement) {
  if (activeGroupElement === group) {
    return
  }
  activeItemElement?.classList.remove("is-selected")
  activeItemElement = null
  activeGroupElement?.classList.remove("is-active")
  activeGroupElement = group
  activeGroupElement.classList.add("is-active")
}

function createCitationGroup(entry: CitationEntry): HTMLLIElement {
  const group = document.createElement("li")
  group.tabIndex = 0

  const citationLine = document.createElement("div")
  citationLine.className = "citation-line"

  const itemsList = createCitationItems(entry)
  const toggle = createToggleButton(itemsList)

  const page = document.createElement("span")
  page.className = "citation-page"
  page.textContent = t("taskpane.pageLabel", { page: entry.page })

  const jumpButton = createJumpButton(entry)
  jumpButton.addEventListener("click", () => {
    markActiveGroup(group)
  })

  group.addEventListener("focus", () => {
    markActiveGroup(group)
  })

  group.addEventListener("keydown", (event) => {
    switch (event.key) {
      case "ArrowLeft": {
        event.preventDefault()
        setGroupExpanded(toggle, itemsList, false)
        break
      }
      case "ArrowRight": {
        event.preventDefault()
        setGroupExpanded(toggle, itemsList, true)
        break
      }
      case "ArrowUp": {
        event.preventDefault()
        moveCitationFocus(group, "up")
        break
      }
      case "ArrowDown": {
        event.preventDefault()
        moveCitationFocus(group, "down")
        break
      }
      case "Enter":
      case " ": {
        event.preventDefault()
        markActiveGroup(group)
        try {
          entry.field.Select()
        }
        catch (error) {
          alert(t("taskpane.jumpError", { message: getErrorMessage(error) }))
        }
        break
      }
      default:
        break
    }
  })

  keepChildOutOfKeyboardFocus(toggle, group)
  keepChildOutOfKeyboardFocus(jumpButton, group)
  itemsList.addEventListener("focusin", (event) => {
    const target = event.target as HTMLElement | null
    if (!target) {
      return
    }
    if (target.classList.contains("library-btn")) {
      group.focus()
    }
  })

  citationLine.appendChild(toggle)
  citationLine.appendChild(jumpButton)
  citationLine.appendChild(page)

  group.appendChild(citationLine)
  group.appendChild(itemsList)

  group.addEventListener("click", (event) => {
    const target = event.target as Node
    if (target === toggle || toggle.contains(target)) {
      return
    }
    if (target === itemsList || itemsList.contains(target)) {
      return
    }
    markActiveGroup(group)
  })

  return group
}

function renderCitationList(entries: CitationEntry[], message?: string) {
  clearList()
  renderTitleWithCount(entries.length)

  if (message) {
    renderListMessage(message)
    return
  }

  if (entries.length === 0) {
    renderListMessage(t("taskpane.empty"))
    return
  }

  for (const entry of entries) {
    nodes.citationList.appendChild(createCitationGroup(entry))
  }
}

async function showInLibrary(item: CitationItem) {
  try {
    await request("showInLibrary", { uri: item.uri })
  }
  catch (error) {
    alert(t("taskpane.showLibraryError", { message: getErrorMessage(error) }))
  }
}

async function refreshPane() {
  nodes.reloadBtn.disabled = true
  try {
    const { entries, message } = await loadCitationEntries()
    renderCitationList(entries, message)
  }
  catch (error) {
    renderCitationList([], t("taskpane.loadError", {
      message: getErrorMessage(error),
    }))
  }
  finally {
    nodes.reloadBtn.disabled = false
    updateChapterNavigationAvailability()
    rememberCurrentSelection()
  }
}

function bindEvents() {
  nodes.reloadBtn.addEventListener("click", () => {
    void refreshPane()
  })
  nodes.prevChapterBtn.addEventListener("click", () => {
    goToPreviousChapter()
    void refreshPane()
  })
  nodes.nextChapterBtn.addEventListener("click", () => {
    goToNextChapter()
    void refreshPane()
  })
  window.addEventListener("focus", () => {
    syncThemeMode()
    void refreshPane()
  })
  window.addEventListener("beforeunload", () => {
    disposeThemeSync()
  })
}

function initialize() {
  setupThemeSync()
  setupTaskpaneSync()
  setupSelectionSync()
  renderStaticText()
  bindEvents()
  void refreshPane()
}

initialize()

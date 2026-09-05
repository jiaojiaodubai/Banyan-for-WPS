import { useI10n } from "../utils/i10n"
import { withProgress } from "../utils/progress"
import { joinPath } from "../utils/path"
import {
  type BanyanFieldData,
  applyRichTextLinksToRange,
  getBibliographyBookmarkName,
  isBibliographyEntry,
  isBibliographyTitle,
  isIntextCitation,
  isNoteCitation,
  readFieldData,
  renderRange,
} from "../utils/field"
import { isChapterBreak } from "./chapter-break"
import { refreshAll } from "./refresh"
import { logWarn } from "../utils/log"

type FileSystemApi = {
  copyFileSync?: (sourcePath: string, targetPath: string, flags?: number) => boolean | void
  existsSync?: (path: string) => boolean
  unlinkSync?: (path: string) => boolean | void
}

type UnlinkStats = {
  total: number
  success: number
  failed: number
}

const FINALIZE_MESSAGE_ZH = {
  confirmTitle: "定稿确认",
  confirmBody:
    "执行“定稿”会将 Banyan 域代码替换为当前显示结果，且不可撤销。\n\n"
    + "系统会先自动备份当前文档，再继续执行。\n\n"
    + "是否继续？",
  progressReason: "正在执行文档定稿...",
  noFilePath: "当前文档尚未保存，无法在原路径创建备份。请先保存文档后再执行定稿。",
  copyApiMissing: "当前环境不支持文件复制接口，无法创建备份。",
  deleteApiMissing: "当前环境不支持文件删除接口，无法安全覆盖同名备份。",
  success:
    "定稿完成。\n\n"
    + "备份文件：{backupPath}\n"
    + "已处理域：{success}\n"
    + "处理失败：{failed}",
  failed:
    "定稿失败：{message}",
  partial:
    "定稿部分完成。\n\n"
    + "备份文件：{backupPath}\n"
    + "已处理域：{success}\n"
    + "处理失败：{failed}",
}

const FINALIZE_MESSAGE_EN = {
  confirmTitle: "Finalize Confirmation",
  confirmBody:
    "Finalize will replace Banyan fields with their rendered results and this action cannot be undone.\n\n"
    + "A backup will be created in the same folder before processing.\n\n"
    + "Do you want to continue?",
  progressReason: "Finalizing document...",
  noFilePath: "The current document is not saved. Please save it first so a backup can be created.",
  copyApiMissing: "File copy API is not available in the current environment. Cannot create backup.",
  deleteApiMissing: "File delete API is not available in the current environment. Cannot safely replace an existing backup.",
  success:
    "Finalize completed.\n\n"
    + "Backup: {backupPath}\n"
    + "Fields processed: {success}\n"
    + "Failed: {failed}",
  failed: "Finalize failed: {message}",
  partial:
    "Finalize partially completed.\n\n"
    + "Backup: {backupPath}\n"
    + "Fields processed: {success}\n"
    + "Failed: {failed}",
}

const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { finalize: FINALIZE_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDChineseSingapore]: { finalize: FINALIZE_MESSAGE_ZH },
  [wps.Enum.msoLanguageIDEnglishUS]: { finalize: FINALIZE_MESSAGE_EN },
})

type DocPathInfo = {
  fullPath: string
  dirPath: string
  fileName: string
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message
  }
  return String(error)
}

function splitDocPath(fullPath: string): DocPathInfo {
  const slashIndex = Math.max(fullPath.lastIndexOf("\\"), fullPath.lastIndexOf("/"))
  if (slashIndex <= 0 || slashIndex >= fullPath.length - 1) {
    throw new Error(t("finalize.noFilePath"))
  }
  return {
    fullPath,
    dirPath: fullPath.slice(0, slashIndex),
    fileName: fullPath.slice(slashIndex + 1),
  }
}

function splitFileName(fileName: string): { name: string; ext: string } {
  const dotIndex = fileName.lastIndexOf(".")
  if (dotIndex <= 0 || dotIndex === fileName.length - 1) {
    return {
      name: fileName,
      ext: "",
    }
  }
  return {
    name: fileName.slice(0, dotIndex),
    ext: fileName.slice(dotIndex),
  }
}

function formatBackupTimestamp(date: Date = new Date()): string {
  const month = String(date.getMonth() + 1).padStart(2, "0")
  const day = String(date.getDate()).padStart(2, "0")
  const hour = String(date.getHours()).padStart(2, "0")
  const minute = String(date.getMinutes()).padStart(2, "0")
  const second = String(date.getSeconds()).padStart(2, "0")
  return `${month}-${day} ${hour}-${minute}-${second}`
}

function resolveCurrentDocumentPath(): DocPathInfo {
  const doc = wps.ActiveDocument as Wps.Document & {
    FullName?: string
    Path?: string
    Name?: string
  }
  const fullName = typeof doc.FullName === "string" ? doc.FullName.trim() : ""
  if (fullName && (fullName.includes("\\") || fullName.includes("/"))) {
    return splitDocPath(fullName)
  }

  const dirPath = typeof doc.Path === "string" ? doc.Path.trim() : ""
  const fileName = typeof doc.Name === "string" ? doc.Name.trim() : ""
  if (!dirPath || !fileName) {
    throw new Error(t("finalize.noFilePath"))
  }

  return {
    fullPath: joinPath(dirPath, fileName),
    dirPath,
    fileName,
  }
}

function createBackupPath(info: DocPathInfo): string {
  const { name, ext } = splitFileName(info.fileName)
  const stamp = formatBackupTimestamp()
  return joinPath(info.dirPath, `${name}-${stamp}${ext}`)
}

function copyFileOrThrow(fs: FileSystemApi, sourcePath: string, targetPath: string): void {
  if (typeof fs.copyFileSync !== "function") {
    throw new Error(t("finalize.copyApiMissing"))
  }
  const result = fs.copyFileSync(sourcePath, targetPath)
  if (result === false) {
    throw new Error(t("finalize.copyApiMissing"))
  }
}

function fileExists(fs: FileSystemApi, filePath: string): boolean {
  if (typeof fs.existsSync !== "function") {
    throw new Error(t("finalize.copyApiMissing"))
  }
  return fs.existsSync(filePath)
}

function deleteFileOrThrow(fs: FileSystemApi, filePath: string): void {
  if (typeof fs.unlinkSync !== "function") {
    throw new Error(t("finalize.deleteApiMissing"))
  }
  const result = fs.unlinkSync(filePath)
  if (result === false) {
    throw new Error(t("finalize.deleteApiMissing"))
  }
}

function cleanupTempFile(fs: FileSystemApi, filePath: string): void {
  if (!fileExists(fs, filePath)) {
    return
  }
  try {
    deleteFileOrThrow(fs, filePath)
  }
  catch (error) {
    logWarn("Finalize", "Failed to clean up temp backup file.", error)
  }
}

function replaceBackupFile(fs: FileSystemApi, sourcePath: string, backupPath: string): void {
  if (!fileExists(fs, backupPath)) {
    copyFileOrThrow(fs, sourcePath, backupPath)
    return
  }

  const tempExistingPath = `${backupPath}.banyan-existing`
  const tempNewPath = `${backupPath}.banyan-new`
  cleanupTempFile(fs, tempExistingPath)
  cleanupTempFile(fs, tempNewPath)

  copyFileOrThrow(fs, backupPath, tempExistingPath)
  try {
    copyFileOrThrow(fs, sourcePath, tempNewPath)
    deleteFileOrThrow(fs, backupPath)
    try {
      copyFileOrThrow(fs, tempNewPath, backupPath)
    }
    catch (error) {
      copyFileOrThrow(fs, tempExistingPath, backupPath)
      throw error
    }
  }
  finally {
    cleanupTempFile(fs, tempNewPath)
    cleanupTempFile(fs, tempExistingPath)
  }
}

function backupCurrentDocument(): string {
  const fs = wps.FileSystem as FileSystemApi

  const info = resolveCurrentDocumentPath()
  const backupPath = createBackupPath(info)
  replaceBackupFile(fs, info.fullPath, backupPath)
  return backupPath
}

function getFinalizeAction(field: Wps.Field): "unlink" | "delete" | null {
  if (field.Type !== wps.Enum.wdFieldAddin) {
    return null
  }
  const data = readFieldData(field)
  if (!data) {
    return null
  }
  if (isIntextCitation(data) || isNoteCitation(data) || isBibliographyTitle(data) || isBibliographyEntry(data)) {
    return "unlink"
  }
  if (isChapterBreak(data)) {
    return "delete"
  }
  return null
}

function unlinkField(field: Wps.Field): boolean {
  try {
    const data = readFieldData(field)
    const resultRange = field.Result.Duplicate
    if (field.Locked) {
      field.Locked = false
    }
    field.Unlink()
    restoreUnlinkedResult(resultRange, data)
    return true
  }
  catch (error) {
    logWarn("Finalize", "Failed to unlink field.", error)
    return false
  }
}

function restoreUnlinkedResult(range: Wps.Range, data: BanyanFieldData | null): void {
  if (!data) {
    return
  }
  try {
    if (isBibliographyEntry(data)) {
      restoreBookmarkToRange(range, getBibliographyBookmarkName(data.id))
    }
    if (isIntextCitation(data) || isNoteCitation(data)) {
      // WPS leaves nested HYPERLINK fields inside an unlinked ADDIN result.
      // Rewriting citation text keeps the Result range but removes that stale field structure.
      renderRange(range, data.content)
      return
    }
    applyRichTextLinksToRange(range, data.content)
  }
  catch (error) {
    logWarn("Finalize", "Failed to restore finalized field result.", error)
  }
}

function restoreBookmarkToRange(range: Wps.Range, bookmarkName: string): void {
  if (!bookmarkName || bookmarkName.trim() === "") {
    return
  }

  try {
    const bookmarks = wps.ActiveDocument.Bookmarks
    for (let i = 1; i <= bookmarks.Count; i += 1) {
      if (bookmarks.Item(i).Name === bookmarkName) {
        bookmarks.Item(i).Delete()
        break
      }
    }
    bookmarks.Add(bookmarkName, range)
  }
  catch (error) {
    logWarn("Finalize", `Failed to restore bookmark \"${bookmarkName}\" after unlinking field.`, error)
  }
}

function deleteField(field: Wps.Field): boolean {
  try {
    if (field.Locked) {
      field.Locked = false
    }
    field.Delete()
    return true
  }
  catch (error) {
    logWarn("Finalize", "Failed to delete field.", error)
    return false
  }
}

function unlinkBanyanFields(): UnlinkStats {
  const stats: UnlinkStats = {
    total: 0,
    success: 0,
    failed: 0,
  }

  const mainFields = wps.ActiveDocument.Content.Fields
  for (let i = mainFields.Count; i >= 1; i -= 1) {
    const field = mainFields.Item(i)
    const action = getFinalizeAction(field)
    if (!action) {
      continue
    }
    stats.total += 1
    if ((action === "unlink" ? unlinkField(field) : deleteField(field))) {
      stats.success += 1
    }
    else {
      stats.failed += 1
    }
  }

  const footnotes = wps.ActiveDocument.Footnotes
  for (let i = footnotes.Count; i >= 1; i -= 1) {
    const footnote = footnotes.Item(i)
    const fields = footnote.Range.Fields
    for (let j = fields.Count; j >= 1; j -= 1) {
      const field = fields.Item(j)
      const action = getFinalizeAction(field)
      if (!action) {
        continue
      }
      stats.total += 1
      if ((action === "unlink" ? unlinkField(field) : deleteField(field))) {
        stats.success += 1
      }
      else {
        stats.failed += 1
      }
    }
  }

  return stats
}

export async function onFinalizeEvent(): Promise<void> {
  const confirmed = confirm(`${t("finalize.confirmTitle")}\n\n${t("finalize.confirmBody")}`)
  if (!confirmed) {
    return
  }

  await withProgress(t("finalize.progressReason"), async () => {
    try {
      const backupPath = backupCurrentDocument()

      // 定稿前先进行一次全量刷新，确保引注和参考文献表是最新的
      try {
        await refreshAll()
      }
      catch (error) {
        logWarn("Finalize", "Full refresh before finalize failed, continuing with current data.", error)
      }

      const stats = unlinkBanyanFields()

      if (stats.failed > 0) {
        alert(t("finalize.partial", {
          backupPath,
          success: stats.success,
          failed: stats.failed,
        }))
        return
      }

      alert(t("finalize.success", {
        backupPath,
        success: stats.success,
        failed: stats.failed,
      }))
    }
    catch (error) {
      alert(t("finalize.failed", {
        message: getErrorMessage(error),
      }))
    }
  })
}

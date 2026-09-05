const LOG_PREFIX = "[Banyan]"
const DESKTOP_LOG_FILE = "Banyan-for-WPS-debug.log"

let pendingDesktopLines: string[] = []
let desktopFlushScheduled = false
let desktopLogInitialized = false

function formatMessage(scope: string, message: string): string {
  return `${LOG_PREFIX}[${scope}] ${message}`
}

function nowIsoTimestamp(): string {
  return new Date().toISOString()
}

function stringifyArg(arg: unknown): string {
  if (arg instanceof Error) {
    return `${arg.name}: ${arg.message}`
  }
  if (typeof arg === "string") {
    return arg
  }
  try {
    return JSON.stringify(arg)
  }
  catch {
    return String(arg)
  }
}

function getDesktopLogPath(): string | null {
  try {
    return `${wps.Env.GetDesktopPath()}\\${DESKTOP_LOG_FILE}`
  }
  catch {
    return null
  }
}

function writeDesktopLog(content: string): void {
  const filePath = getDesktopLogPath()
  if (!filePath) {
    return
  }

  const fileSystem = wps.FileSystem as {
    AppendFile?: (path: string, data: string) => boolean
    WriteFile?: (path: string, data: string) => boolean
    Exists?: (path: string) => boolean
  }

  try {
    if (desktopLogInitialized && typeof fileSystem.AppendFile === "function") {
      fileSystem.AppendFile(filePath, content)
      return
    }

    const exists = typeof fileSystem.Exists === "function"
      ? fileSystem.Exists(filePath)
      : false

    if (exists && typeof fileSystem.AppendFile === "function") {
      fileSystem.AppendFile(filePath, content)
      desktopLogInitialized = true
      return
    }

    if (typeof fileSystem.WriteFile === "function") {
      fileSystem.WriteFile(filePath, content)
      desktopLogInitialized = true
    }
  }
  catch {
    // Do not throw if debug log persistence fails.
  }
}

function flushDesktopLogQueue(): void {
  desktopFlushScheduled = false
  if (pendingDesktopLines.length === 0) {
    return
  }
  const content = pendingDesktopLines.join("")
  pendingDesktopLines = []
  writeDesktopLog(content)
}

function scheduleDesktopFlush(): void {
  if (desktopFlushScheduled) {
    return
  }
  desktopFlushScheduled = true
  setTimeout(flushDesktopLogQueue, 0)
}

function enqueueDesktopLog(line: string): void {
  pendingDesktopLines.push(line)
  if (pendingDesktopLines.length >= 16) {
    flushDesktopLogQueue()
    return
  }
  scheduleDesktopFlush()
}

function logDesktopLine(level: "INFO" | "WARN" | "ERROR" | "TRACE", scope: string, message: string, args: unknown[]): void {
  const serializedArgs = args.map(arg => stringifyArg(arg)).join(" ")
  const line = `${nowIsoTimestamp()} [${level}] ${formatMessage(scope, message)}${serializedArgs ? ` ${serializedArgs}` : ""}\n`
  enqueueDesktopLog(line)
}

export function logInfo(scope: string, message: string, ...args: unknown[]): void {
  console.info(formatMessage(scope, message), ...args)
  logDesktopLine("INFO", scope, message, args)
}

export function logWarn(scope: string, message: string, ...args: unknown[]): void {
  console.warn(formatMessage(scope, message), ...args)
  logDesktopLine("WARN", scope, message, args)
}

export function logError(scope: string, message: string, ...args: unknown[]): void {
  console.error(formatMessage(scope, message), ...args)
  logDesktopLine("ERROR", scope, message, args)
}

export function logTrace(scope: string, message: string, ...args: unknown[]): void {
  console.debug(formatMessage(scope, message), ...args)
  logDesktopLine("TRACE", scope, message, args)
}

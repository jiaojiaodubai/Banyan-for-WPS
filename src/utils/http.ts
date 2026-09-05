import type { HttpPath, ResponsePayload, RouteTable } from "../typings/http"
import { logError } from "./log"

const OFFLINE_ALERT_MESSAGE =
  "无法连接到 Zotero 本地服务器，请启动 Zotero 客户端并确保 Banyan 插件已启用后重试。"

export const DEFAULT_ZOTERO_PORT = "23119"
export const FALLBACK_ZOTERO_PORT = "23124"

let activeZoteroPort = DEFAULT_ZOTERO_PORT

function getCandidatePorts(): string[] {
  const alternate = activeZoteroPort === DEFAULT_ZOTERO_PORT
    ? FALLBACK_ZOTERO_PORT
    : DEFAULT_ZOTERO_PORT
  return [activeZoteroPort, alternate]
}

export function getZoteroApiBaseUrl(): string {
  return `http://localhost:${activeZoteroPort}/api/`
}

export function getBanyanApiBaseUrl(): string {
  return `http://localhost:${activeZoteroPort}/banyan`
}

export function getDocumentId(): string {
  const doc = Application.ActiveDocument
  if (!doc) {
    return "__no_document__"
  }
  return doc.FullName || doc.Name || "__unnamed__"
}

export async function request<T extends HttpPath>(
  path: T,
  data: RouteTable[T]["req"],
): Promise<RouteTable[T]["res"] | null> {
  let lastError: unknown
  for (const port of getCandidatePorts()) {
    try {
      const result = await requestAtPort(port, path, data)
      activeZoteroPort = port
      return result
    }
    catch (error) {
      lastError = error
      if (!(error instanceof Error && error.message.includes("Failed to fetch"))) {
        logError("HTTP", `Request failed for path: ${path}.`, error)
        throw error
      }
    }
  }

  alert(OFFLINE_ALERT_MESSAGE)
  logError("HTTP", `Request failed for path: ${path}.`, lastError)
  throw lastError
}

async function requestAtPort<T extends HttpPath>(
  port: string,
  path: T,
  data: RouteTable[T]["req"],
): Promise<RouteTable[T]["res"] | null> {
  const url = `http://localhost:${port}/banyan/${path}`
  const requestInit: RequestInit = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Zotero-Allowed-Request": "1",
      "X-Banyan-Client": "Banyan for WPS",
    },
  }
  if (data !== undefined) {
    requestInit.body = JSON.stringify(data)
  }

  const response = await fetch(url, requestInit)
  if (!response.ok) {
    throw new Error(`Network response was not ok: ${response.statusText}`)
  }

  const result = (await response.json()) as ResponsePayload<T>
  if (result.ok === true) return result.data
  if (result.error.code === "cancelled") return null
  alert(`HTTP error: [${result.error.code}] ${result.error.message}`)
  throw result.error
}

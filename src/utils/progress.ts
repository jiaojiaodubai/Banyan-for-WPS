import { request } from "./http"
import { logWarn } from "./log"

/**
 * 用进度条包装一个异步操作
 * @param reason - The reason to display in the progress bar
 * @param operation - The async operation to execute
 * @returns The result of the operation
 */
export async function withProgress<T>(
  reason: string,
  operation: () => Promise<T>
): Promise<T> {
  try {
    await request("progress", {
      action: "open",
      reason,
    })
  }
  catch (error) {
    logWarn("Progress", "Failed to open progress bar.", error)
  }

  try {
    return await operation()
  }
  finally {
    try {
      await request("progress", {
        action: "close",
      })
    }
    catch (error) {
      logWarn("Progress", "Failed to close progress bar.", error)
    }
  }
}

let batchDepth = 0
let previousScreenUpdating = true
let customUndoStarted = false

/**
 * 把一次文档更新合并为一个重绘/撤销事务。
 * WPS 与 Word 暴露同样的 UndoRecord 接口，但旧版本实现不可靠，因此撤销分组
 * 只做尽力而为，重绘状态则必定恢复。
 */
export function beginBatchUpdate(name = "Banyan Update"): void {
  if (batchDepth === 0) {
    previousScreenUpdating = Application.ScreenUpdating
    try {
      Application.ScreenUpdating = false
    }
    catch {
      // 宿主拒绝修改重绘状态时也继续执行。
    }

    try {
      const undoRecord = Application.UndoRecord
      if (undoRecord && !undoRecord.IsRecordingCustomRecord) {
        undoRecord.StartCustomRecord(name)
        customUndoStarted = true
      }
    }
    catch {
      customUndoStarted = false
    }
  }
  batchDepth += 1
}

export function endBatchUpdate(): void {
  if (batchDepth <= 0) return
  batchDepth -= 1
  if (batchDepth > 0) return

  try {
    if (customUndoStarted) {
      Application.UndoRecord.EndCustomRecord()
    }
  }
  catch {
    // 文档更新本身已经完成。
  }
  finally {
    customUndoStarted = false
    try {
      Application.ScreenUpdating = previousScreenUpdating
    }
    catch {
      // 宿主拒绝赋值时已无其它可安全做的事。
    }
  }
}

export async function withBatchUpdate<T>(name: string, callback: () => T | Promise<T>): Promise<T> {
  beginBatchUpdate(name)
  try {
    return await callback()
  }
  finally {
    endBatchUpdate()
  }
}

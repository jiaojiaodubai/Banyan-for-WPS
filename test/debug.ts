/**
 * Ribbon “运行调试代码” 按钮的入口：人工临时调试专用，按钮点击后会以
 * `withOperationLock` 串行执行，避免和刷新等操作撞车。
 */

/** 统计当前文档里各前缀的域数量（只读，不修改文档）。 */
function countFieldsByPrefix(prefix: string): number {
  const fields = Application.ActiveDocument?.Fields
  if (!fields) return 0

  let count = 0
  for (let index = 1; index <= fields.Count; index += 1) {
    const code = fields.Item(index).Code?.Text ?? ""
    if (code.startsWith(prefix)) count += 1
  }
  return count
}

/** 按钮入口：收集调试信息，控制台回显一份、弹窗一份。 */
export function runDebugEntry(): void {
  const document = Application.ActiveDocument
  const lines = [
    `WPS ${Application.Version}`,
    `当前文档：${document?.Name ?? "（无）"}`,
    `域总数：${document?.Fields?.Count ?? 0}`,
    `BANYAN_CITATION 域：${countFieldsByPrefix("BANYAN_CITATION")}`,
    `BANYAN_BIBLIOGRAPHY 域：${countFieldsByPrefix("BANYAN_BIBLIOGRAPHY")}`,
  ]
  for (const line of lines) console.info("[debug]", line)
  alert(`Banyan 调试信息\n\n${lines.join("\n")}`)
}

export default runDebugEntry

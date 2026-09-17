import { runBibliographyTests } from "./bibliography.test"
import { runCitationTests } from "./citation.test"
import { runFieldTests } from "./field.test"
import { TestContext, type PerformanceMetric } from "./framework"
import { runHttpTests } from "./http.test"
import { runRefreshTests } from "./refresh.test"

const LOG_FILE_NAME = "Banyan-for-WPS-performance.log"

/** 临时文档的标记：清理上次异常中断（刷新页面、杀进程、命令被中止）留下的残留。 */
const FIXTURE_PROPERTY = "BANYAN_TEST_FIXTURE"
const FIXTURE_TEXT = "Banyan performance test. "

/**
 * 同一时刻只允许一轮测试。
 *
 * 调试桥的页面侧给每条命令加了执行上限（见桥的 commandTimeoutMs）：超时只会把命令
 * 判为失败，桥会接着执行下一条命令。若此时又跑一次测试，就会
 * 两轮并发建文档、吃内存。所以入口自己把关：正在跑就拒掉。
 */
let running = false

type FixtureProperties = {
  Item: (name: string) => { Value?: unknown }
  Add: (name: string, link: boolean, type: number, value: string) => unknown
}

function getFixtureProperties(document: Wps.Document): FixtureProperties | undefined {
  return (document as unknown as { CustomDocumentProperties?: FixtureProperties }).CustomDocumentProperties
}

/** 给临时文档打标记，让后续运行能认出异常中断后的残留。 */
function markFixture(document: Wps.Document): void {
  try {
    getFixtureProperties(document)?.Add(FIXTURE_PROPERTY, false, 4, "1")
  }
  catch (error) {
    // 打标记失败不致命：清理时还会按“未保存 + 开头是我们写入的标记文本”兜底
    console.warn("[Banyan Test] 临时文档标记失败：", String(error))
  }
}

function isFixture(document: Wps.Document): boolean {
  try {
    if (getFixtureProperties(document)?.Item(FIXTURE_PROPERTY)?.Value) return true
  }
  catch {
    // 没这个属性：走文本兜底
  }
  try {
    return document.Path === "" && (document.Content.Text ?? "").startsWith(FIXTURE_TEXT)
  }
  catch {
    return false
  }
}

/** 关掉上次异常中断留下的临时文档（只动带标记的，不碰用户文档），返回清理数量。 */
function closeStaleFixtures(): number {
  const documents = Application.Documents
  let closed = 0
  for (let index = documents.Count; index >= 1; index -= 1) {
    try {
      const document = documents.Item(index)
      if (!isFixture(document)) continue
      document.Close(wps.Enum.wdDoNotSaveChanges)
      closed += 1
    }
    catch (error) {
      // 单个文档失败不影响其余
      console.warn("[Banyan Test] 清理临时文档失败：", String(error))
    }
  }
  return closed
}

function formatMs(milliseconds: number): string {
  if (milliseconds >= 1000) return `${(milliseconds / 1000).toFixed(3)} s`
  return `${milliseconds.toFixed(2)} ms`
}

function formatMetric(metric: PerformanceMetric): string {
  const attributes = [
    metric.size === undefined ? "" : `size=${metric.size}`,
    metric.mode ? `mode=${metric.mode}` : "",
    metric.iterations === undefined ? "" : `iterations=${metric.iterations}`,
  ].filter(Boolean).join(" ")
  const perItemCount = metric.size ?? metric.iterations
  const perItem = perItemCount ? ` perItem=${formatMs(metric.totalMs / perItemCount)}` : ""
  const statistic = metric.statistic && metric.statistic !== "single" ? ` statistic=${metric.statistic}` : ""
  return `[TIME] ${metric.module}.${metric.name}${attributes ? ` ${attributes}` : ""} total=${formatMs(metric.totalMs)}${perItem}${statistic}`
}

function buildReport(context: TestContext, totalMs: number): string {
  const passed = context.results.filter(result => result.status === "pass").length
  const failed = context.results.length - passed
  const lines = [
    "Banyan for WPS test and performance report",
    "========================================================================",
    `[INFO] generatedAt=${new Date().toISOString()}`,
    `[INFO] WPS=${Application.Version}; userAgent=${navigator.userAgent}`,
    "[INFO] scope=WPS object model, citation/field/bibliography paths, local JSON and Banyan backend hello endpoint",
    "[INFO] benchmarkSizes=10,50,100; fieldRepetitions=4; styleIterations=20",
    "[INFO] stagedMetrics=field text/style/rich-text/link stages; citation source/content decisions; bibliography incremental/delete strategies; host-option diagnostic uses size=10; selected comparisons use 3-sample median",
    "",
    "[TESTS]",
  ]

  for (const result of context.results) {
    const status = result.status === "pass" ? "PASS" : "FAIL"
    lines.push(`[${status}] ${result.module}.${result.name} duration=${formatMs(result.durationMs)}${result.error ? ` error=${result.error}` : ""}`)
  }

  lines.push("", "[PERFORMANCE]")
  for (const metric of context.metrics) lines.push(formatMetric(metric))
  if (context.notes.length > 0) {
    lines.push("", "[NOTES]", ...context.notes.map(note => `[INFO] ${note}`))
  }
  lines.push(
    "",
    "[SUMMARY]",
    `tests=${context.results.length} passed=${passed} failed=${failed} metrics=${context.metrics.length} total=${formatMs(totalMs)}`,
    "",
  )
  return lines.join("\n")
}

function writeReportToDesktop(report: string): string {
  const path = `${wps.Env.GetDesktopPath()}\\${LOG_FILE_NAME}`
  const wrote = wps.FileSystem.WriteFile(path, report)
  if (!wrote) throw new Error(`WPS FileSystem.WriteFile returned false for ${path}`)
  return path
}

export async function runPerformanceTests(): Promise<void> {
  if (running) {
    console.warn("[Banyan Test] 上一轮测试还没结束，忽略本次调用（并发跑会重复建文档、吃内存）")
    return
  }
  running = true
  try {
    await runPerformanceTestsOnce()
  }
  finally {
    running = false
  }
}

async function runPerformanceTestsOnce(): Promise<void> {
  const context = new TestContext()
  const startedAt = performance.now()
  const originalDocument = Application.ActiveDocument
  const originalRange = originalDocument ? wps.Selection.Range.Duplicate : null
  const originalScreenUpdating = Application.ScreenUpdating
  let testDocument: Wps.Document | null = null

  try {
    const stale = closeStaleFixtures()
    if (stale > 0) console.info(`[Banyan Test] 已清理 ${stale} 个上次中断留下的临时文档`)

    testDocument = Application.Documents.Add()
    markFixture(testDocument)
    testDocument.Activate()
    testDocument.Content.Text = FIXTURE_TEXT

    const suites = [runHttpTests, runFieldTests, runCitationTests, runRefreshTests, runBibliographyTests]
    for (const runSuite of suites) {
      // 让出事件循环：一轮同步宿主调用占满主线程时，长轮询与主线程 tick 都发不出去
      // （桥的信标在 Web Worker 里，不受影响）。留一拍让它们有机会上报。
      await new Promise(resolve => setTimeout(resolve, 50))
      try {
        await runSuite(context)
      }
      catch (error) {
        context.fail("runner", runSuite.name, error)
      }
    }
  }
  catch (error) {
    context.fail("runner", "temporary document lifecycle", error)
  }
  finally {
    Application.ScreenUpdating = originalScreenUpdating
    try {
      testDocument?.Close(wps.Enum.wdDoNotSaveChanges)
    }
    catch (error) {
      context.fail("runner", "close temporary document", error)
    }
    try {
      originalDocument?.Activate()
      originalRange?.Select()
    }
    catch (error) {
      context.fail("runner", "restore original document", error)
    }
  }

  const totalMs = performance.now() - startedAt
  const report = buildReport(context, totalMs)
  let logPath: string | null = null
  try {
    logPath = writeReportToDesktop(report)
  }
  catch (error) {
    console.error("[Banyan Test] Failed to write performance report.", error)
  }
  console.info(report)
  console.info(logPath ? `[Banyan Test] 报告已写入 ${logPath}` : "[Banyan Test] 报告写入桌面失败，详见上方错误。")
}

export default runPerformanceTests

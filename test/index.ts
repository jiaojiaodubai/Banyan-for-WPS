import { runBibliographyTests } from "./bibliography.test"
import { runFieldTests } from "./field.test"
import { TestContext, type PerformanceMetric } from "./framework"
import { runHttpTests } from "./http.test"
import { runRefreshTests } from "./refresh.test"

const LOG_FILE_NAME = "Banyan-for-WPS-performance.log"

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
  return `[TIME] ${metric.module}.${metric.name}${attributes ? ` ${attributes}` : ""} total=${formatMs(metric.totalMs)}${perItem}`
}

function buildReport(context: TestContext, totalMs: number): string {
  const passed = context.results.filter(result => result.status === "pass").length
  const failed = context.results.length - passed
  const lines = [
    "Banyan for WPS test and performance report",
    "========================================================================",
    `[INFO] generatedAt=${new Date().toISOString()}`,
    `[INFO] WPS=${Application.Version}; userAgent=${navigator.userAgent}`,
    "[INFO] scope=WPS object model, local JSON and Banyan backend hello endpoint",
    "[INFO] benchmarkSizes=10,50,100; fieldRepetitions=4; styleIterations=20",
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

function showSummary(context: TestContext, totalMs: number, logPath: string | null): void {
  const passed = context.results.filter(result => result.status === "pass").length
  const failed = context.results.length - passed
  const logLine = logPath ? `日志：${logPath}` : "日志写入失败，请查看开发者控制台。"
  alert(
    "Banyan 性能测试完成\n\n"
    + `测试：${context.results.length}\n`
    + `通过：${passed}\n`
    + `失败：${failed}\n`
    + `性能指标：${context.metrics.length}\n`
    + `总耗时：${formatMs(totalMs)}\n\n`
    + logLine,
  )
}

export async function runPerformanceTests(): Promise<void> {
  const context = new TestContext()
  const startedAt = performance.now()
  const originalDocument = Application.ActiveDocument
  const originalRange = originalDocument ? wps.Selection.Range.Duplicate : null
  const originalScreenUpdating = Application.ScreenUpdating
  let testDocument: Wps.Document | null = null

  try {
    testDocument = Application.Documents.Add()
    testDocument.Activate()
    testDocument.Content.Text = "Banyan performance test. "

    const suites = [runHttpTests, runFieldTests, runRefreshTests, runBibliographyTests]
    for (const runSuite of suites) {
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
  showSummary(context, totalMs, logPath)
}

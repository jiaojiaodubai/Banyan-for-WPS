export type TestStatus = "pass" | "fail"

export type TestResult = {
  module: string
  name: string
  status: TestStatus
  durationMs: number
  error?: string
}

export type PerformanceMetric = {
  module: string
  name: string
  totalMs: number
  size?: number
  mode?: string
  iterations?: number
}

type TestCallback = () => void | Promise<void>
type MeasureCallback = () => void | Promise<void>

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return `${error.name}: ${error.message}`
  }
  return String(error)
}

function describe(value: unknown): string {
  if (typeof value === "string") return JSON.stringify(value)
  try {
    return JSON.stringify(value)
  }
  catch {
    return String(value)
  }
}

export class AssertionError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "AssertionError"
  }
}

type Assert = {
  ok(value: unknown, message?: string): asserts value
  equal<T>(actual: T, expected: T, message?: string): void
  deepEqual(actual: unknown, expected: unknown, message?: string): void
}

export const assert: Assert = {
  ok(value: unknown, message = "Expected value to be truthy"): asserts value {
    if (!value) throw new AssertionError(message)
  },

  equal<T>(actual: T, expected: T, message?: string): void {
    if (Object.is(actual, expected)) return
    throw new AssertionError(message ?? `Expected ${describe(expected)}, received ${describe(actual)}`)
  },

  deepEqual(actual: unknown, expected: unknown, message?: string): void {
    if (JSON.stringify(actual) === JSON.stringify(expected)) return
    throw new AssertionError(message ?? `Expected ${describe(expected)}, received ${describe(actual)}`)
  },
}

export class TestContext {
  readonly results: TestResult[] = []
  readonly metrics: PerformanceMetric[] = []
  readonly notes: string[] = []

  async test(module: string, name: string, callback: TestCallback): Promise<void> {
    const startedAt = performance.now()
    try {
      await callback()
      this.results.push({
        module,
        name,
        status: "pass",
        durationMs: performance.now() - startedAt,
      })
    }
    catch (error) {
      this.results.push({
        module,
        name,
        status: "fail",
        durationMs: performance.now() - startedAt,
        error: getErrorMessage(error),
      })
    }
  }

  async measure(
    module: string,
    name: string,
    callback: MeasureCallback,
    options: Pick<PerformanceMetric, "size" | "mode" | "iterations"> = {},
  ): Promise<number> {
    const startedAt = performance.now()
    await callback()
    const totalMs = performance.now() - startedAt
    this.metrics.push({ module, name, totalMs, ...options })
    return totalMs
  }

  fail(module: string, name: string, error: unknown): void {
    this.results.push({
      module,
      name,
      status: "fail",
      durationMs: 0,
      error: getErrorMessage(error),
    })
  }

  note(message: string): void {
    this.notes.push(message)
  }
}

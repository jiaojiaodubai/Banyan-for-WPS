import { getBanyanApiBaseUrl, getZoteroApiBaseUrl, request } from "../src/utils/http"
import { assert, type TestContext } from "./framework"

const MODULE = "http"
const WARM_REQUEST_COUNT = 4

export async function runHttpTests(context: TestContext): Promise<void> {
  await context.test(MODULE, "API base URLs", () => {
    assert.ok(getZoteroApiBaseUrl().endsWith("/api/"))
    assert.ok(getBanyanApiBaseUrl().endsWith("/banyan"))
  })

  await context.test(MODULE, "backend hello endpoint", async () => {
    let response: string | null = null
    await context.measure(MODULE, "hello first request", async () => {
      response = await request("hello", undefined as never)
    }, { iterations: 1, mode: "production-path" })
    assert.equal(response, "Hello from Banyan server!")

    await context.measure(MODULE, "hello warm requests", async () => {
      for (let i = 0; i < WARM_REQUEST_COUNT; i += 1) {
        const warmResponse = await request("hello", undefined as never)
        assert.equal(warmResponse, "Hello from Banyan server!")
      }
    }, { iterations: WARM_REQUEST_COUNT, mode: "production-path" })
  })
}

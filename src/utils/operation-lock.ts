import { getDocumentId } from "./http"

const operationLocks = new Map<string, boolean>()

export function isOperationPending(docId?: string): boolean {
  const key = docId ?? getDocumentId()
  return operationLocks.get(key) === true
}

function startOperation(docId?: string): void {
  const key = docId ?? getDocumentId()
  operationLocks.set(key, true)
}

function endOperation(docId?: string): void {
  const key = docId ?? getDocumentId()
  operationLocks.delete(key)
}

export async function withOperationLock<T>(
  fn: () => Promise<T>,
  docId?: string,
  _operationName: string = "operation"
): Promise<T> {
  const key = docId ?? getDocumentId()

  if (isOperationPending(key)) {
    return Promise.resolve(undefined as T)
  }

  startOperation(key)
  try {
    return await fn()
  }
  finally {
    endOperation(key)
  }
}

const TASKPANE_ID_KEY = "taskpane_id"
const TASKPANE_SYNC_CHANNEL_NAME = "banyan-taskpane-sync"
const TASKPANE_SYNC_STORAGE_KEY = "banyan_taskpane_sync_event"

const taskpaneUrl = import.meta.env.DEV
  ? new URL("/src/ui/taskpane.html", window.location.href).toString()
  : new URL("ui/taskpane.html", window.location.href).toString()

function getStoredTaskpaneId(): number | undefined {
  const taskpaneIdRaw = Application.PluginStorage.getItem(TASKPANE_ID_KEY)
  if (!taskpaneIdRaw) {
    return undefined
  }
  const taskpaneId = Number(taskpaneIdRaw)
  return Number.isNaN(taskpaneId) ? undefined : taskpaneId
}

function getTaskpane(): WPS.CustomTaskpane | undefined {
  const taskpaneId = getStoredTaskpaneId()
  if (typeof taskpaneId !== "number") {
    return undefined
  }
  try {
    const pane = Application.GetTaskPane(taskpaneId)
    return pane ?? undefined
  }
  catch {
    return undefined
  }
}

function createTaskpane(): WPS.CustomTaskpane {
  const pane = Application.CreateTaskPane(taskpaneUrl)
  Application.PluginStorage.setItem(TASKPANE_ID_KEY, pane.ID)
  return pane
}

export function toggleTaskpaneVisibility() {
  const pane = getTaskpane() ?? createTaskpane()
  pane.Visible = !pane.Visible
}

export function isTaskpaneVisible(): boolean {
  const pane = getTaskpane()
  return !!pane?.Visible
}

export function notifyTaskpaneCitationsRefreshed() {
  if (!isTaskpaneVisible()) {
    return
  }

  const payload = JSON.stringify({
    type: "citations-refreshed",
    ts: Date.now(),
  })

  try {
    const channel = new BroadcastChannel(TASKPANE_SYNC_CHANNEL_NAME)
    channel.postMessage(payload)
    channel.close()
  }
  catch {
    // Ignore environments without BroadcastChannel support.
  }

  try {
    window.localStorage.setItem(TASKPANE_SYNC_STORAGE_KEY, payload)
  }
  catch {
    // Ignore storage write failures.
  }
}

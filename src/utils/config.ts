import type { ConfigKey, ConfigMap, PersistOption } from "../typings/config"

export async function loadConfig(): Promise<void> {
  await Promise.resolve()
}

export function getConfig<K extends ConfigKey>(key: K, fallback: ConfigMap[K]): ConfigMap[K] {
  const value = Application.PluginStorage.getItem(key)
  return typeof value === "string" ? value as ConfigMap[K] : fallback
}

export function setConfig<K extends ConfigKey>(
  key: K,
  value: ConfigMap[K],
  options?: PersistOption,
): void {
  void options
  Application.PluginStorage.setItem(key, String(value))
}

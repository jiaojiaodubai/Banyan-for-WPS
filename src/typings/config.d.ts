export interface ConfigMap {
  ribbon_theme_mode: "light" | "dark"
}

export type ConfigKey = keyof ConfigMap

export type PersistOption = {
  persistent?: boolean
}

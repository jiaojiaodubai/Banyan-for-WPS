/// <reference types="vite/client" />

declare const __BUILD_TIME__: string

// i10n types
type LocaleParams = Record<string, unknown>
interface LocaleStructure {
  [key: string]: LocaleStructure | string
}
type LocaleData = Record<number, LocaleStructure>

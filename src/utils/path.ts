export function joinPath(...segments: string[]): string {
  const filtered = segments.filter(segment => segment && segment.length > 0)
  if (filtered.length === 0) {
    return ""
  }
  // 始终使用 "/" 拼接，再通过 WPS 官方 API 转换为原生分隔符
  return wps.FileSystem.toNativeSeparators(filtered.join("/"))
}

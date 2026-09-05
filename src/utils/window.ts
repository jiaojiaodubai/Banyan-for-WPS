export function openDialog(
  url: string,
  title?: string,
  width?: number,
  height?: number,
  isModal: boolean = true,
) {
  const { protocol, href, origin } = window.location
  const base = protocol === "file:" ? new URL(".", href).toString() : origin
  const dialogUrl = new URL(url, base).toString()
  window.Application.ShowDialog(dialogUrl, title, width, height, isModal)
}

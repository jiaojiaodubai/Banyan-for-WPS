/**
 * WPS Add-in Installer for Zotero Plugin
 *
 * This module provides install/uninstall functions for deploying the WPS add-in
 * to the local WPS jsaddons directory from within a Zotero plugin.
 *
 * Usage in Zotero plugin:
 *   import { installWpsAddin, uninstallWpsAddin } from './install'
 *   await installWpsAddin({ sourceDir: '...', name: 'Banyan', version: '1.0.0' })
 */

interface InstallOptions {
  sourceDir: string
  name: string
  version: string
  addonType?: string
}

interface Result {
  ok: boolean
  error?: string
}

/**
 * Get WPS jsaddons directory path based on platform
 */
function getAddonPath(): string {
  const os = Services.appinfo.OS
  const home = PathUtils.profileDir.replace(/\/Profiles\/.*$/, '')

  if (os === 'WINNT') {
    const appdata = Services.env.get('APPDATA')
    return PathUtils.join(appdata, 'kingsoft', 'wps', 'jsaddons')
  } else if (os === 'Darwin') {
    return PathUtils.join(
      PathUtils.profileDir.replace(/\/Library\/.*$/, ''),
      'Library',
      'Containers',
      'com.kingsoft.wpsoffice.mac',
      'Data',
      '.kingsoft',
      'wps',
      'jsaddons'
    )
  } else {
    return PathUtils.join(home, '.local', 'share', 'Kingsoft', 'wps', 'jsaddons')
  }
}

/**
 * Recursively copy directory
 */
async function copyDirectory(src: string, dest: string): Promise<void> {
  await IOUtils.makeDirectory(dest, { ignoreExisting: true })

  const entries = await IOUtils.getChildren(src)
  for (const entry of entries) {
    const stat = await IOUtils.stat(entry)
    const basename = PathUtils.filename(entry)
    const destPath = PathUtils.join(dest, basename)

    if (stat.type === 'directory') {
      await copyDirectory(entry, destPath)
    } else {
      await IOUtils.copy(entry, destPath)
    }
  }
}

/**
 * Parse XML string to DOM
 */
function parseXML(xmlString: string): Document {
  const parser = new DOMParser()
  return parser.parseFromString(xmlString, 'text/xml')
}

/**
 * Serialize DOM to XML string
 */
function serializeXML(doc: Document): string {
  const serializer = new XMLSerializer()
  return serializer.serializeToString(doc)
}

/**
 * Ensure publish.xml exists with skeleton structure
 */
async function ensurePublishXml(publishXmlPath: string): Promise<void> {
  const exists = await IOUtils.exists(publishXmlPath)
  if (!exists) {
    const skeleton = '<?xml version="1.0" encoding="UTF-8"?>\n<jsplugins>\n</jsplugins>\n'
    await IOUtils.writeUTF8(publishXmlPath, skeleton)
  }
}

/**
 * Add or update jsplugin entry in publish.xml
 */
async function registerAddin(
  publishXmlPath: string,
  name: string,
  version: string,
  addonType: string
): Promise<void> {
  await ensurePublishXml(publishXmlPath)

  const xmlString = await IOUtils.readUTF8(publishXmlPath)
  const doc = parseXML(xmlString)
  const root = doc.documentElement

  if (root.nodeName !== 'jsplugins') {
    throw new Error('Invalid publish.xml: root element must be <jsplugins>')
  }

  const existingNodes = Array.from(root.getElementsByTagName('jsplugin'))
  for (const node of existingNodes) {
    if (node.getAttribute('name') === name) {
      root.removeChild(node)
    }
  }

  const newNode = doc.createElement('jsplugin')
  newNode.setAttribute('name', name)
  newNode.setAttribute('type', addonType)
  newNode.setAttribute('url', `${name}_${version}`)
  newNode.setAttribute('version', version)
  newNode.setAttribute('enable', 'enable_dev')
  newNode.setAttribute('install', 'null')
  newNode.setAttribute('customDomain', '')

  root.appendChild(doc.createTextNode('\n  '))
  root.appendChild(newNode)
  root.appendChild(doc.createTextNode('\n'))

  const newXmlString = serializeXML(doc)
  await IOUtils.writeUTF8(publishXmlPath, newXmlString)
}

/**
 * Remove jsplugin entry from publish.xml
 */
async function unregisterAddin(publishXmlPath: string, name: string): Promise<void> {
  const exists = await IOUtils.exists(publishXmlPath)
  if (!exists) {
    return
  }

  const xmlString = await IOUtils.readUTF8(publishXmlPath)
  const doc = parseXML(xmlString)
  const root = doc.documentElement

  if (root.nodeName !== 'jsplugins') {
    return
  }

  const nodes = Array.from(root.getElementsByTagName('jsplugin'))
  for (const node of nodes) {
    if (node.getAttribute('name') === name) {
      const prevSibling = node.previousSibling
      if (prevSibling && prevSibling.nodeType === Node.TEXT_NODE) {
        root.removeChild(prevSibling)
      }
      root.removeChild(node)
    }
  }

  const newXmlString = serializeXML(doc)
  await IOUtils.writeUTF8(publishXmlPath, newXmlString)
}

/**
 * Install WPS add-in to local jsaddons directory
 */
export async function installWpsAddin(opts: InstallOptions): Promise<Result> {
  const { sourceDir, name, version, addonType = 'wps' } = opts

  try {
    const addonPath = getAddonPath()
    const targetDir = PathUtils.join(addonPath, `${name}_${version}`)
    const publishXmlPath = PathUtils.join(addonPath, 'publish.xml')

    await IOUtils.makeDirectory(addonPath, { ignoreExisting: true })

    const targetExists = await IOUtils.exists(targetDir)
    if (targetExists) {
      await IOUtils.remove(targetDir, { recursive: true })
    }

    await copyDirectory(sourceDir, targetDir)

    await registerAddin(publishXmlPath, name, version, addonType)

    return { ok: true }
  } catch (err) {
    return { ok: false, error: String(err) }
  }
}

/**
 * Uninstall WPS add-in from local jsaddons directory
 */
export async function uninstallWpsAddin(opts: { name: string }): Promise<Result> {
  const { name } = opts

  try {
    const addonPath = getAddonPath()
    const publishXmlPath = PathUtils.join(addonPath, 'publish.xml')

    const addonPathExists = await IOUtils.exists(addonPath)
    if (!addonPathExists) {
      return { ok: true }
    }

    const entries = await IOUtils.getChildren(addonPath)
    for (const entry of entries) {
      const basename = PathUtils.filename(entry)
      if (basename.startsWith(`${name}_`)) {
        const stat = await IOUtils.stat(entry)
        if (stat.type === 'directory') {
          await IOUtils.remove(entry, { recursive: true })
        }
      }
    }

    await unregisterAddin(publishXmlPath, name)

    return { ok: true }
  } catch (err) {
    return { ok: false, error: String(err) }
  }
}

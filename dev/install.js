#!/usr/bin/env node
/**
 * WPS Add-in Installer (Node.js standalone version for testing)
 *
 * Usage:
 *   node dev/install.js install <sourceDir> <name> <version>
 *   node dev/install.js uninstall <name>
 *
 * Example:
 *   node dev/install.js install ./release/banyan-wps-addin_1.0.0 Banyan 1.0.0
 *   node dev/install.js uninstall Banyan
 */

import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'

/**
 * Get WPS jsaddons directory path based on platform
 */
function getAddonPath() {
  const platform = os.platform()
  const home = os.homedir()

  if (platform === 'win32') {
    const appdata = process.env.APPDATA
    if (!appdata) {
      throw new Error('APPDATA environment variable not found')
    }
    return path.join(appdata, 'kingsoft', 'wps', 'jsaddons')
  } else if (platform === 'darwin') {
    return path.join(
      home,
      'Library',
      'Containers',
      'com.kingsoft.wpsoffice.mac',
      'Data',
      '.kingsoft',
      'wps',
      'jsaddons'
    )
  } else {
    return path.join(home, '.local', 'share', 'Kingsoft', 'wps', 'jsaddons')
  }
}

/**
 * Recursively copy directory
 */
function copyDirectory(src, dest) {
  fs.mkdirSync(dest, { recursive: true })

  const entries = fs.readdirSync(src, { withFileTypes: true })
  for (const entry of entries) {
    const srcPath = path.join(src, entry.name)
    const destPath = path.join(dest, entry.name)

    if (entry.isDirectory()) {
      copyDirectory(srcPath, destPath)
    } else {
      fs.copyFileSync(srcPath, destPath)
    }
  }
}

/**
 * Ensure publish.xml exists with skeleton structure
 */
function ensurePublishXml(publishXmlPath) {
  if (!fs.existsSync(publishXmlPath)) {
    const skeleton = '<?xml version="1.0" encoding="UTF-8"?>\n<jsplugins>\n</jsplugins>\n'
    fs.writeFileSync(publishXmlPath, skeleton, 'utf-8')
  }
}

/**
 * Add or update jsplugin entry in publish.xml
 */
function registerAddin(publishXmlPath, name, version, addonType) {
  ensurePublishXml(publishXmlPath)

  const xmlString = fs.readFileSync(publishXmlPath, 'utf-8')
  const parser = new DOMParser()
  const doc = parser.parseFromString(xmlString, 'text/xml')
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

  const serializer = new XMLSerializer()
  const newXmlString = serializer.serializeToString(doc)
  fs.writeFileSync(publishXmlPath, newXmlString, 'utf-8')
}

/**
 * Remove jsplugin entry from publish.xml
 */
function unregisterAddin(publishXmlPath, name) {
  if (!fs.existsSync(publishXmlPath)) {
    return
  }

  const xmlString = fs.readFileSync(publishXmlPath, 'utf-8')
  const parser = new DOMParser()
  const doc = parser.parseFromString(xmlString, 'text/xml')
  const root = doc.documentElement

  if (root.nodeName !== 'jsplugins') {
    return
  }

  const nodes = Array.from(root.getElementsByTagName('jsplugin'))
  for (const node of nodes) {
    if (node.getAttribute('name') === name) {
      const prevSibling = node.previousSibling
      if (prevSibling && prevSibling.nodeType === 3) {
        root.removeChild(prevSibling)
      }
      root.removeChild(node)
    }
  }

  const serializer = new XMLSerializer()
  const newXmlString = serializer.serializeToString(doc)
  fs.writeFileSync(publishXmlPath, newXmlString, 'utf-8')
}

/**
 * Install WPS add-in
 */
function install(sourceDir, name, version, addonType = 'wps') {
  try {
    const addonPath = getAddonPath()
    const targetDir = path.join(addonPath, `${name}_${version}`)
    const publishXmlPath = path.join(addonPath, 'publish.xml')

    console.log(`Installing ${name} v${version}...`)
    console.log(`Source: ${sourceDir}`)
    console.log(`Target: ${targetDir}`)

    fs.mkdirSync(addonPath, { recursive: true })

    if (fs.existsSync(targetDir)) {
      console.log('Removing existing installation...')
      fs.rmSync(targetDir, { recursive: true, force: true })
    }

    console.log('Copying files...')
    copyDirectory(sourceDir, targetDir)

    console.log('Registering in publish.xml...')
    registerAddin(publishXmlPath, name, version, addonType)

    console.log('✓ Installation complete')
    return true
  } catch (err) {
    console.error('✗ Installation failed:', err.message)
    return false
  }
}

/**
 * Uninstall WPS add-in
 */
function uninstall(name) {
  try {
    const addonPath = getAddonPath()
    const publishXmlPath = path.join(addonPath, 'publish.xml')

    console.log(`Uninstalling ${name}...`)

    if (!fs.existsSync(addonPath)) {
      console.log('No jsaddons directory found')
      return true
    }

    const entries = fs.readdirSync(addonPath, { withFileTypes: true })
    for (const entry of entries) {
      if (entry.name.startsWith(`${name}_`) && entry.isDirectory()) {
        const fullPath = path.join(addonPath, entry.name)
        console.log(`Removing ${fullPath}...`)
        fs.rmSync(fullPath, { recursive: true, force: true })
      }
    }

    console.log('Unregistering from publish.xml...')
    unregisterAddin(publishXmlPath, name)

    console.log('✓ Uninstallation complete')
    return true
  } catch (err) {
    console.error('✗ Uninstallation failed:', err.message)
    return false
  }
}

// CLI
const args = process.argv.slice(2)
const command = args[0]

if (command === 'install') {
  const [, sourceDir, name, version] = args
  if (!sourceDir || !name || !version) {
    console.error('Usage: node install.js install <sourceDir> <name> <version>')
    process.exit(1)
  }
  const success = install(sourceDir, name, version)
  process.exit(success ? 0 : 1)
} else if (command === 'uninstall') {
  const [, name] = args
  if (!name) {
    console.error('Usage: node install.js uninstall <name>')
    process.exit(1)
  }
  const success = uninstall(name)
  process.exit(success ? 0 : 1)
} else {
  console.error('Usage:')
  console.error('  node install.js install <sourceDir> <name> <version>')
  console.error('  node install.js uninstall <name>')
  process.exit(1)
}

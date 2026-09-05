#!/usr/bin/env node
import { execSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import process from 'node:process'
import { fileURLToPath } from 'node:url'
import { DOMParser } from '@xmldom/xmldom'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const manifestXml = fs.readFileSync(path.join(__dirname, 'public/manifest.xml'), 'utf-8')
const doc = new DOMParser().parseFromString(manifestXml, 'text/xml')

const name = doc.getElementsByTagName('Name')[0].textContent
const version = doc.getElementsByTagName('ApiVersion')[0].textContent
const distDir = path.join(__dirname, 'dist')
const releaseDir = path.join(__dirname, 'release')
const targetDir = path.join(releaseDir, `${name}_${version}`)

console.log('Building WPS add-in...')
console.log(`Name: ${name}`)
console.log(`Version: ${version}`)

// Step 1: Run vite build
console.log('\n[1/3] Running vite build...')
try {
  execSync('npm run build:vite', { stdio: 'inherit', cwd: __dirname })
}
catch (err) {
  console.error('Vite build failed')
  console.error(err)
  process.exit(1)
}

// Step 2: Clean and prepare release directory
console.log('\n[2/3] Preparing release directory...')
if (fs.existsSync(releaseDir)) {
  fs.rmSync(releaseDir, { recursive: true, force: true })
}
fs.mkdirSync(releaseDir, { recursive: true })

// Step 3: Copy dist to release/<name>_<version>
console.log(`\n[3/3] Copying dist to release/${name}_${version}/...`)
function copyRecursive(src, dest) {
  const stat = fs.statSync(src)
  if (stat.isDirectory()) {
    fs.mkdirSync(dest, { recursive: true })
    for (const entry of fs.readdirSync(src)) {
      copyRecursive(path.join(src, entry), path.join(dest, entry))
    }
  }
  else {
    fs.copyFileSync(src, dest)
  }
}
copyRecursive(distDir, targetDir)

console.log(`\n✓ Build complete: ${targetDir}`)
console.log('\nNext steps:')
console.log('  1. Copy this directory to your Zotero plugin chrome/content/')
console.log('  2. Use dev/install.js or dev/install.ts to deploy to WPS jsaddons')

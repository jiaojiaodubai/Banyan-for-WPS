import { defineConfig } from 'vite'
import { resolve } from 'node:path'
import { fileURLToPath, URL } from 'node:url'
import {
    createClassicUiEntryPlugin,
    createMoveUiHtmlFilesPlugin,
    getUiHtmlInputs,
} from './dev/vite-ui-build.js'

const projectRoot = fileURLToPath(new URL('.', import.meta.url))
const uiDir = resolve(projectRoot, 'src/ui')
const classicUiEntries = {
    preference: resolve(uiDir, 'preference.ts'),
}
const uiHtmlInputs = getUiHtmlInputs(uiDir)

function isUiEntryName(name) {
    return name.startsWith('ui/')
}

function toUiOutputName(name) {
    return name.replace(/^ui\//, '')
}

// https://vitejs.dev/config/
export default defineConfig(() => {
    const buildTime = new Date().toLocaleString()
    const bundleOptions = {
        input: {
            main: resolve(projectRoot, 'index.html'),
            ...uiHtmlInputs,
        },
        output: {
            entryFileNames(chunkInfo) {
                if (isUiEntryName(chunkInfo.name)) {
                    return `ui/${toUiOutputName(chunkInfo.name)}-[hash].js`
                }
                return 'assets/[name]-[hash].js'
            },
            chunkFileNames: 'assets/[name]-[hash].js',
            assetFileNames: 'assets/[name]-[hash][extname]',
        },
    }

    return {
        base: './',
        define: {
            __BUILD_TIME__: JSON.stringify(buildTime),
        },
        plugins: [
            createClassicUiEntryPlugin({ buildTime, entries: classicUiEntries }),
            createMoveUiHtmlFilesPlugin(),
        ],
        build: {
            assetsInlineLimit: 0,
            rollupOptions: bundleOptions,
            rolldownOptions: bundleOptions,
        },
        server: {
            host: '0.0.0.0',
        },
    }
})

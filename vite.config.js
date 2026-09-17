import { defineConfig } from 'vite'
import { resolve } from 'node:path'
import { fileURLToPath, URL } from 'node:url'
import {
    createClassicUiEntryPlugin,
    createMoveUiHtmlFilesPlugin,
    getUiHtmlInputs,
} from './dev/vite-ui-build.js'
import { wpsDebugBridge } from 'wps-js-addin-debug-bridge'

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
            // 开发期调试桥（wps-js-addin-debug-bridge）：仅 serve 生效，生产构建不含它。
            // 性能测试要跑 3–4 分钟，超过桥默认的 120 s 页面执行上限，因此单独放宽；
            // 桥还会往项目目录写 .wps-bridge.json（端口/token 发现文件，已 gitignore）。
            wpsDebugBridge({ commandTimeoutMs: 600_000 }),
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

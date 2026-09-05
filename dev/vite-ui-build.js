import { build as esbuildBuild } from 'esbuild'
import fs from 'node:fs'
import { resolve } from 'node:path'

export function getUiHtmlInputs(uiDir) {
    return Object.fromEntries(
        fs.readdirSync(uiDir)
            .filter((name) => name.endsWith('.html'))
            .map((name) => {
                const entryName = `ui/${name.replace(/\.html$/, '')}`
                return [entryName, resolve(uiDir, name)]
            })
    )
}

export function createMoveUiHtmlFilesPlugin() {
    return {
        name: 'move-ui-html-files',
        enforce: 'post',
        generateBundle(_options, bundle) {
            for (const [fileName, asset] of Object.entries(bundle)) {
                if (fileName.startsWith('src/ui/') && fileName.endsWith('.html')) {
                    const newFileName = fileName.replace('src/ui/', 'ui/')
                    asset.fileName = newFileName
                    if (asset.source && typeof asset.source === 'string') {
                        asset.source = asset.source.replace(/\.\.\/\.\.\/assets\//g, '../assets/')
                        asset.source = asset.source.replace(/\.\.\/\.\.\/ui\//g, './')
                    }
                    delete bundle[fileName]
                    bundle[newFileName] = asset
                }
            }
        },
    }
}

async function bundleClassicUiEntry(name, entryPath, buildTime, isDev = false) {
    const result = await esbuildBuild({
        entryPoints: [entryPath],
        outfile: `${name}.js`,
        bundle: true,
        format: 'iife',
        platform: 'browser',
        target: 'es2018',
        write: false,
        minify: !isDev,
        sourcemap: isDev ? 'inline' : false,
        legalComments: 'none',
        define: {
            __BUILD_TIME__: JSON.stringify(buildTime),
            'import.meta.env.DEV': JSON.stringify(isDev),
            'import.meta.env.PROD': JSON.stringify(!isDev),
            'import.meta.env.MODE': JSON.stringify(isDev ? 'development' : 'production'),
        },
    })

    const output = result.outputFiles.find(file => file.path.endsWith('.js'))
    if (!output) {
        throw new Error(`Failed to bundle classic UI entry: ${entryPath}`)
    }

    return output.text
}

export function createClassicUiEntryPlugin({ buildTime, entries }) {
    return {
        name: 'classic-ui-entry',
        configureServer(server) {
            server.middlewares.use(async (req, res, next) => {
                const pathname = req.url?.split('?')[0]
                const match = Object.entries(entries).find(([name]) => pathname === `/src/ui/${name}.js`)
                if (!match) {
                    next()
                    return
                }

                try {
                    const [name, entryPath] = match
                    const code = await bundleClassicUiEntry(name, entryPath, buildTime, true)
                    res.statusCode = 200
                    res.setHeader('Content-Type', 'application/javascript; charset=utf-8')
                    res.end(code)
                }
                catch (error) {
                    next(error)
                }
            })
        },
        async generateBundle() {
            for (const [name, entryPath] of Object.entries(entries)) {
                const code = await bundleClassicUiEntry(name, entryPath, buildTime)
                this.emitFile({
                    type: 'asset',
                    fileName: `ui/${name}.js`,
                    source: code,
                })
            }
        },
    }
}
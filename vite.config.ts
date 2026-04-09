import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'
import path from 'path'

const httpsConfig = (() => {
  try {
    return {
      key: fs.readFileSync('./certs/key.pem'),
      cert: fs.readFileSync('./certs/cert.pem'),
    }
  } catch {
    return undefined
  }
})()

// mkcert で生成した証明書を使用（Windows の信頼済みルート CA に登録済み）
// 証明書の生成: npx mkcert -install && npx mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1
export default defineConfig({
  plugins: [
    react(),
    {
      // kuromoji 辞書ファイルを開発サーバーで /panel-for-word/dict/ として提供
      name: 'kuromoji-dict-serve',
      configureServer(server) {
        const dictDir = path.resolve('./node_modules/kuromoji/dict')
        server.middlewares.use('/panel-for-word/dict/', (req, res, next) => {
          const fileName = (req.url ?? '').replace(/^\/?/, '')
          if (!fileName) { next(); return }
          const filePath = path.join(dictDir, fileName)
          try {
            const buf = fs.readFileSync(filePath)
            res.setHeader('Content-Type', 'application/octet-stream')
            res.end(buf)
          } catch {
            next()
          }
        })
      },
    },
  ],
  base: '/panel-for-word/',
  resolve: {
    alias: {
      // kuromoji の DictionaryLoader が require("path") を使うため、
      // ブラウザ環境で動く最小シムに差し替える（path-browserify は https:// を壊すため不可）
      path: path.resolve('./src/utils/path-browser-shim.js'),
    },
  },
  server: {
    port: 3000,
    https: httpsConfig,
  },
  build: {
    outDir: 'dist',
  },
  optimizeDeps: {
    include: ['kuromoji'],
  },
})

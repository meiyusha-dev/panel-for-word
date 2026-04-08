import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'

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
  plugins: [react()],
  base: '/panel-for-word/',
  server: {
    port: 3000,
    https: httpsConfig,
  },
  build: {
    outDir: 'dist',
  },
})

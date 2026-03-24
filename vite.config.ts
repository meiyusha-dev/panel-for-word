import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'

// mkcert で生成した証明書を使用（Windows の信頼済みルート CA に登録済み）
// 証明書の生成: npx mkcert -install && npx mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1
export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    https: {
      key: fs.readFileSync('./certs/key.pem'),
      cert: fs.readFileSync('./certs/cert.pem'),
    },
  },
  build: {
    outDir: 'dist',
  },
})

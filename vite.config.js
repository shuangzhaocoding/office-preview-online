import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { fileProxyPlugin } from './vite.file-proxy.js'
import { pptxLazyRenderPlugin } from './vite.pptx-lazy.js'
import { pdfScalePlugin } from './vite.pdf-scale.js'

const root = fileURLToPath(new URL('.', import.meta.url))

export default defineConfig({
  define: {
    'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV || 'development'),
  },
  plugins: [vue(), fileProxyPlugin(), pptxLazyRenderPlugin(), pdfScalePlugin()],
  resolve: {
    alias: {
      '@vue-office/pptx': path.join(root, 'node_modules/@vue-office/pptx/lib/v3/vue-office-pptx.mjs'),
      '@vue-office/pdf': path.join(root, 'node_modules/@vue-office/pdf/lib/v3/vue-office-pdf.mjs'),
    },
  },
  optimizeDeps: {
    include: [
      'vue',
      'vue-demi',
      '@vue-office/docx',
      '@vue-office/excel',
    ],
    exclude: ['@vue-office/pptx', '@vue-office/pdf'],
  },
  server: {
    host: true,
    port: 5173,
    fs: {
      strict: false,
      allow: ['..'],
    },
  },
  preview: {
    host: true,
    port: 5173,
  },
})

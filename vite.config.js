import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { fileProxyPlugin } from './vite.file-proxy.js'
import { pptxLazyRenderPlugin } from './vite.pptx-lazy.js'
import { pdfScalePlugin } from './vite.pdf-scale.js'
import { excelPreviewPlugin } from './vite.excel-preview.js'

const root = fileURLToPath(new URL('.', import.meta.url))

export default defineConfig({
  define: {
    'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV || 'development'),
  },
  plugins: [vue(), fileProxyPlugin(), pptxLazyRenderPlugin(), pdfScalePlugin(), excelPreviewPlugin()],
  resolve: {
    alias: [
      {
        find: '@vue-office/pptx',
        replacement: path.join(root, 'node_modules/@vue-office/pptx/lib/v3/vue-office-pptx.mjs'),
      },
      {
        find: '@vue-office/pdf',
        replacement: path.join(root, 'node_modules/@vue-office/pdf/lib/v3/vue-office-pdf.mjs'),
      },
      {
        find: /^@vue-office\/excel$/,
        replacement: path.join(root, 'node_modules/@vue-office/excel/lib/v3/vue-office-excel.mjs'),
      },
    ],
  },
  optimizeDeps: {
    include: [
      'vue',
      'vue-demi',
      '@vue-office/docx',
    ],
    exclude: ['@vue-office/pptx', '@vue-office/pdf', '@vue-office/excel'],
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

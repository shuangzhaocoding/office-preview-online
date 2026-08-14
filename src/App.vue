<template>
  <div id="app-wrapper" :class="{ 'is-home': showHome }">
    <div v-if="showHome" class="home">
      <div class="home-card">
        <div class="home-marks" aria-hidden="true">
          <span class="home-mark" style="--mark: #2B579A">W</span>
          <span class="home-mark" style="--mark: #217346">X</span>
          <span class="home-mark" style="--mark: #E5252A">P</span>
          <span class="home-mark" style="--mark: #C43E1C">T</span>
        </div>
        <h1>Office 文件预览</h1>
        <p class="home-lead">
          在浏览器中直接预览 Word、Excel、PDF 与 PPT，无需安装办公软件。
        </p>

        <form class="home-form" @submit.prevent="submitHome">
          <label class="home-field">
            <span>文件地址</span>
            <input
              v-model="homeUrl"
              class="home-input"
              type="text"
              inputmode="url"
              autocomplete="off"
              placeholder="https://example.com/file.docx"
            />
          </label>

          <label class="home-field home-field-type">
            <span>文件类型</span>
            <select v-model="homeType" class="home-select">
              <option value="auto">自动推断</option>
              <option value="docx">Word（docx）</option>
              <option value="excel">Excel（xlsx / xls）</option>
              <option value="pdf">PDF</option>
              <option value="pptx">PPT（pptx）</option>
            </select>
          </label>

          <p v-if="homeTypeHint" class="home-hint-ok">{{ homeTypeHint }}</p>
          <p v-if="homeError" class="home-hint-err">{{ homeError }}</p>

          <button type="submit" class="home-submit">预览</button>
        </form>
      </div>
    </div>

    <div
      v-else
      class="preview-area"
      :class="{
        'has-page-toolbar': showPageToolbar,
        'is-pptx': previewType === 'pptx',
        'is-pdf': previewType === 'pdf',
        'is-docx': previewType === 'docx',
        'is-excel': previewType === 'excel',
        'excel-finding': previewType === 'excel' && excelFindOpen && excelFindHits.length > 0,
      }"
    >
      <div v-if="loading || hint" class="loading-overlay" :class="{ error: !!hint }">
        <div class="loading-circle">
          <svg class="progress-ring" :class="{ spin: phase === 'render' && !hint }" viewBox="0 0 120 120">
            <circle class="progress-track" cx="60" cy="60" r="54" />
            <circle
              class="progress-indicator"
              cx="60"
              cy="60"
              r="54"
              :stroke-dasharray="ringLength"
              :stroke-dashoffset="ringOffset"
            />
          </svg>
          <div class="loading-inner">
            <template v-if="hint">
              <div class="loading-text">{{ hint }}</div>
            </template>
            <template v-else>
              <div class="loading-elapsed">{{ elapsedText }}</div>
              <div class="loading-label">{{ statusLabel }}</div>
              <div v-if="downloadSizeText" class="loading-size">{{ downloadSizeText }}</div>
            </template>
          </div>
        </div>
      </div>

      <aside v-if="showPageToolbar" class="pptx-sidebar">
        <div ref="pptxSidebarList" class="pptx-sidebar-list">
          <button
            v-for="page in pageList"
            :key="page"
            type="button"
            class="pptx-thumb"
            :class="{ active: page === pageCurrent, ready: pageReady[page] }"
            :data-page="page"
            @click="goPage(page)"
          >
            <div class="pptx-thumb-frame">
              <div class="pptx-thumb-scale" :data-thumb-page="page"></div>
              <div v-if="!pageReady[page]" class="pptx-thumb-placeholder">{{ page }}</div>
            </div>
            <div class="pptx-thumb-label">{{ page }}</div>
          </button>
        </div>
      </aside>

      <vue-office-docx
        v-if="previewType === 'docx' && src"
        :src="src"
        style="height: 100%;"
        @rendered="renderedHandler"
        @error="errorHandler"
      />
      <vue-office-excel
        v-else-if="previewType === 'excel' && src"
        :src="src"
        :options="excelOptions"
        style="height: 100%;"
        @rendered="renderedHandler"
        @error="errorHandler"
      />
      <vue-office-pdf
        v-else-if="previewType === 'pdf' && src"
        :src="src"
        style="height: 100%;"
        @rendered="renderedHandler"
        @error="errorHandler"
      />
      <vue-office-pptx
        v-else-if="previewType === 'pptx' && src"
        :src="src"
        :options="pptxOptions"
        :request-options="requestOptions"
        style="height: 100%;"
        @rendered="renderedHandler"
        @error="errorHandler"
      />

      <div
        v-if="previewType === 'excel' && excelFindOpen"
        class="excel-find-bar"
        @mousedown.stop
        @click.stop
      >
        <input
          ref="excelFindInput"
          v-model="excelFindQuery"
          class="excel-find-input"
          type="search"
          placeholder="查找"
          autocomplete="off"
          @input="runExcelFind"
          @keydown="onExcelFindKeydown"
        />
        <span class="excel-find-count">{{ excelFindCountText }}</span>
        <button type="button" class="excel-find-btn" title="上一个" :disabled="!excelFindHits.length" @click="stepExcelFind(-1)">↑</button>
        <button type="button" class="excel-find-btn" title="下一个" :disabled="!excelFindHits.length" @click="stepExcelFind(1)">↓</button>
        <button type="button" class="excel-find-btn excel-find-close" title="关闭" @click="closeExcelFind">×</button>
      </div>

      <div v-if="excelCopyToast" class="excel-copy-toast" role="status">{{ excelCopyToast }}</div>

      <div v-if="showPageToolbar" class="pptx-toolbar">
        <button type="button" class="toolbar-btn" :disabled="pageCurrent <= 1" @click="goPage(pageCurrent - 1)">上一页</button>
        <div class="toolbar-page">
          <input
            v-model="pageInput"
            class="toolbar-input"
            type="number"
            min="1"
            :max="pageTotal || 1"
            @keyup.enter="jumpPage"
          />
          <span class="toolbar-sep">/</span>
          <span class="toolbar-total">{{ pageTotal || '-' }}</span>
        </div>
        <button type="button" class="toolbar-btn" @click="jumpPage">跳转</button>
        <button
          type="button"
          class="toolbar-btn"
          :disabled="!pageTotal || pageCurrent >= pageTotal"
          @click="goPage(pageCurrent + 1)"
        >
          下一页
        </button>

        <template v-if="showZoomControls">
          <span class="toolbar-divider" aria-hidden="true" />
          <button
            type="button"
            class="toolbar-btn"
            :disabled="viewZoom <= VIEW_ZOOM_MIN"
            @click="zoomViewBy(-VIEW_ZOOM_STEP)"
          >
            缩小
          </button>
          <span class="toolbar-zoom-label">{{ viewZoomPercent }}%</span>
          <button
            type="button"
            class="toolbar-btn"
            :disabled="viewZoom >= VIEW_ZOOM_MAX"
            @click="zoomViewBy(VIEW_ZOOM_STEP)"
          >
            放大
          </button>
          <button type="button" class="toolbar-btn" :disabled="viewZoom === VIEW_ZOOM_DEFAULT" @click="resetViewZoom">
            重置
          </button>
        </template>
      </div>
    </div>
  </div>
</template>

<script>
import { defineAsyncComponent } from 'vue'
import '@vue-office/docx/lib/index.css'
import '@vue-office/excel/lib/index.css'
import {
  cellAddress,
  collectExcelMatches,
  getExcelSelection,
  getExcelSpreadsheet,
  gotoExcelCell,
} from './excelPreview.js'

const TYPE_MAP = {
  docx: 'docx',
  doc: 'docx',
  word: 'docx',
  xlsx: 'excel',
  xls: 'excel',
  excel: 'excel',
  spreadsheet: 'excel',
  pdf: 'pdf',
  pptx: 'pptx',
  ppt: 'pptx',
  powerpoint: 'pptx',
}

function readLocationParams() {
  const search = new URLSearchParams(window.location.search)
  let url = search.get('url')
  let resourceType = search.get('resource_type')

  if (!url || !resourceType) {
    const hash = window.location.hash || ''
    const queryIndex = hash.indexOf('?')
    if (queryIndex >= 0) {
      const hashParams = new URLSearchParams(hash.slice(queryIndex + 1))
      url = url || hashParams.get('url')
      resourceType = resourceType || hashParams.get('resource_type')
    }
  }

  return { url, resourceType }
}

function inferTypeFromUrl(url) {
  try {
    const pathname = decodeURIComponent(new URL(url).pathname).toLowerCase()
    if (pathname.endsWith('.docx') || pathname.endsWith('.doc')) return 'docx'
    if (pathname.endsWith('.xlsx') || pathname.endsWith('.xls')) return 'excel'
    if (pathname.endsWith('.pdf')) return 'pdf'
    if (pathname.endsWith('.pptx') || pathname.endsWith('.ppt')) return 'pptx'
  } catch {
    return ''
  }
  return ''
}

function toProxySrc(url) {
  if (!url) return ''
  if (url.startsWith('/') && !url.startsWith('//')) return url
  if (url.startsWith('blob:') || url.startsWith('data:')) return url
  return `/file-proxy?url=${encodeURIComponent(url)}`
}

function formatSize(bytes) {
  if (!bytes) return '0M'
  if (bytes < 1024 * 1024) {
    return `${Math.max(0.1, bytes / (1024 * 1024)).toFixed(1)}M`
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)}M`
}

function formatElapsed(ms) {
  return `${(Math.max(0, ms) / 1000).toFixed(3)}s`
}

const RING_RADIUS = 54
const RING_LENGTH = 2 * Math.PI * RING_RADIUS
const PPTX_SIDEBAR_WIDTH = 148
const PPTX_TOOLBAR_HEIGHT = 52
const VIEW_ZOOM_MIN = 0.5
const VIEW_ZOOM_MAX = 2.5
const VIEW_ZOOM_STEP = 0.1
const VIEW_ZOOM_DEFAULT = 1

const SITE_THEME = {
  docx: { color: '#2B579A', label: 'W', title: 'Word 预览' },
  excel: { color: '#217346', label: 'X', title: 'Excel 预览' },
  pdf: { color: '#E5252A', label: 'PDF', title: 'PDF 预览' },
  pptx: { color: '#C43E1C', label: 'P', title: 'PPT 预览' },
  default: { color: '#4B5563', label: 'O', title: 'Office Preview' },
}

function buildFaviconHref(color, label) {
  const fontSize = label.length > 2 ? 10 : label.length > 1 ? 13 : 16
  const svg = `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
  <rect width="32" height="32" rx="7" fill="${color}"/>
  <text x="16" y="21.5" text-anchor="middle" fill="#fff"
    font-family="Segoe UI, Arial, sans-serif" font-size="${fontSize}" font-weight="700">${label}</text>
</svg>`.trim()
  return `data:image/svg+xml,${encodeURIComponent(svg)}`
}

function applySiteChrome(previewType) {
  const theme = SITE_THEME[previewType] || SITE_THEME.default
  document.title = theme.title

  let link = document.querySelector("link[rel='icon']")
  if (!link) {
    link = document.createElement('link')
    link.setAttribute('rel', 'icon')
    document.head.appendChild(link)
  }
  link.setAttribute('type', 'image/svg+xml')
  link.setAttribute('href', buildFaviconHref(theme.color, theme.label))

  let meta = document.querySelector('meta[name="theme-color"]')
  if (!meta) {
    meta = document.createElement('meta')
    meta.setAttribute('name', 'theme-color')
    document.head.appendChild(meta)
  }
  meta.setAttribute('content', theme.color)
}

function getPptxLayoutSize() {
  const width = Math.max((window.innerWidth || 960) - PPTX_SIDEBAR_WIDTH, 320)
  const height = Math.max((window.innerHeight || 540) - PPTX_TOOLBAR_HEIGHT, 240)
  return { width, height }
}

export default {
  name: 'App',
  components: {
    VueOfficeDocx: defineAsyncComponent(() => import('@vue-office/docx')),
    VueOfficeExcel: defineAsyncComponent(() => import('@vue-office/excel')),
    VueOfficePdf: defineAsyncComponent(() => import('@vue-office/pdf')),
    VueOfficePptx: defineAsyncComponent(() => import('@vue-office/pptx')),
  },
  data() {
    return {
      loading: true,
      loadingText: '正在下载...',
      elapsedText: '0.000s',
      hint: '',
      homeUrl: '',
      homeType: 'auto',
      homeError: '',
      src: '',
      originalUrl: '',
      usedProxy: false,
      previewType: '',
      downloadReceived: 0,
      downloadTotal: 0,
      phase: '',
      timerStartedAt: 0,
      statusTimer: null,
      pptxPage: 1,
      pptxPageInput: '1',
      pptxTotalPages: 0,
      pptxScrollEl: null,
      pptxReadyPages: {},
      pptxThumbTimer: null,
      pptxSlideObserver: null,
      pdfInited: false,
      viewZoom: VIEW_ZOOM_DEFAULT,
      zoomWheelEl: null,
      VIEW_ZOOM_MIN,
      VIEW_ZOOM_MAX,
      VIEW_ZOOM_STEP,
      VIEW_ZOOM_DEFAULT,
      requestOptions: {
        headers: {},
      },
      pptxOptions: getPptxLayoutSize(),
      excelOptions: {
        xls: false,
        minColLength: 0,
        minRowLength: 0,
        widthOffset: 10,
        heightOffset: 10,
        beforeTransformData: (workbookData) => workbookData,
        transformData: (workbookData) => workbookData,
      },
      excelFindOpen: false,
      excelFindQuery: '',
      excelFindHits: [],
      excelFindIndex: 0,
      excelCopyToast: '',
      excelCopyTimer: null,
    }
  },
  computed: {
    showHome() {
      return !this.originalUrl && !this.src && !this.loading
    },
    homeTypeHint() {
      if (this.homeType !== 'auto') return ''
      const inferred = inferTypeFromUrl(this.homeUrl.trim())
      if (!inferred) return ''
      const labels = { docx: 'Word', excel: 'Excel', pdf: 'PDF', pptx: 'PPT' }
      return `已识别为 ${labels[inferred] || inferred}`
    },
    excelFindCountText() {
      if (!this.excelFindQuery.trim()) return ''
      if (!this.excelFindHits.length) return '无结果'
      const hit = this.excelFindHits[this.excelFindIndex]
      const where = hit ? `${hit.sheetName}!${cellAddress(hit.ri, hit.ci)}` : ''
      return `${this.excelFindIndex + 1}/${this.excelFindHits.length}${where ? `  ${where}` : ''}`
    },
    downloadPercent() {
      if (!this.downloadTotal) return 0
      return Math.min(100, (this.downloadReceived / this.downloadTotal) * 100)
    },
    ringLength() {
      return RING_LENGTH
    },
    ringOffset() {
      if (this.phase === 'render') {
        return RING_LENGTH * 0.75
      }
      const percent = this.downloadPercent / 100
      return RING_LENGTH * (1 - percent)
    },
    statusLabel() {
      if (this.phase === 'download') return '正在下载'
      if (this.phase === 'render') return '正在渲染'
      return '加载中'
    },
    downloadSizeText() {
      if (this.phase !== 'download' || !this.downloadTotal) return ''
      return `${formatSize(this.downloadReceived)} / ${formatSize(this.downloadTotal)}`
    },
    showPageToolbar() {
      return (
        (this.previewType === 'pptx' ||
          this.previewType === 'pdf' ||
          this.previewType === 'docx') &&
        !this.loading &&
        !this.hint &&
        this.pageTotal > 0
      )
    },
    pageCurrent: {
      get() {
        return this.pptxPage
      },
      set(value) {
        this.pptxPage = value
      },
    },
    pageInput: {
      get() {
        return this.pptxPageInput
      },
      set(value) {
        this.pptxPageInput = value
      },
    },
    pageTotal() {
      return this.pptxTotalPages
    },
    pageReady() {
      return this.pptxReadyPages
    },
    pageList() {
      return Array.from({ length: this.pageTotal }, (_, index) => index + 1)
    },
    showZoomControls() {
      return (
        (this.previewType === 'docx' ||
          this.previewType === 'pdf' ||
          this.previewType === 'pptx') &&
        this.showPageToolbar
      )
    },
    viewZoomPercent() {
      return Math.round(this.viewZoom * 100)
    },
  },
  watch: {
    pageCurrent(page) {
      this.$nextTick(() => this.scrollSidebarToActive(page))
    },
  },
  created() {
    this.applyQuery()
  },
  mounted() {
    window.addEventListener('popstate', this.applyQuery)
    window.addEventListener('keydown', this.onPreviewKeydown, true)
    window.addEventListener('resize', this.syncSidebarOverflow)
    document.addEventListener('copy', this.onPreviewCopy, true)
  },
  beforeUnmount() {
    window.removeEventListener('popstate', this.applyQuery)
    window.removeEventListener('keydown', this.onPreviewKeydown, true)
    window.removeEventListener('resize', this.syncSidebarOverflow)
    document.removeEventListener('copy', this.onPreviewCopy, true)
    this.stopStatusTimer()
    this.unbindPageScroll()
    this.stopThumbCapture()
    this.unbindViewZoom()
    this.closeExcelFind()
    this.clearExcelCopyToast()
  },
  methods: {
    onPreviewKeydown(event) {
      if (this.previewType !== 'excel' || this.loading) return
      const key = event.key.toLowerCase()
      const mod = event.ctrlKey || event.metaKey
      if (mod && key === 'f') {
        event.preventDefault()
        event.stopPropagation()
        this.openExcelFind()
        return
      }
      if (this.excelFindOpen && key === 'escape') {
        event.preventDefault()
        this.closeExcelFind()
        return
      }
      if (this.excelFindOpen && key === 'f3') {
        event.preventDefault()
        this.stepExcelFind(event.shiftKey ? -1 : 1)
      }
    },
    onPreviewCopy(event) {
      if (this.previewType !== 'excel' || this.loading) return
      const target = event.target
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable)) {
        return
      }
      const selection = getExcelSelection(getExcelSpreadsheet())
      if (!selection) {
        this.showExcelCopyToast('没有可复制的内容')
        return
      }
      event.preventDefault()
      event.stopPropagation()
      event.clipboardData?.setData('text/plain', selection.text)
      const cells = selection.rowCount * selection.colCount
      const size =
        cells === 1
          ? selection.address
          : `${selection.address}（${selection.rowCount}×${selection.colCount}）`
      this.showExcelCopyToast(`已复制 ${size}`)
    },
    showExcelCopyToast(message) {
      this.excelCopyToast = message
      if (this.excelCopyTimer) clearTimeout(this.excelCopyTimer)
      this.excelCopyTimer = setTimeout(() => {
        this.excelCopyToast = ''
        this.excelCopyTimer = null
      }, 1800)
    },
    clearExcelCopyToast() {
      if (this.excelCopyTimer) {
        clearTimeout(this.excelCopyTimer)
        this.excelCopyTimer = null
      }
      this.excelCopyToast = ''
    },
    openExcelFind() {
      this.excelFindOpen = true
      this.$nextTick(() => {
        const input = this.$refs.excelFindInput
        if (input) {
          input.focus()
          input.select()
        }
        if (this.excelFindQuery.trim()) this.runExcelFind()
      })
    },
    closeExcelFind() {
      this.excelFindOpen = false
      this.excelFindHits = []
      this.excelFindIndex = 0
    },
    runExcelFind() {
      const xs = getExcelSpreadsheet()
      this.excelFindHits = collectExcelMatches(xs, this.excelFindQuery)
      this.excelFindIndex = 0
      if (this.excelFindHits.length) {
        gotoExcelCell(xs, this.excelFindHits[0])
      }
    },
    stepExcelFind(delta) {
      const hits = this.excelFindHits
      if (!hits.length) return
      const next = (this.excelFindIndex + delta + hits.length) % hits.length
      this.excelFindIndex = next
      gotoExcelCell(getExcelSpreadsheet(), hits[next])
    },
    onExcelFindKeydown(event) {
      if (event.key === 'Enter') {
        event.preventDefault()
        this.stepExcelFind(event.shiftKey ? -1 : 1)
      } else if (event.key === 'Escape') {
        event.preventDefault()
        this.closeExcelFind()
      }
    },
    applyQuery() {
      const { url, resourceType } = readLocationParams()
      if (!url) {
        this.originalUrl = ''
        this.src = ''
        this.hint = ''
        this.loading = false
        this.previewType = ''
        applySiteChrome('')
        return
      }

      const normalized = (resourceType || '').toLowerCase().replace(/^\./, '')
      const previewType = TYPE_MAP[normalized] || inferTypeFromUrl(url)
      if (!previewType) {
        this.originalUrl = ''
        this.src = ''
        this.homeUrl = url
        this.homeType = 'auto'
        this.homeError = `无法识别文件类型，请手动选择。支持 Word / Excel / PDF / PPT。`
        this.loading = false
        this.previewType = ''
        applySiteChrome('')
        return
      }

      this.hint = ''
      this.homeError = ''
      this.previewType = previewType
      applySiteChrome(previewType)
      this.excelOptions.xls = normalized === 'xls'
      this.originalUrl = url
      this.usedProxy = false
      this.src = ''
      this.downloadReceived = 0
      this.downloadTotal = 0
      this.pptxPage = 1
      this.pptxPageInput = '1'
      this.pptxTotalPages = 0
      this.pptxReadyPages = {}
      this.pdfInited = false
      this.viewZoom = VIEW_ZOOM_DEFAULT
      this.closeExcelFind()
      this.unbindPageScroll()
      this.stopThumbCapture()
      this.unbindViewZoom()
      if (previewType === 'pptx') {
        this.pptxOptions = getPptxLayoutSize()
      }
      this.loading = true
      this.loadingText = '正在下载... 0.000s'
      this.elapsedText = '0.000s'
      this.loadFile(url)
    },
    submitHome() {
      const raw = this.homeUrl.trim()
      this.homeError = ''
      if (!raw) {
        this.homeError = '请输入文件地址'
        return
      }
      const url = /^[a-z][a-z0-9+.-]*:/i.test(raw) ? raw : `https://${raw}`
      try {
        // eslint-disable-next-line no-new
        new URL(url)
      } catch {
        this.homeError = '请输入有效的 URL'
        return
      }
      this.homeUrl = url
      const type = this.homeType === 'auto' ? inferTypeFromUrl(url) : this.homeType
      if (!type) {
        this.homeError = '无法从地址识别类型，请手动选择'
        return
      }
      const params = new URLSearchParams({ url, resource_type: type })
      window.history.pushState({}, '', `?${params.toString()}`)
      this.applyQuery()
    },
    async loadFile(url) {
      try {
        const buffer = await this.fetchBuffer(url)
        this.downloadTotal = buffer.byteLength
        this.downloadReceived = buffer.byteLength
        this.startStatusTimer('render')
        this.src = buffer
      } catch (error) {
        this.stopStatusTimer()
        console.error('文件下载失败', error)
        if (!this.usedProxy && this.originalUrl) {
          this.usedProxy = true
          this.downloadReceived = 0
          this.downloadTotal = 0
          this.loadingText = '直连失败，改用代理下载...'
          await this.loadFile(toProxySrc(this.originalUrl))
          return
        }
        this.loading = false
        this.src = ''
        this.hint = error instanceof Error ? error.message : '文件下载失败'
      }
    },
    async fetchBuffer(url) {
      this.startStatusTimer('download')
      const response = await fetch(url, this.requestOptions)
      if (!response.ok) {
        throw new Error(`下载失败 HTTP ${response.status}`)
      }

      const total = Number(response.headers.get('content-length')) || 0
      this.downloadTotal = total
      this.downloadReceived = 0
      this.updateDownloadText(0, total)

      if (!response.body) {
        const buffer = await response.arrayBuffer()
        this.downloadReceived = buffer.byteLength
        this.downloadTotal = buffer.byteLength
        this.updateDownloadText(buffer.byteLength, buffer.byteLength)
        return buffer
      }

      const reader = response.body.getReader()
      const chunks = []
      let received = 0

      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        chunks.push(value)
        received += value.byteLength
        this.downloadReceived = received
        this.updateDownloadText(received, total)
      }

      const bytes = new Uint8Array(received)
      let offset = 0
      for (const chunk of chunks) {
        bytes.set(chunk, offset)
        offset += chunk.byteLength
      }
      return bytes.buffer
    },
    updateDownloadText(received, total) {
      const elapsed = formatElapsed(performance.now() - this.timerStartedAt)
      this.elapsedText = elapsed
      if (total) {
        this.loadingText = `正在下载 ${formatSize(received)} / ${formatSize(total)} ${elapsed}`
        return
      }
      this.loadingText = received
        ? `正在下载 ${formatSize(received)} ${elapsed}`
        : `正在下载... ${elapsed}`
    },
    startStatusTimer(phase) {
      this.stopStatusTimer()
      this.phase = phase
      this.timerStartedAt = performance.now()
      this.elapsedText = '0.000s'
      const tick = () => {
        const elapsed = formatElapsed(performance.now() - this.timerStartedAt)
        this.elapsedText = elapsed
        if (this.phase === 'download') {
          this.updateDownloadText(this.downloadReceived, this.downloadTotal)
        } else if (this.phase === 'render') {
          this.loadingText = `正在渲染... ${elapsed}`
        }
        this.statusTimer = window.requestAnimationFrame(tick)
      }
      tick()
    },
    stopStatusTimer() {
      if (this.statusTimer) {
        window.cancelAnimationFrame(this.statusTimer)
        this.statusTimer = null
      }
      this.phase = ''
    },
    getPptxRoot() {
      return document.querySelector('.preview-area > .vue-office-pptx')
    },
    getPptxScrollEl() {
      return this.getPptxRoot()?.querySelector('.pptx-preview-wrapper') || null
    },
    getPptxSlides() {
      const scrollEl = this.getPptxScrollEl()
      if (!scrollEl) return []
      // Only direct children of the main preview wrapper — never sidebar clones.
      return Array.from(scrollEl.children).filter((el) =>
        el.classList?.contains('pptx-preview-slide-wrapper'),
      )
    },
    fitPptxSlideShapes() {
      const wrapper = this.getPptxScrollEl()
      const slides = this.getPptxSlides()
      if (!wrapper || !slides.length) return
      const classic = wrapper.offsetWidth - wrapper.clientWidth
      const base = Math.max(1, wrapper.clientWidth - (classic > 0 ? 0 : 16))
      const target = Math.max(1, Math.round(base * this.viewZoom))
      if (target < 2) return
      for (const slide of slides) {
        const layers = Array.from(slide.children).filter(
          (el) =>
            el.classList.contains('slide-wrapper') ||
            el.classList.contains('slide-master-wrapper') ||
            el.classList.contains('slide-layout-wrapper'),
        )
        let nativeW = 0
        let nativeH = 0
        for (const layer of layers) {
          nativeW = Math.max(nativeW, parseFloat(layer.style.width) || 0)
          nativeH = Math.max(nativeH, parseFloat(layer.style.height) || 0)
          layer.querySelectorAll('.shape-wrapper, .group').forEach((el) => {
            nativeW = Math.max(
              nativeW,
              (parseFloat(el.style.left) || 0) + (parseFloat(el.style.width) || 0),
            )
          })
        }
        if (nativeW < 2) continue
        const zoom = target / nativeW
        slide.style.width = `${target}px`
        if (nativeH) slide.style.height = `${Math.round(nativeH * zoom)}px`
        for (const layer of layers) {
          layer.style.transform = 'none'
          layer.style.zoom = String(zoom)
        }
      }
    },
    getDocxRoot() {
      return document.querySelector('.preview-area > .vue-office-docx')
    },
    getDocxScrollEl() {
      return this.getDocxRoot()
    },
    getDocxWrapper() {
      return this.getDocxRoot()?.querySelector('.docx-wrapper') || null
    },
    getDocxPages() {
      const wrapper = this.getDocxWrapper()
      if (!wrapper) return []
      // Only direct section.docx under main wrapper — never sidebar clones.
      return Array.from(wrapper.children).filter(
        (el) => el.tagName === 'SECTION' && el.classList?.contains('docx'),
      )
    },
    applyDocxZoom() {
      const wrapper = this.getDocxWrapper()
      if (!wrapper) return
      // 只缩放文档显示区域，侧栏/底栏不受影响
      wrapper.style.zoom = String(this.viewZoom)
    },
    applyPdfZoom() {
      const wrapper = this.getPdfWrapper()
      if (!wrapper) return
      // CSS zoom，避免 setScale 整页重绘导致闪屏/缩略图变黑
      wrapper.style.zoom = String(this.viewZoom)
    },
    applyPptxZoom() {
      this.fitPptxSlideShapes()
    },
    applyViewZoom() {
      if (this.previewType === 'docx') {
        this.applyDocxZoom()
        return
      }
      if (this.previewType === 'pdf') {
        this.applyPdfZoom()
        return
      }
      if (this.previewType === 'pptx') {
        this.applyPptxZoom()
      }
    },
    setViewZoom(next) {
      const zoom = Math.min(
        VIEW_ZOOM_MAX,
        Math.max(VIEW_ZOOM_MIN, Number(next) || VIEW_ZOOM_DEFAULT),
      )
      this.viewZoom = Math.round(zoom * 100) / 100
      this.applyViewZoom()
    },
    zoomViewBy(delta) {
      this.setViewZoom(this.viewZoom + delta)
    },
    resetViewZoom() {
      this.setViewZoom(VIEW_ZOOM_DEFAULT)
    },
    bindViewZoom() {
      this.unbindViewZoom()
      this.applyViewZoom()
      const root =
        this.previewType === 'pdf'
          ? this.getPdfScrollEl()
          : this.previewType === 'docx'
            ? this.getDocxScrollEl()
            : this.previewType === 'pptx'
              ? this.getPptxScrollEl()
              : null
      if (!root) return
      this.zoomWheelEl = root
      this.onViewWheel = (event) => {
        if (!(event.ctrlKey || event.metaKey)) return
        event.preventDefault()
        const delta = event.deltaY > 0 ? -VIEW_ZOOM_STEP : VIEW_ZOOM_STEP
        this.zoomViewBy(delta)
      }
      root.addEventListener('wheel', this.onViewWheel, { passive: false })
    },
    unbindViewZoom() {
      if (this.zoomWheelEl && this.onViewWheel) {
        this.zoomWheelEl.removeEventListener('wheel', this.onViewWheel)
      }
      this.zoomWheelEl = null
      this.onViewWheel = null
      const docxWrapper = this.getDocxWrapper()
      if (docxWrapper) docxWrapper.style.zoom = ''
      const pdfWrapper = this.getPdfWrapper()
      if (pdfWrapper) pdfWrapper.style.zoom = ''
    },
    getPdfRoot() {
      return document.querySelector('.preview-area > .vue-office-pdf')
    },
    getPdfScrollEl() {
      return this.getPdfRoot()
    },
    getPdfWrapper() {
      return this.getPdfRoot()?.querySelector('.vue-office-pdf-wrapper') || null
    },
    getPdfCanvases() {
      const wrapper = this.getPdfWrapper()
      if (!wrapper) return []
      return Array.from(
        wrapper.querySelectorAll('.vue-office-pdf-page canvas, :scope > canvas[data-id]'),
      )
    },
    getPdfMetrics() {
      const wrapper = this.getPdfWrapper()
      if (!wrapper) return null
      const canvas = this.getPdfCanvases()[0]
      if (!canvas) return null
      const pageBox = canvas.closest('.vue-office-pdf-page')
      const pageH =
        parseFloat(pageBox?.style.height) ||
        parseFloat(canvas.style.height) ||
        canvas.clientHeight
      if (!pageH) return null
      const gap = 10
      const contentH = parseFloat(wrapper.style.height) || 0
      if (!contentH) return null
      const total = Math.max(1, Math.round((contentH + gap) / (pageH + gap)))
      const zoom = parseFloat(wrapper.style.zoom) || this.viewZoom || 1
      return {
        pageH,
        gap,
        stride: pageH + gap,
        total,
        zoom,
        padding: parseInt(getComputedStyle(wrapper).paddingTop, 10) || 0,
      }
    },
    bindPageScroll() {
      this.unbindPageScroll()
      if (this.previewType === 'pdf') {
        const scrollEl = this.getPdfScrollEl()
        if (!scrollEl) return
        this.pptxScrollEl = scrollEl
        this.onPptxScroll = () => this.syncPdfPageFromScroll()
        scrollEl.addEventListener('scroll', this.onPptxScroll, { passive: true })
        this.syncPdfPageFromScroll()
        return
      }
      if (this.previewType === 'docx') {
        const scrollEl = this.getDocxScrollEl()
        if (!scrollEl) return
        this.pptxScrollEl = scrollEl
        this.onPptxScroll = () => this.syncDocxPageFromScroll()
        scrollEl.addEventListener('scroll', this.onPptxScroll, { passive: true })
        this.syncDocxPageFromScroll()
        return
      }
      const scrollEl = this.getPptxScrollEl()
      if (!scrollEl) return
      this.pptxScrollEl = scrollEl
      this.onPptxScroll = () => this.syncPptxPageFromScroll()
      scrollEl.addEventListener('scroll', this.onPptxScroll, { passive: true })
      this.syncPptxPageFromScroll()
    },
    unbindPageScroll() {
      if (this.pptxScrollEl && this.onPptxScroll) {
        this.pptxScrollEl.removeEventListener('scroll', this.onPptxScroll)
      }
      this.pptxScrollEl = null
      this.onPptxScroll = null
    },
    syncPptxPageFromScroll() {
      const scrollEl = this.pptxScrollEl || this.getPptxScrollEl()
      const slides = this.getPptxSlides()
      if (!scrollEl || !slides.length) return

      const top = scrollEl.scrollTop + 8
      let current = 1
      for (let i = 0; i < slides.length; i += 1) {
        if (slides[i].offsetTop <= top) {
          current = i + 1
        } else {
          break
        }
      }
      this.pptxPage = current
      if (document.activeElement?.classList?.contains('toolbar-input')) return
      this.pptxPageInput = String(current)
    },
    getElementScrollTop(el, scrollEl) {
      if (!el || !scrollEl) return 0
      const elRect = el.getBoundingClientRect()
      const scrollRect = scrollEl.getBoundingClientRect()
      return elRect.top - scrollRect.top + scrollEl.scrollTop
    },
    syncDocxPageFromScroll() {
      const scrollEl = this.pptxScrollEl || this.getDocxScrollEl()
      const pages = this.getDocxPages()
      if (!scrollEl || !pages.length) return

      const top = scrollEl.scrollTop + 8
      let current = 1
      for (let i = 0; i < pages.length; i += 1) {
        if (this.getElementScrollTop(pages[i], scrollEl) <= top) {
          current = i + 1
        } else {
          break
        }
      }
      this.pptxPage = current
      if (document.activeElement?.classList?.contains('toolbar-input')) return
      this.pptxPageInput = String(current)
    },
    syncPdfPageFromScroll() {
      const scrollEl = this.pptxScrollEl || this.getPdfScrollEl()
      const metrics = this.getPdfMetrics()
      if (!scrollEl || !metrics) return
      const zoom = metrics.zoom || 1
      const current = Math.min(
        metrics.total,
        Math.max(1, Math.floor((scrollEl.scrollTop / zoom + 8) / metrics.stride) + 1),
      )
      this.pptxPage = current
      if (document.activeElement?.classList?.contains('toolbar-input')) return
      this.pptxPageInput = String(current)
    },
    scrollSidebarToActive(page) {
      const list = this.$refs.pptxSidebarList
      if (!list) return
      const active = list.querySelector(`.pptx-thumb[data-page="${page}"]`)
      if (!active) return
      const top = active.offsetTop - list.clientHeight / 2 + active.clientHeight / 2
      list.scrollTo({ top: Math.max(0, top), behavior: 'smooth' })
    },
    syncSidebarOverflow() {
      const list = this.$refs.pptxSidebarList
      if (!list) return
      list.style.overflowY = 'hidden'
      const overflowing = list.scrollHeight > list.clientHeight + 1
      list.style.overflowY = overflowing ? 'auto' : 'hidden'
    },
    capturePptxThumb(page, slideEl) {
      if (!slideEl || this.pptxReadyPages[page]) return
      const host = document.querySelector(`.pptx-sidebar [data-thumb-page="${page}"]`)
      if (!host) return
      const width = slideEl.offsetWidth || 0
      const height = slideEl.offsetHeight || 0
      if (!width || !height) return

      const frame = host.parentElement
      const frameWidth = frame?.clientWidth || 112
      const scale = frameWidth / width
      host.innerHTML = ''
      const clone = slideEl.cloneNode(true)
      clone.classList.remove('pptx-preview-slide-wrapper')
      clone.classList.add('pptx-thumb-slide-clone')
      clone.style.margin = '0'
      clone.style.pointerEvents = 'none'
      host.style.width = `${width}px`
      host.style.height = `${height}px`
      host.style.transform = `scale(${scale})`
      host.appendChild(clone)
      this.pptxReadyPages = { ...this.pptxReadyPages, [page]: true }
    },
    captureDocxThumb(page, pageEl) {
      if (!pageEl || this.pptxReadyPages[page]) return
      const host = document.querySelector(`.pptx-sidebar [data-thumb-page="${page}"]`)
      if (!host) return
      const width = pageEl.offsetWidth || 0
      const height = pageEl.offsetHeight || 0
      if (!width || !height) return

      const frame = host.parentElement
      const frameWidth = frame?.clientWidth || 112
      const scale = frameWidth / width
      host.innerHTML = ''
      const clone = pageEl.cloneNode(true)
      clone.classList.remove('docx')
      clone.classList.add('docx-thumb-page-clone')
      clone.style.margin = '0'
      clone.style.boxShadow = 'none'
      clone.style.pointerEvents = 'none'
      host.style.width = `${width}px`
      host.style.height = `${height}px`
      host.style.transform = `scale(${scale})`
      host.appendChild(clone)
      this.pptxReadyPages = { ...this.pptxReadyPages, [page]: true }
    },
    capturePdfThumb(page, canvas) {
      if (!canvas || this.pptxReadyPages[page]) return
      const host = document.querySelector(`.pptx-sidebar [data-thumb-page="${page}"]`)
      if (!host) return
      if (canvas.width < 2 || canvas.height < 2) return

      // 跳过尚未绘制完成的空白/近黑画布，避免侧栏变黑
      try {
        const ctx = canvas.getContext('2d', { willReadFrequently: true })
        if (ctx) {
          const sample = ctx.getImageData(
            Math.floor(canvas.width / 2),
            Math.floor(canvas.height / 2),
            1,
            1,
          ).data
          const [r, g, b, a] = sample
          if (a < 10 || r + g + b < 24) return
        }
      } catch {
        // ignore cross-origin / security errors and still try dataURL
      }

      let dataUrl = ''
      try {
        dataUrl = canvas.toDataURL('image/jpeg', 0.72)
      } catch {
        return
      }
      if (!dataUrl || dataUrl === 'data:,') return

      host.innerHTML = ''
      host.style.width = '100%'
      host.style.height = '100%'
      host.style.transform = 'none'
      const img = document.createElement('img')
      img.src = dataUrl
      img.alt = `第 ${page} 页`
      img.draggable = false
      img.style.width = '100%'
      img.style.height = '100%'
      img.style.objectFit = 'cover'
      img.style.objectPosition = 'top center'
      img.style.pointerEvents = 'none'
      host.appendChild(img)
      this.pptxReadyPages = { ...this.pptxReadyPages, [page]: true }
    },
    refreshPptxThumbs() {
      this.fitPptxSlideShapes()
      const slides = this.getPptxSlides()
      const limit = this.pptxTotalPages > 0 ? this.pptxTotalPages : slides.length
      for (let index = 0; index < Math.min(slides.length, limit); index += 1) {
        this.capturePptxThumb(index + 1, slides[index])
      }
    },
    refreshDocxThumbs() {
      const pages = this.getDocxPages()
      const limit = this.pptxTotalPages > 0 ? this.pptxTotalPages : pages.length
      for (let index = 0; index < Math.min(pages.length, limit); index += 1) {
        this.captureDocxThumb(index + 1, pages[index])
      }
    },
    refreshPdfThumbs() {
      const canvases = this.getPdfCanvases()
      for (const canvas of canvases) {
        const page = Number(
          canvas.getAttribute('data-id') ||
            canvas.closest('[data-id]')?.getAttribute('data-id'),
        )
        if (!page) continue
        this.capturePdfThumb(page, canvas)
      }
    },
    startThumbCapture() {
      this.stopThumbCapture()
      this.$nextTick(() => {
        this.syncSidebarOverflow()
        if (this.previewType === 'pdf') {
          this.refreshPdfThumbs()
          this.pptxThumbTimer = window.setInterval(() => {
            this.refreshPdfThumbs()
            if (
              this.pptxTotalPages > 0 &&
              Object.keys(this.pptxReadyPages).length >= this.pptxTotalPages
            ) {
              this.stopThumbCapture(false)
            }
          }, 600)

          const wrapper = this.getPdfWrapper()
          if (wrapper && typeof MutationObserver !== 'undefined') {
            this.pptxSlideObserver = new MutationObserver(() => this.refreshPdfThumbs())
            this.pptxSlideObserver.observe(wrapper, { childList: true })
          }
          return
        }

        if (this.previewType === 'docx') {
          this.refreshDocxThumbs()
          this.pptxThumbTimer = window.setInterval(() => {
            this.refreshDocxThumbs()
            if (
              this.pptxTotalPages > 0 &&
              Object.keys(this.pptxReadyPages).length >= this.pptxTotalPages
            ) {
              this.stopThumbCapture(false)
            }
          }, 600)

          const wrapper = this.getDocxWrapper()
          if (wrapper && typeof MutationObserver !== 'undefined') {
            this.pptxSlideObserver = new MutationObserver(() => this.refreshDocxThumbs())
            this.pptxSlideObserver.observe(wrapper, { childList: true })
          }
          return
        }

        this.refreshPptxThumbs()
        this.pptxThumbTimer = window.setInterval(() => {
          this.refreshPptxThumbs()
          if (
            this.pptxTotalPages > 0 &&
            Object.keys(this.pptxReadyPages).length >= this.pptxTotalPages
          ) {
            this.stopThumbCapture(false)
          }
        }, 600)

        const wrapper = this.getPptxScrollEl()
        if (wrapper && typeof MutationObserver !== 'undefined') {
          this.pptxSlideObserver = new MutationObserver(() => this.refreshPptxThumbs())
          this.pptxSlideObserver.observe(wrapper, { childList: true })
        }
      })
    },
    stopThumbCapture(clearObserver = true) {
      if (this.pptxThumbTimer) {
        window.clearInterval(this.pptxThumbTimer)
        this.pptxThumbTimer = null
      }
      if (clearObserver && this.pptxSlideObserver) {
        this.pptxSlideObserver.disconnect()
        this.pptxSlideObserver = null
      }
    },
    async waitForPptxSlide(page, timeoutMs = 8000) {
      const index = page - 1
      const started = Date.now()
      while (Date.now() - started < timeoutMs) {
        const slides = this.getPptxSlides()
        if (slides[index]) return slides[index]
        await new Promise((resolve) => setTimeout(resolve, 50))
      }
      return null
    },
    async waitForDocxPage(page, timeoutMs = 8000) {
      const index = page - 1
      const started = Date.now()
      while (Date.now() - started < timeoutMs) {
        const pages = this.getDocxPages()
        if (pages[index]) return pages[index]
        await new Promise((resolve) => setTimeout(resolve, 50))
      }
      return null
    },
    async goPage(page) {
      if (this.previewType === 'pdf') {
        this.goPdfPage(page)
        return
      }
      if (this.previewType === 'docx') {
        await this.goDocxPage(page)
        return
      }
      await this.goPptxPage(page)
    },
    goPdfPage(page) {
      const metrics = this.getPdfMetrics()
      if (!metrics) return
      const target = Math.min(metrics.total, Math.max(1, Number(page) || 1))
      this.pptxPageInput = String(target)
      this.pptxPage = target
      const scrollEl = this.getPdfScrollEl()
      if (!scrollEl) return
      const zoom = metrics.zoom || 1
      scrollEl.scrollTo({
        top: (target - 1) * metrics.stride * zoom,
        behavior: 'smooth',
      })
    },
    async goDocxPage(page) {
      if (!this.pptxTotalPages) return
      const target = Math.min(this.pptxTotalPages, Math.max(1, Number(page) || 1))
      this.pptxPageInput = String(target)
      this.pptxPage = target
      const pageEl = await this.waitForDocxPage(target)
      if (!pageEl) return
      const scrollEl = this.getDocxScrollEl()
      if (scrollEl) {
        scrollEl.scrollTo({
          top: this.getElementScrollTop(pageEl, scrollEl),
          behavior: 'smooth',
        })
      } else {
        pageEl.scrollIntoView({ behavior: 'smooth', block: 'start' })
      }
    },
    async goPptxPage(page) {
      if (!this.pptxTotalPages) return
      const target = Math.min(this.pptxTotalPages, Math.max(1, Number(page) || 1))
      this.pptxPageInput = String(target)
      this.pptxPage = target
      const slide = await this.waitForPptxSlide(target)
      if (!slide) return
      const scrollEl = this.getPptxScrollEl()
      if (scrollEl) {
        scrollEl.scrollTo({
          top: slide.offsetTop,
          behavior: 'smooth',
        })
      } else {
        slide.scrollIntoView({ behavior: 'smooth', block: 'start' })
      }
    },
    jumpPage() {
      this.goPage(this.pptxPageInput)
    },
    renderedHandler(payload) {
      this.stopStatusTimer()
      this.loading = false
      this.hint = ''
      if (this.previewType === 'pptx') {
        // Trust the pptx model page count; do not infer from DOM (thumb clones used to inflate it).
        const total = payload?.slides?.length || 0
        this.pptxTotalPages = total
        this.pptxPage = 1
        this.pptxPageInput = '1'
        this.pptxReadyPages = {}
        this.$nextTick(() => {
          this.fitPptxSlideShapes()
          this.bindPageScroll()
          this.startThumbCapture()
          this.bindViewZoom()
        })
      } else if (this.previewType === 'pdf') {
        // PDF virtual list re-emits rendered as pages load; init once, then only refresh thumbs.
        if (!this.pdfInited) {
          const metrics = this.getPdfMetrics()
          const total = metrics?.total || 0
          if (total > 0) {
            this.pdfInited = true
            this.pptxTotalPages = total
            this.pptxPage = 1
            this.pptxPageInput = '1'
            this.pptxReadyPages = {}
            this.$nextTick(() => {
              this.bindPageScroll()
              this.startThumbCapture()
              this.bindViewZoom()
            })
          }
        } else {
          this.refreshPdfThumbs()
        }
      } else if (this.previewType === 'docx') {
        this.$nextTick(() => {
          const total = this.getDocxPages().length
          this.pptxTotalPages = total
          this.pptxPage = 1
          this.pptxPageInput = '1'
          this.pptxReadyPages = {}
          this.viewZoom = VIEW_ZOOM_DEFAULT
          if (total > 0) {
            this.bindPageScroll()
            this.startThumbCapture()
            this.$nextTick(() => {
              this.bindViewZoom()
            })
          }
        })
      }
      console.log('渲染完成')
    },
    errorHandler(error) {
      this.stopStatusTimer()
      this.unbindPageScroll()
      this.stopThumbCapture()
      this.unbindViewZoom()
      this.pptxTotalPages = 0
      this.pptxReadyPages = {}
      this.pdfInited = false
      this.viewZoom = VIEW_ZOOM_DEFAULT
      this.closeExcelFind()
      console.error('渲染失败', error)
      this.loading = false
      this.hint = error?.message || '文件渲染失败，请检查 url 是否可访问，以及 resource_type 是否匹配。'
    },
  },
}
</script>

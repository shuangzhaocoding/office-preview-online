export function getExcelSpreadsheet() {
  return document.querySelector('.vue-office-excel-main')?.__spreadsheet || null
}

export function colLabel(ci) {
  let n = Number(ci) + 1
  let label = ''
  while (n > 0) {
    n -= 1
    label = String.fromCharCode(65 + (n % 26)) + label
    n = Math.floor(n / 26)
  }
  return label
}

export function cellAddress(ri, ci) {
  return `${colLabel(ci)}${Number(ri) + 1}`
}

export function getExcelSelection(xs) {
  if (!xs?.sheet?.data) return null
  const data = xs.sheet.data
  const range = data.selector?.range
  if (!range) return null
  const { sri, eri, sci, eci } = range
  const rows = []
  for (let ri = sri; ri <= eri; ri += 1) {
    const row = []
    for (let ci = sci; ci <= eci; ci += 1) {
      const cell = data.getCell(ri, ci)
      row.push(cell && cell.text != null ? String(cell.text) : '')
    }
    rows.push(row)
  }
  const rowCount = eri - sri + 1
  const colCount = eci - sci + 1
  const start = cellAddress(sri, sci)
  const end = cellAddress(eri, eci)
  return {
    text: rows.map((row) => row.join('\t')).join('\n'),
    rowCount,
    colCount,
    address: rowCount === 1 && colCount === 1 ? start : `${start}:${end}`,
  }
}

export function getExcelSelectionText(xs) {
  return getExcelSelection(xs)?.text || ''
}

export function collectExcelMatches(xs, query) {
  const q = String(query || '').trim().toLowerCase()
  if (!q || !xs?.datas) return []
  const hits = []
  xs.datas.forEach((sheet, sheetIndex) => {
    const rows = sheet.rows?._ || {}
    Object.keys(rows).forEach((riKey) => {
      const ri = Number(riKey)
      if (Number.isNaN(ri)) return
      const cells = rows[riKey]?.cells || {}
      Object.keys(cells).forEach((ciKey) => {
        const ci = Number(ciKey)
        if (Number.isNaN(ci)) return
        const text = cells[ciKey]?.text
        if (text == null) return
        const value = String(text)
        if (value.toLowerCase().includes(q)) {
          hits.push({
            sheetIndex,
            sheetName: sheet.name || `Sheet${sheetIndex + 1}`,
            ri,
            ci,
            text: value,
          })
        }
      })
    })
  })
  return hits
}

export function gotoExcelCell(xs, match) {
  if (!xs?.sheet || !match) return
  const current = xs.datas.indexOf(xs.sheet.data)
  if (match.sheetIndex !== current && xs.bottombar?.items?.[match.sheetIndex]) {
    xs.bottombar.clickSwap2(xs.bottombar.items[match.sheetIndex])
  }
  const sheet = xs.sheet
  sheet.selector.set(match.ri, match.ci)
  sheet.focusing = true
  scrollSheetToSelection(sheet)
}

function scrollSheetToSelection(sheet) {
  const data = sheet.data
  const rect = data.getSelectedRect()
  if (!rect) return
  const tableOffset = sheet.getTableOffset()
  const { l, t, left, top, width, height } = rect
  if (Math.abs(left) + width > tableOffset.width) {
    sheet.horizontalScrollbar.move({ left: l + width - tableOffset.width })
  } else {
    const fsw = data.freezeTotalWidth()
    if (left < fsw) {
      sheet.horizontalScrollbar.move({ left: l - 1 - fsw })
    }
  }
  if (Math.abs(top) + height > tableOffset.height) {
    sheet.verticalScrollbar.move({ top: t + height - tableOffset.height - 1 })
  } else {
    const fsh = data.freezeTotalHeight()
    if (top < fsh) {
      sheet.verticalScrollbar.move({ top: t - 1 - fsh })
    }
  }
}

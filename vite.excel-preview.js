// Keep preview read-only, but allow copying the current selection.
// Expose the x-spreadsheet instance on the mount element for find/copy.

const COPY_RE =
  /function copy\(evt\) \{\s*var data = this\.data, selector = this\.selector;\s*if \(data\.settings\.mode === "read"\)\s*return;\s*data\.copy\(\);\s*data\.copyToSystemClipboard\(evt\);\s*selector\.showClipboard\(\);\s*\}/

const COPY_FIXED = `function copy(evt) {
  var data = this.data, selector = this.selector;
  data.copy();
  data.copyToSystemClipboard(evt);
  selector.showClipboard();
}`

const EXPOSE_RE =
  /autoFocus: false\s*\}\)\.loadData\(\{\}\);\s*xs\.on\("cell-selected"/

const EXPOSE_FIXED = `autoFocus: false
        }).loadData({});
        if (rootRef.value) rootRef.value.__spreadsheet = xs;
        xs.on("cell-selected"`

export function excelPreviewPlugin() {
  return {
    name: 'excel-preview-fix',
    transform(code, id) {
      const normalized = id.replace(/\\/g, '/')
      if (!normalized.includes('@vue-office/excel')) {
        return null
      }
      let next = code
      if (COPY_RE.test(next)) next = next.replace(COPY_RE, COPY_FIXED)
      if (EXPOSE_RE.test(next)) next = next.replace(EXPOSE_RE, EXPOSE_FIXED)
      return next === code ? null : { code: next, map: null }
    },
  }
}

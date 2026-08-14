// Make PDF CSS zoom (on .vue-office-pdf-wrapper) work with the virtual list:
// scrollTop is in zoomed coordinates, while page metrics (u) are not.

const SCROLL_RE =
  /const Kn = BC\.debounce\(function\(v\) \{\s*const \{ scrollTop: p \} = v\.target;\s*let w = parseInt\(getComputedStyle\(a\.value\)\.paddingTop\) \|\| 0;\s*const d = e\.options\.gap \|\| 10, S = Math\.max\(\s*Math\.floor\(Math\.max\(p - w, 0\) \/ \(u \+ d\)\) - 1,\s*1\s*\), x = Math\.min\(S \+ g - 1, s\);\s*Ge\(S, x\);\s*\}, 60\);/

const SCROLL_FIXED = `const Kn = BC.debounce(function(v) {
      const { scrollTop: p } = v.target;
      let w = parseInt(getComputedStyle(a.value).paddingTop) || 0;
      const z = parseFloat(a.value.style.zoom) || 1;
      const d = e.options.gap || 10, S = Math.max(
        Math.floor(Math.max(p / z - w, 0) / (u + d)) - 1,
        1
      ), x = Math.min(S + g - 1, s);
      Ge(S, x);
    }, 60);`

export function pdfScalePlugin() {
  return {
    name: 'pdf-scale-fix',
    transform(code, id) {
      const normalized = id.replace(/\\/g, '/')
      if (!normalized.includes('@vue-office/pdf')) {
        return null
      }

      let next = code
      if (SCROLL_RE.test(next)) {
        next = next.replace(SCROLL_RE, SCROLL_FIXED)
      }
      return next === code ? null : { code: next, map: null }
    },
  }
}

// 1) CSS zoom: scrollTop is zoomed, page metrics are not.
// 2) Overlay pdf.js TextLayer so visible pages are selectable / Ctrl+F-able.

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

const CREATE_PAGE_RE =
  /function le\(v\) \{\s*let p = parseInt\(getComputedStyle\(a\.value\)\.paddingTop\) \|\| 0, w = e\.options\.gap \|\| 10;\s*const d = document\.createElement\("canvas"\);\s*return d\.style\.position = "absolute", d\.style\.top = \(v - 1\) \* \(u \+ w\) \+ p \+ "px", d\.style\.left = "50%", d\.style\.transform = "translate\(-50%\)", d\.style\.backgroundColor = "#fff", d\.setAttribute\("data-id", v\), d\.width = h, d\.height = y, d\.style\.width = `\$\{c\}px`, d\.style\.height = `\$\{u\}px`, d;\s*\}/

const CREATE_PAGE_FIXED = `function le(v) {
      let p = parseInt(getComputedStyle(a.value).paddingTop) || 0, w = e.options.gap || 10;
      const box = document.createElement("div");
      box.className = "vue-office-pdf-page";
      box.setAttribute("data-id", v);
      box.style.position = "absolute";
      box.style.top = (v - 1) * (u + w) + p + "px";
      box.style.left = "50%";
      box.style.transform = "translate(-50%)";
      box.style.width = \`\${c}px\`;
      box.style.height = \`\${u}px\`;
      box.style.backgroundColor = "#fff";
      const d = document.createElement("canvas");
      d.width = h;
      d.height = y;
      d.style.width = \`\${c}px\`;
      d.style.height = \`\${u}px\`;
      d.style.display = "block";
      const textLayer = document.createElement("div");
      textLayer.className = "textLayer";
      textLayer.style.setProperty("--scale-factor", "1");
      box.appendChild(d);
      box.appendChild(textLayer);
      return box;
    }`

const RENDER_PAGE_RE =
  /function ge\(v, p\) \{\s*return r\.getPage\(v\)\.then\(\(w\) => \{\s*const d = w\.getViewport\(\{ scale: I \* \(f\.value < 1 \? f\.value : 1\) \}\);\s*let S = window\.devicePixelRatio > 2 \? 1\.5 : 2;\s*if \(h \* d\.height !== y \* d\.width\) \{\s*let k = Math\.floor\(d\.width \* S\), T = Math\.floor\(d\.height \* S\), J = y \/ T;\s*S = S \* J, p\.width = k \* J, p\.style\.width = parseInt\(p\.style\.width\) \* J \+ "px";\s*\}\s*const x = p\.getContext\("2d"\), L = S !== 1 \? \[S, 0, 0, S, 0, 0\] : null;\s*return w\.render\(\{\s*canvasContext: x,\s*transform: L,\s*viewport: d\s*\}\)\.promise;\s*\}\);\s*\}/

const RENDER_PAGE_FIXED = `function ge(v, p) {
      const canvas = p && p.tagName === "CANVAS" ? p : p.querySelector("canvas");
      const textLayerEl = p && p.querySelector ? p.querySelector(".textLayer") : null;
      return r.getPage(v).then((w) => {
        const d = w.getViewport({ scale: I * (f.value < 1 ? f.value : 1) });
        let S = window.devicePixelRatio > 2 ? 1.5 : 2;
        if (h * d.height !== y * d.width) {
          let k = Math.floor(d.width * S), T = Math.floor(d.height * S), J = y / T;
          S = S * J;
          canvas.width = k * J;
          canvas.style.width = parseInt(canvas.style.width) * J + "px";
          if (p !== canvas) {
            p.style.width = canvas.style.width;
          }
        }
        const x = canvas.getContext("2d"), L = S !== 1 ? [S, 0, 0, S, 0, 0] : null;
        return w.render({
          canvasContext: x,
          transform: L,
          viewport: d
        }).promise.then(() => {
          if (!textLayerEl || !window.pdfjsLib || !window.pdfjsLib.TextLayer) return;
          textLayerEl.innerHTML = "";
          textLayerEl.style.setProperty("--scale-factor", String(d.scale));
          textLayerEl.style.width = canvas.style.width;
          textLayerEl.style.height = canvas.style.height;
          return w.getTextContent().then((textContent) => {
            const layer = new window.pdfjsLib.TextLayer({
              textContentSource: textContent,
              container: textLayerEl,
              viewport: d
            });
            return layer.render();
          });
        });
      });
    }`

const VH_RE =
  /function vh\(\)\{if\(null===Bl\(_l,this,bh\)\._\)\{var t=document\.createElement\("div"\);t\.style\.opacity=0,t\.style\.lineHeight=1,t\.style\.fontSize="1px",t\.style\.position="absolute",t\.textContent="X",document\.body\.append\(t\),bh\._=Bl\(_l,this,t\.getBoundingClientRect\(\)\.height\),t\.remove\(\)\}\}/

const VH_FIXED = `function vh(){if(null===Bl(_l,this,bh)._){bh._=Bl(_l,this,1)}}`

const HH_SCALE_RE = /if\(bh\._>1&&\(a="scale\("\.concat\(1\/bh\._,"\)"\)\),/

const HH_SCALE_FIXED = 'if(0&&bh._>1&&(a="scale(".concat(1/bh._,")")),'

export function pdfScalePlugin() {
  return {
    name: 'pdf-scale-fix',
    transform(code, id) {
      const normalized = id.replace(/\\/g, '/')
      if (!normalized.includes('@vue-office/pdf')) {
        return null
      }

      let next = code
      if (SCROLL_RE.test(next)) next = next.replace(SCROLL_RE, SCROLL_FIXED)
      if (CREATE_PAGE_RE.test(next)) next = next.replace(CREATE_PAGE_RE, CREATE_PAGE_FIXED)
      if (RENDER_PAGE_RE.test(next)) next = next.replace(RENDER_PAGE_RE, RENDER_PAGE_FIXED)
      if (VH_RE.test(next)) next = next.replace(VH_RE, VH_FIXED)
      if (HH_SCALE_RE.test(next)) next = next.replace(HH_SCALE_RE, HH_SCALE_FIXED)
      return next === code ? null : { code: next, map: null }
    },
  }
}

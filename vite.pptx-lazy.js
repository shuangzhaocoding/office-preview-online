const SLIDE_LOOP_RE = /for \(var o = new Rot\(t\.wrapper, i, \{ viewPort: \{ width: t\.options\.width, height: t\.options\.height \} \}\), s = 0; s < i\.slides\.length; s\+\+\) o\.renderSlide\(s\);\s*n\(i\);/

const SLIDE_LOOP_LAZY = `var o = new Rot(t.wrapper, i, { viewPort: { width: t.options.width, height: t.options.height } });
          var s = 0;
          var total = i.slides.length;
          if (!total) {
            n(i);
            return;
          }
          var renderNext = function() {
            try {
              o.renderSlide(s++);
            } catch (err) {
              a(err);
              return;
            }
            if (s === 1) n(i);
            if (s < total) {
              var sched = typeof requestAnimationFrame === "function" ? requestAnimationFrame : function(cb) { setTimeout(cb, 0); };
              sched(renderNext);
            }
          };
          renderNext();`

const ARRAY_BUFFER_RE = /if \(s instanceof ArrayBuffer\)\s*return Promise\.resolve\(s\);/

const ARRAY_BUFFER_WITH_BLOB = `if (s instanceof ArrayBuffer)
        return Promise.resolve(s);
      if (typeof Blob !== "undefined" && s instanceof Blob)
        return s.arrayBuffer();`

// 缩放分母取 pptx.width 与图形实际右边界的较大值，避免 shape 比 sldSz 更宽。
const CALC_SCALE_RE = /r\.prototype\._calcScaleAndRenderPort = function\(\) \{\s*var e = this\.options\.viewPort\.width \/ this\.pptx\.width;/

const CALC_SCALE_FIT = `r.prototype._calcScaleAndRenderPort = function() {
    var contentW = this.pptx.width || 1;
    var walk = function(nodes) {
      if (!nodes) return;
      for (var i = 0; i < nodes.length; i++) {
        var node = nodes[i];
        if (!node) continue;
        if (node.offset && node.extend) {
          var right = (node.offset.x || 0) + (node.extend.w || 0);
          if (right > contentW) contentW = right;
        }
        if (node.nodes) walk(node.nodes);
      }
    };
    var slides = this.pptx.slides || [];
    for (var si = 0; si < slides.length; si++) {
      var slide = slides[si];
      walk(slide.nodes);
      if (slide.slideMaster) walk(slide.slideMaster.nodes);
      if (slide.slideLayout) walk(slide.slideLayout.nodes);
    }
    var e = this.options.viewPort.width / contentW;`

// transform:scale 不改变布局盒；外层 overflow:hidden 仍按未缩放的 1382px 裁切。
// zoom 会缩小布局盒，图形会落到 slide-wrapper 宽度内。
const LAYER_SCALE_RE = /n\.style\.setProperty\("transform", "scale\("\.concat\(this\.scale, "\)"\)\), n\.style\.setProperty\("transform-origin", "0 0"\)/g

const LAYER_SCALE_ZOOM = `n.style.setProperty("zoom", String(this.scale))`

export function pptxLazyRenderPlugin() {
  return {
    name: 'pptx-lazy-render',
    transform(code, id) {
      const normalized = id.replace(/\\/g, '/')
      if (!normalized.includes('@vue-office/pptx')) {
        return null
      }

      let next = code
      if (SLIDE_LOOP_RE.test(next)) {
        next = next.replace(SLIDE_LOOP_RE, SLIDE_LOOP_LAZY)
      }
      if (ARRAY_BUFFER_RE.test(next)) {
        next = next.replace(ARRAY_BUFFER_RE, ARRAY_BUFFER_WITH_BLOB)
      }
      if (CALC_SCALE_RE.test(next)) {
        next = next.replace(CALC_SCALE_RE, CALC_SCALE_FIT)
      }
      if (LAYER_SCALE_RE.test(next)) {
        next = next.replace(LAYER_SCALE_RE, LAYER_SCALE_ZOOM)
      }
      return next === code ? null : { code: next, map: null }
    },
  }
}

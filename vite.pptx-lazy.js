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
      return next === code ? null : { code: next, map: null }
    },
  }
}

import http from 'node:http'
import https from 'node:https'

const MAX_REDIRECTS = 5
const REQUEST_TIMEOUT_MS = 180000

function getTargetUrl(req) {
  const raw = req.originalUrl || req.url || ''
  return new URL(raw, 'http://localhost').searchParams.get('url')
}

function pipeUpstream(target, res, redirectCount = 0) {
  let parsed
  try {
    parsed = new URL(target)
  } catch {
    res.statusCode = 400
    res.end('invalid url')
    return
  }

  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    res.statusCode = 400
    res.end('invalid protocol')
    return
  }

  const client = parsed.protocol === 'https:' ? https : http
  const req = client.get(parsed, {
    headers: {
      'User-Agent': 'office-preview',
      Accept: '*/*',
      'Accept-Encoding': 'identity',
    },
  }, (upstream) => {
    const status = upstream.statusCode || 500
    const location = upstream.headers.location

    if (status >= 300 && status < 400 && location) {
      upstream.resume()
      if (redirectCount >= MAX_REDIRECTS) {
        res.statusCode = 502
        res.end('too many redirects')
        return
      }
      pipeUpstream(new URL(location, parsed).href, res, redirectCount + 1)
      return
    }

    res.statusCode = status
    if (upstream.headers['content-type']) {
      res.setHeader('content-type', upstream.headers['content-type'])
    }
    if (upstream.headers['content-length']) {
      res.setHeader('content-length', upstream.headers['content-length'])
    }
    res.setHeader('cache-control', 'private, max-age=60')
    upstream.pipe(res)
  })

  req.on('error', (error) => {
    if (!res.headersSent) {
      res.statusCode = 502
    }
    res.end(error.message)
  })
  req.setTimeout(REQUEST_TIMEOUT_MS, () => {
    req.destroy(new Error('upstream timeout'))
  })
}

export function fileProxyPlugin() {
  const middleware = (req, res, next) => {
    const path = (req.originalUrl || req.url || '').split('?')[0]
    if (path !== '/file-proxy') {
      next()
      return
    }
    const target = getTargetUrl(req)
    if (!target) {
      res.statusCode = 400
      res.end('missing url')
      return
    }
    pipeUpstream(target, res)
  }

  return {
    name: 'file-proxy',
    configureServer(server) {
      server.middlewares.use(middleware)
    },
    configurePreviewServer(server) {
      server.middlewares.use(middleware)
    },
  }
}

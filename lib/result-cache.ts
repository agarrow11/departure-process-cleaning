// In-memory cache for pipeline output files, keyed by a resultId.
//
// The /api/pipeline route computes four output files (full xlsx, no-PII
// xlsx, ecode mapping csv, no-PII csv) and used to inline all four as base64
// directly in its JSON response. That made the response grow with the
// combined size of every output file at once, which stopped fitting through
// an intermediary in this environment once the schema widened. Caching the
// buffers here lets the route return a small JSON payload and the client
// fetch each file individually, only when a download button is clicked.
//
// This is a single-process, in-memory cache — correct for this internal
// single-user tool, which runs as one long-lived Node process, but it will
// NOT survive a server restart or work across multiple serverless instances.
// If this tool is ever deployed behind multiple function instances, replace
// this with a real shared store (e.g. Blob) before relying on it.

export interface PipelineFiles {
  xlsx: Buffer
  ecodeMap: Buffer
  xlsxRedacted: Buffer
  csvRedacted: Buffer
}

interface CacheEntry {
  files: PipelineFiles
  createdAt: number
}

const TTL_MS = 60 * 60 * 1000 // 1 hour — long enough to click every download button
const MAX_ENTRIES = 5 // this tool is single-user; keep memory bounded regardless

const cache = new Map<string, CacheEntry>()

function evictStale() {
  const now = Date.now()
  for (const [id, entry] of cache) {
    if (now - entry.createdAt > TTL_MS) cache.delete(id)
  }
  while (cache.size > MAX_ENTRIES) {
    const oldestKey = cache.keys().next().value
    if (oldestKey === undefined) break
    cache.delete(oldestKey)
  }
}

export function cacheResult(id: string, files: PipelineFiles): void {
  evictStale()
  cache.set(id, { files, createdAt: Date.now() })
}

export function getCachedResult(id: string): PipelineFiles | null {
  const entry = cache.get(id)
  return entry ? entry.files : null
}

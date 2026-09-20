import {
  deserializeMetrics,
  serializeMetrics,
  type FontMetrics,
  type SerializedFontMetrics,
} from './metrics'

const DATABASE_NAME = 'atlas-docx-font-metrics'
const STORE_NAME = 'metrics'

const memoryCache = new Map<string, SerializedFontMetrics>()
let databasePromise: Promise<IDBDatabase> | null = null

/**
 * MEM-01 follow-up — this in-memory L1 cache (mirroring the persistent
 * IndexedDB store) has no explicit size cap, unlike `canvasMetrics.ts`'s
 * fragment-width cache or `downscaleImage.ts`'s LRU. That's safe by
 * construction rather than by omission: every key is
 * `${substituteName}@${variant}@${hash(fileUrl)}` (see `loadFontMetrics` in
 * `loader.ts`), and both `substituteName` and `fileUrl` are drawn from
 * `FONT_FAMILIES` in `families.ts` — a fixed, bundled catalog of 5 substitute
 * families x 4 `FontVariant`s, independent of how many documents (or how
 * many distinct Word font names) get opened in a session. So this map can
 * never hold more than `FONT_FAMILIES.length * 4` entries no matter how long
 * the app runs — see `cache.test.ts`'s regression test, which opens far more
 * distinct (arbitrary, made-up) font names than that and asserts the cache
 * never grows past the fixed bound.
 */
export function __memoryCacheSizeForTests(): number {
  return memoryCache.size
}

export async function getCachedMetrics(key: string): Promise<FontMetrics | null> {
  if (!hasIndexedDb()) {
    return deserializeFromMemory(key)
  }

  try {
    const database = await getDatabase()
    const serialized = await readFromStore(database, key)
    if (!serialized) {
      return null
    }

    memoryCache.set(key, serialized)
    return deserializeMetrics(serialized)
  } catch {
    return deserializeFromMemory(key)
  }
}

export async function setCachedMetrics(key: string, metrics: FontMetrics): Promise<void> {
  const serialized = serializeMetrics(metrics)
  memoryCache.set(key, serialized)

  if (!hasIndexedDb()) {
    return
  }

  try {
    const database = await getDatabase()
    await writeToStore(database, key, serialized)
  } catch {
    return
  }
}

function hasIndexedDb(): boolean {
  return typeof indexedDB !== 'undefined'
}

function deserializeFromMemory(key: string): FontMetrics | null {
  const serialized = memoryCache.get(key)
  return serialized ? deserializeMetrics(serialized) : null
}

function getDatabase(): Promise<IDBDatabase> {
  if (!databasePromise) {
    databasePromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(DATABASE_NAME, 1)

      request.onerror = () => {
        reject(request.error ?? new Error('Failed to open font metrics cache'))
      }

      request.onupgradeneeded = () => {
        const database = request.result
        if (!database.objectStoreNames.contains(STORE_NAME)) {
          database.createObjectStore(STORE_NAME)
        }
      }

      request.onsuccess = () => {
        resolve(request.result)
      }
    })
  }

  return databasePromise
}

function readFromStore(database: IDBDatabase, key: string): Promise<SerializedFontMetrics | null> {
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readonly')
    const request = transaction.objectStore(STORE_NAME).get(key)

    request.onerror = () => {
      reject(request.error ?? new Error('Failed to read font metrics cache'))
    }

    request.onsuccess = () => {
      const value = request.result
      resolve(isSerializedFontMetrics(value) ? value : null)
    }
  })
}

function writeToStore(database: IDBDatabase, key: string, serialized: SerializedFontMetrics): Promise<void> {
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readwrite')
    const request = transaction.objectStore(STORE_NAME).put(serialized, key)

    request.onerror = () => {
      reject(request.error ?? new Error('Failed to write font metrics cache'))
    }

    transaction.oncomplete = () => {
      resolve()
    }

    transaction.onerror = () => {
      reject(transaction.error ?? new Error('Failed to commit font metrics cache write'))
    }
  })
}

function isSerializedFontMetrics(value: unknown): value is SerializedFontMetrics {
  if (!value || typeof value !== 'object') {
    return false
  }

  const candidate = value as Record<string, unknown>
  return typeof candidate.unitsPerEm === 'number'
    && typeof candidate.ascender === 'number'
    && typeof candidate.descender === 'number'
    && typeof candidate.lineGap === 'number'
    && typeof candidate.xHeight === 'number'
    && typeof candidate.capHeight === 'number'
    && typeof candidate.widths === 'object'
    && candidate.widths !== null
}

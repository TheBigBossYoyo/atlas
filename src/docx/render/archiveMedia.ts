/**
 * Atlas — archive-backed media resolver
 *
 * Turns the image relationships of a loaded DOCX into blob: URLs for
 * `MediaContext`. URL ownership lives outside render:
 *   - URLs are cached per archive path AND byte identity, so an unchanged
 *     picture keeps its URL across edits, while a different document reusing
 *     the same relationship id / path never receives a stale image.
 *   - URLs no longer referenced are revoked after commit (never mid-render,
 *     so the page on screen never points at a revoked URL).
 *   - Unmount revocation is deferred one task so React StrictMode's simulated
 *     unmount/remount does not revoke URLs the mounted tree still uses.
 */

import { useEffect, useMemo, useState } from 'react'

import type { Relationship } from '../parser/relationships'
import type { MediaResolver } from './mediaContext'
import { imageMimeType, resolveImagePartPath } from './mediaParts'

type CachedUrl = { readonly bytes: Uint8Array; readonly url: string }

type ArchiveMediaResolver = MediaResolver & { readonly urls: ReadonlySet<string> }

class BlobUrlRegistry {
  private readonly byPath = new Map<string, CachedUrl>()
  private readonly live = new Set<string>()
  private pendingRevokeAll: ReturnType<typeof setTimeout> | null = null

  acquire(path: string, bytes: Uint8Array, mime: string): string | null {
    const cached = this.byPath.get(path)
    if (cached !== undefined && cached.bytes === bytes) {
      return cached.url
    }
    if (typeof URL.createObjectURL !== 'function') {
      return null
    }

    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
    const url = URL.createObjectURL(new Blob([buffer], { type: mime }))
    this.byPath.set(path, { bytes, url })
    this.live.add(url)
    return url
  }

  retainOnly(keep: ReadonlySet<string>): void {
    this.cancelScheduledRevoke()
    for (const url of this.live) {
      if (!keep.has(url)) {
        this.revoke(url)
      }
    }
  }

  scheduleRevokeAll(): void {
    this.cancelScheduledRevoke()
    this.pendingRevokeAll = setTimeout(() => {
      this.pendingRevokeAll = null
      this.retainOnly(new Set())
    }, 0)
  }

  private cancelScheduledRevoke(): void {
    if (this.pendingRevokeAll !== null) {
      clearTimeout(this.pendingRevokeAll)
      this.pendingRevokeAll = null
    }
  }

  private revoke(url: string): void {
    this.live.delete(url)
    for (const [path, cached] of this.byPath) {
      if (cached.url === url) {
        this.byPath.delete(path)
      }
    }
    URL.revokeObjectURL(url)
  }
}

export function useArchiveMediaResolver(
  archive: ReadonlyMap<string, Uint8Array> | undefined,
  relationships: ReadonlyArray<Relationship> | undefined,
): MediaResolver {
  const [registry] = useState(() => new BlobUrlRegistry())

  const resolver = useMemo<ArchiveMediaResolver>(
    () => buildResolver(registry, archive, relationships ?? []),
    [registry, archive, relationships],
  )

  useEffect(() => {
    registry.retainOnly(resolver.urls)
    return () => registry.scheduleRevokeAll()
  }, [registry, resolver])

  return resolver
}

function buildResolver(
  registry: BlobUrlRegistry,
  archive: ReadonlyMap<string, Uint8Array> | undefined,
  relationships: ReadonlyArray<Relationship>,
): ArchiveMediaResolver {
  const urlByRelationshipId = new Map<string, string>()

  if (archive !== undefined) {
    for (const relationship of relationships) {
      const path = resolveImagePartPath(relationship)
      // Unpaintable formats (EMF/WMF/TIFF) stay unresolved so the page
      // shows a visible placeholder rather than a broken image.
      const mime = path === null ? null : imageMimeType(path)
      const bytes = path === null ? undefined : archive.get(path)
      if (path === null || mime === null || bytes === undefined) {
        continue
      }

      const url = registry.acquire(path, bytes, mime)
      if (url !== null) {
        urlByRelationshipId.set(relationship.id, url)
      }
    }
  }

  return {
    urls: new Set(urlByRelationshipId.values()),
    resolve: (relationshipId: string) => urlByRelationshipId.get(relationshipId) ?? null,
  }
}

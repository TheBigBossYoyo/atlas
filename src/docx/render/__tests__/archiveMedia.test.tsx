import { StrictMode, type ReactNode } from 'react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { renderHook } from '@testing-library/react'

import type { Relationship } from '../../parser/relationships'
import { useArchiveMediaResolver } from '../archiveMedia'
import { IMAGE_RELATIONSHIP_TYPE } from '../mediaParts'

let urlCounter = 0
const revoked = new Set<string>()

beforeEach(() => {
  vi.useFakeTimers()
  urlCounter = 0
  revoked.clear()
  vi.stubGlobal('URL', Object.assign(URL, {
    createObjectURL: vi.fn(() => `blob:atlas/${++urlCounter}`),
    revokeObjectURL: vi.fn((url: string) => {
      revoked.add(url)
    }),
  }))
})

afterEach(() => {
  vi.useRealTimers()
  vi.unstubAllGlobals()
})

function imageRel(id: string, target: string): Relationship {
  return { id, type: IMAGE_RELATIONSHIP_TYPE, target }
}

function archiveWith(entries: Record<string, Uint8Array>): ReadonlyMap<string, Uint8Array> {
  return new Map(Object.entries(entries))
}

const strict = ({ children }: { children: ReactNode }) => <StrictMode>{children}</StrictMode>

type HookProps = {
  archive: ReadonlyMap<string, Uint8Array> | undefined
  relationships: ReadonlyArray<Relationship> | undefined
}

function renderResolver(initial: HookProps) {
  return renderHook(
    ({ archive, relationships }: HookProps) => useArchiveMediaResolver(archive, relationships),
    { initialProps: initial, wrapper: strict },
  )
}

describe('useArchiveMediaResolver', () => {
  it('resolves embedded image relationships to blob URLs', () => {
    const { result } = renderResolver({
      archive: archiveWith({ 'word/media/image1.png': new Uint8Array([1]) }),
      relationships: [imageRel('rId5', 'media/image1.png')],
    })

    expect(result.current.resolve('rId5')).toMatch(/^blob:atlas\//)
    expect(result.current.resolve('rId404')).toBeNull()
  })

  it('does not revoke live URLs during StrictMode effect replay', () => {
    const { result } = renderResolver({
      archive: archiveWith({ 'word/media/image1.png': new Uint8Array([1]) }),
      relationships: [imageRel('rId5', 'media/image1.png')],
    })

    vi.runAllTimers()

    const url = result.current.resolve('rId5')
    expect(url).not.toBeNull()
    expect(revoked.has(url as string)).toBe(false)
  })

  it('reuses the URL for unchanged bytes when relationships change', () => {
    const bytes = new Uint8Array([1])
    const archive = archiveWith({ 'word/media/image1.png': bytes, 'word/media/image2.png': new Uint8Array([2]) })
    const { result, rerender } = renderResolver({
      archive,
      relationships: [imageRel('rId5', 'media/image1.png')],
    })
    const before = result.current.resolve('rId5')

    rerender({
      archive,
      relationships: [imageRel('rId5', 'media/image1.png'), imageRel('rId6', 'media/image2.png')],
    })
    vi.runAllTimers()

    expect(result.current.resolve('rId5')).toBe(before)
    expect(result.current.resolve('rId6')).not.toBeNull()
    expect(revoked.has(before as string)).toBe(false)
  })

  it('never serves a previous document image for a reused relationship id', () => {
    const { result, rerender } = renderResolver({
      archive: archiveWith({ 'word/media/image1.png': new Uint8Array([1]) }),
      relationships: [imageRel('rId5', 'media/image1.png')],
    })
    const firstDocumentUrl = result.current.resolve('rId5')

    rerender({
      archive: archiveWith({ 'word/media/image1.png': new Uint8Array([9, 9]) }),
      relationships: [imageRel('rId5', 'media/image1.png')],
    })
    vi.runAllTimers()

    expect(result.current.resolve('rId5')).not.toBe(firstDocumentUrl)
    expect(revoked.has(firstDocumentUrl as string)).toBe(true)
  })

  it('leaves unpaintable formats unresolved so a placeholder renders', () => {
    const { result } = renderResolver({
      archive: archiveWith({ 'word/media/image1.emf': new Uint8Array([1]) }),
      relationships: [imageRel('rId5', 'media/image1.emf')],
    })

    expect(result.current.resolve('rId5')).toBeNull()
    expect(URL.createObjectURL).not.toHaveBeenCalled()
  })

  it('revokes every URL after a real unmount', () => {
    const { result, unmount } = renderResolver({
      archive: archiveWith({ 'word/media/image1.png': new Uint8Array([1]) }),
      relationships: [imageRel('rId5', 'media/image1.png')],
    })
    const url = result.current.resolve('rId5') as string

    unmount()
    vi.runAllTimers()

    expect(revoked.has(url)).toBe(true)
  })
})

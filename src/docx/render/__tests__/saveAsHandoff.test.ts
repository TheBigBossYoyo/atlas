/**
 * DOCX-3 — the Save As handoff slot.
 *
 * The behaviour worth pinning here is not "a value goes in and comes out"; it is
 * everything the slot refuses to do. It survives a remount by living outside the
 * React tree, which is exactly what makes a leaked or mis-addressed slot able to
 * drop a caret recorded in one document into a different one.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'

import type { Position, Range } from '../../editor'
import { claimSaveAsHandoff, resetSaveAsHandoff, stashSaveAsHandoff } from '../saveAsHandoff'

function position(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return { paragraphPath: Object.freeze([...paragraphPath]), runIndex, charOffset }
}

const CARET: Range = { anchor: position([1], 0, 16), focus: position([1], 0, 16) }

afterEach(() => {
  resetSaveAsHandoff()
  vi.useRealTimers()
})

describe('saveAsHandoff', () => {
  it('hands the caret and scroll position to the viewer that mounts at the target path', () => {
    stashSaveAsHandoff({ targetPath: '/docs/renamed.docx', range: CARET, scrollTop: 240, scrollLeft: 12 })

    expect(claimSaveAsHandoff('/docs/renamed.docx')).toEqual({
      targetPath: '/docs/renamed.docx',
      range: CARET,
      scrollTop: 240,
      scrollLeft: 12,
    })
  })

  it('is claimed once — a second viewer mounting at the same path gets nothing', () => {
    stashSaveAsHandoff({ targetPath: '/docs/renamed.docx', range: CARET, scrollTop: 0, scrollLeft: 0 })

    expect(claimSaveAsHandoff('/docs/renamed.docx')).not.toBeNull()
    expect(claimSaveAsHandoff('/docs/renamed.docx')).toBeNull()
  })

  it('gives nothing to a different document, and does not keep the slot for later', () => {
    stashSaveAsHandoff({ targetPath: '/docs/renamed.docx', range: CARET, scrollTop: 0, scrollLeft: 0 })

    // The user opened something else before the renamed viewer ever mounted.
    expect(claimSaveAsHandoff('/docs/other.docx')).toBeNull()
    // And the slot is gone rather than waiting: applying a caret recorded in one
    // document to another is worse than applying nothing.
    expect(claimSaveAsHandoff('/docs/renamed.docx')).toBeNull()
  })

  it('expires rather than restoring a caret from an abandoned save much later', () => {
    vi.useFakeTimers()
    stashSaveAsHandoff({ targetPath: '/docs/renamed.docx', range: CARET, scrollTop: 0, scrollLeft: 0 })

    vi.advanceTimersByTime(15_001)

    expect(claimSaveAsHandoff('/docs/renamed.docx')).toBeNull()
  })

  it('still claims within the window, so a slow re-parse does not lose the caret', () => {
    vi.useFakeTimers()
    stashSaveAsHandoff({ targetPath: '/docs/renamed.docx', range: CARET, scrollTop: 0, scrollLeft: 0 })

    vi.advanceTimersByTime(14_000)

    expect(claimSaveAsHandoff('/docs/renamed.docx')?.range).toEqual(CARET)
  })

  it('keeps only the most recent stash', () => {
    stashSaveAsHandoff({ targetPath: '/docs/first.docx', range: CARET, scrollTop: 0, scrollLeft: 0 })
    stashSaveAsHandoff({ targetPath: '/docs/second.docx', range: null, scrollTop: 99, scrollLeft: 0 })

    expect(claimSaveAsHandoff('/docs/second.docx')?.scrollTop).toBe(99)
  })

  it('returns null when nothing was ever stashed, which is the ordinary mount', () => {
    expect(claimSaveAsHandoff('/docs/opened.docx')).toBeNull()
  })
})

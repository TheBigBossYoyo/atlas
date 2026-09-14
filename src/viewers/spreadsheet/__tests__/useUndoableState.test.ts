/**
 * Wave 3 spreadsheet editing — generic undo/redo history hook.
 */
import { act, renderHook } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import { useUndoableState } from '../useUndoableState'

describe('useUndoableState', () => {
  it('starts with the initial value and nothing to undo/redo', () => {
    const { result } = renderHook(() => useUndoableState(0))
    expect(result.current.present).toBe(0)
    expect(result.current.canUndo).toBe(false)
    expect(result.current.canRedo).toBe(false)
  })

  it('set() pushes a new present and enables undo', () => {
    const { result } = renderHook(() => useUndoableState(0))
    act(() => result.current.set(1))
    expect(result.current.present).toBe(1)
    expect(result.current.canUndo).toBe(true)
  })

  it('undo() restores the previous present and enables redo', () => {
    const { result } = renderHook(() => useUndoableState(0))
    act(() => result.current.set(1))
    act(() => result.current.undo())
    expect(result.current.present).toBe(0)
    expect(result.current.canUndo).toBe(false)
    expect(result.current.canRedo).toBe(true)
  })

  it('redo() re-applies an undone change', () => {
    const { result } = renderHook(() => useUndoableState(0))
    act(() => result.current.set(1))
    act(() => result.current.undo())
    act(() => result.current.redo())
    expect(result.current.present).toBe(1)
    expect(result.current.canRedo).toBe(false)
  })

  it('a new set() after undo() discards the redo branch', () => {
    const { result } = renderHook(() => useUndoableState(0))
    act(() => result.current.set(1))
    act(() => result.current.undo())
    act(() => result.current.set(2))
    expect(result.current.present).toBe(2)
    expect(result.current.canRedo).toBe(false)
  })

  it('undo() on an empty history is a no-op', () => {
    const { result } = renderHook(() => useUndoableState(0))
    act(() => result.current.undo())
    expect(result.current.present).toBe(0)
  })

  it('redo() with nothing to redo is a no-op', () => {
    const { result } = renderHook(() => useUndoableState(0))
    act(() => result.current.redo())
    expect(result.current.present).toBe(0)
  })

  it('reset() replaces the present and clears all history', () => {
    const { result } = renderHook(() => useUndoableState(0))
    act(() => result.current.set(1))
    act(() => result.current.set(2))
    act(() => result.current.reset(99))
    expect(result.current.present).toBe(99)
    expect(result.current.canUndo).toBe(false)
    expect(result.current.canRedo).toBe(false)
    act(() => result.current.undo())
    expect(result.current.present).toBe(99)
  })

  it('bounds history to maxHistory entries', () => {
    const { result } = renderHook(() => useUndoableState(0, 2))
    act(() => result.current.set(1))
    act(() => result.current.set(2))
    act(() => result.current.set(3))
    // Only the last 2 past states (1, 2) should survive; undoing 3 times
    // should stop at 1, not reach all the way back to 0.
    act(() => result.current.undo())
    act(() => result.current.undo())
    expect(result.current.present).toBe(1)
    expect(result.current.canUndo).toBe(false)
  })
})

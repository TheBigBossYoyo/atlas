import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import {
  STATE_FILE_NAME,
  clampBoundsToDisplays,
  isSufficientlyVisible,
  loadWindowState,
  saveWindowState,
} from '../../../electron/lib/windowState.cjs'

const PRIMARY_DISPLAY = { workArea: { x: 0, y: 0, width: 1920, height: 1080 } }
const SECONDARY_DISPLAY = { workArea: { x: 1920, y: 0, width: 1280, height: 1024 } }

describe('isSufficientlyVisible (P5.2/ELEC-11)', () => {
  it('is visible when fully inside a display work area', () => {
    expect(isSufficientlyVisible({ x: 100, y: 100, width: 800, height: 600 }, [PRIMARY_DISPLAY])).toBe(true)
  })

  it('is visible when only a corner overlaps, as long as it clears the minimum threshold', () => {
    // Bottom-right corner of the display: exactly 50px of overlap on both axes.
    expect(isSufficientlyVisible({ x: 1870, y: 1030, width: 800, height: 600 }, [PRIMARY_DISPLAY])).toBe(true)
  })

  it('is not visible when entirely off every display (a disconnected external monitor)', () => {
    expect(isSufficientlyVisible({ x: 3200, y: 0, width: 800, height: 600 }, [PRIMARY_DISPLAY])).toBe(false)
  })

  it('checks against every connected display, not just the first', () => {
    expect(isSufficientlyVisible({ x: 2000, y: 100, width: 400, height: 300 }, [PRIMARY_DISPLAY, SECONDARY_DISPLAY])).toBe(
      true,
    )
  })

  it('rejects a sliver overlap below the minimum-visible-pixels threshold', () => {
    expect(isSufficientlyVisible({ x: 1919, y: 0, width: 800, height: 600 }, [PRIMARY_DISPLAY])).toBe(false)
  })
})

describe('clampBoundsToDisplays', () => {
  it('returns the bounds unchanged when they are still recoverable', () => {
    const bounds = { x: 100, y: 100, width: 800, height: 600 }
    expect(clampBoundsToDisplays(bounds, [PRIMARY_DISPLAY])).toEqual(bounds)
  })

  it('returns null when the saved position is off every current display', () => {
    const bounds = { x: 5000, y: 5000, width: 800, height: 600 }
    expect(clampBoundsToDisplays(bounds, [PRIMARY_DISPLAY])).toBeNull()
  })

  it('returns null when there are no bounds to clamp', () => {
    expect(clampBoundsToDisplays(null, [PRIMARY_DISPLAY])).toBeNull()
  })

  it('returns null when there are no connected displays at all', () => {
    expect(clampBoundsToDisplays({ x: 0, y: 0, width: 800, height: 600 }, [])).toBeNull()
  })
})

describe('loadWindowState / saveWindowState', () => {
  let storeDir: string

  beforeEach(() => {
    storeDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-window-state-'))
  })

  afterEach(() => {
    fs.rmSync(storeDir, { recursive: true, force: true })
  })

  it('returns null when no state file exists yet (first launch)', () => {
    expect(loadWindowState(storeDir)).toBeNull()
  })

  it('round-trips a saved state', () => {
    const state = { x: 50, y: 60, width: 1000, height: 700, isMaximized: false }
    saveWindowState(storeDir, state)
    expect(loadWindowState(storeDir)).toEqual(state)
  })

  it('writes its backing file under the given directory', () => {
    saveWindowState(storeDir, { x: 0, y: 0, width: 1200, height: 800, isMaximized: false })
    expect(fs.existsSync(path.join(storeDir, STATE_FILE_NAME))).toBe(true)
  })

  it('preserves isMaximized:true across a round-trip', () => {
    saveWindowState(storeDir, { x: 0, y: 0, width: 1920, height: 1080, isMaximized: true })
    expect(loadWindowState(storeDir)?.isMaximized).toBe(true)
  })

  it('rejects a corrupt (non-JSON) state file instead of throwing', () => {
    fs.writeFileSync(path.join(storeDir, STATE_FILE_NAME), '{ not valid json', 'utf-8')
    expect(() => loadWindowState(storeDir)).not.toThrow()
    expect(loadWindowState(storeDir)).toBeNull()
  })

  it('rejects a structurally-invalid state file (missing numeric fields)', () => {
    fs.writeFileSync(path.join(storeDir, STATE_FILE_NAME), JSON.stringify({ isMaximized: true }), 'utf-8')
    expect(loadWindowState(storeDir)).toBeNull()
  })

  it('rejects non-positive width/height rather than restoring a zero-sized window', () => {
    fs.writeFileSync(
      path.join(storeDir, STATE_FILE_NAME),
      JSON.stringify({ x: 0, y: 0, width: 0, height: 800, isMaximized: false }),
      'utf-8',
    )
    expect(loadWindowState(storeDir)).toBeNull()
  })

  it('never throws even when the store directory cannot be written to', () => {
    const nested = path.join(storeDir, 'nested', 'deep', 'dir')
    expect(() => saveWindowState(nested, { x: 0, y: 0, width: 1200, height: 800, isMaximized: false })).not.toThrow()
  })
})

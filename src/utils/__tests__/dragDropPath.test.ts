import { describe, expect, it, vi } from 'vitest'

import { resolveDroppedFilePath } from '../dragDropPath'

function makeFile(name = 'doc.docx'): File {
  return new File(['content'], name)
}

describe('resolveDroppedFilePath', () => {
  it('returns null when there is no dropped file', async () => {
    const api = { getPathForFile: vi.fn() }
    expect(await resolveDroppedFilePath(api, null)).toBeNull()
    expect(api.getPathForFile).not.toHaveBeenCalled()
  })

  it('returns null when the bridge has no getPathForFile (browser mode)', async () => {
    expect(await resolveDroppedFilePath(undefined, makeFile())).toBeNull()
    expect(await resolveDroppedFilePath({}, makeFile())).toBeNull()
  })

  it('returns null when getPathForFile resolves to an empty path', async () => {
    const api = { getPathForFile: vi.fn().mockReturnValue('') }
    expect(await resolveDroppedFilePath(api, makeFile())).toBeNull()
  })

  it('returns the resolved path when there is no registration step', async () => {
    const api = { getPathForFile: vi.fn().mockReturnValue('C:\\Users\\a\\doc.docx') }
    expect(await resolveDroppedFilePath(api, makeFile())).toBe('C:\\Users\\a\\doc.docx')
  })

  it('registers the resolved path with the main-process allowlist before returning it', async () => {
    const registerDroppedPath = vi.fn().mockResolvedValue({ ok: true })
    const api = {
      getPathForFile: vi.fn().mockReturnValue('C:\\Users\\a\\doc.docx'),
      registerDroppedPath,
    }
    const path = await resolveDroppedFilePath(api, makeFile())
    expect(path).toBe('C:\\Users\\a\\doc.docx')
    expect(registerDroppedPath).toHaveBeenCalledWith('C:\\Users\\a\\doc.docx')
  })

  it('returns null when the main process rejects the path (not registered)', async () => {
    const api = {
      getPathForFile: vi.fn().mockReturnValue('C:\\Windows\\System32\\config'),
      registerDroppedPath: vi.fn().mockResolvedValue({ ok: false }),
    }
    expect(await resolveDroppedFilePath(api, makeFile())).toBeNull()
  })
})

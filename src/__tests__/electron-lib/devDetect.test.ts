import { describe, expect, it } from 'vitest'

import { resolveIsDev } from '../../../electron/lib/devDetect.cjs'

describe('resolveIsDev (RUN-12)', () => {
  it('resolves to prod when a production build exists on disk and there is no override', () => {
    expect(resolveIsDev({ atlasDevEnv: undefined, distIndexExists: () => true })).toBe(false)
  })

  it('resolves to dev when no production build exists on disk and there is no override', () => {
    expect(resolveIsDev({ atlasDevEnv: undefined, distIndexExists: () => false })).toBe(true)
  })

  it('ATLAS_DEV=1 forces dev even when a production build exists on disk', () => {
    expect(resolveIsDev({ atlasDevEnv: '1', distIndexExists: () => true })).toBe(true)
  })

  it('ATLAS_DEV=0 forces prod even when no production build exists on disk', () => {
    expect(resolveIsDev({ atlasDevEnv: '0', distIndexExists: () => false })).toBe(false)
  })

  it('ignores an unrecognized ATLAS_DEV value and falls back to the filesystem check', () => {
    expect(resolveIsDev({ atlasDevEnv: 'yes', distIndexExists: () => true })).toBe(false)
    expect(resolveIsDev({ atlasDevEnv: '', distIndexExists: () => false })).toBe(true)
  })
})

import { describe, expect, it } from 'vitest'

import { countWords } from '../textStats'

describe('countWords', () => {
  it('prefers innerText when available', () => {
    expect(countWords({ innerText: 'one two three', textContent: 'one two three four' })).toBe(3)
  })

  it('falls back to textContent when innerText is undefined (jsdom)', () => {
    expect(countWords({ innerText: undefined as unknown as string, textContent: 'one two three' })).toBe(3)
  })

  it('returns 0 for empty/whitespace-only content', () => {
    expect(countWords({ innerText: '   \n  ', textContent: '' })).toBe(0)
    expect(countWords({ innerText: undefined as unknown as string, textContent: null as unknown as string })).toBe(0)
  })
})

import { describe, expect, it } from 'vitest'

import {
  CLOSE_PROMPT_CHOICE,
  decideAfterPromptChoice,
  decideOnClose,
} from '../../../electron/lib/closeGuard.cjs'

describe('decideOnClose (P2.5/SHELL-02/ELEC-06)', () => {
  it('allows the close when the renderer reports a clean document', () => {
    expect(decideOnClose(false)).toBe('allow')
  })

  it('requires a prompt when the renderer reports unsaved changes', () => {
    expect(decideOnClose(true)).toBe('prompt')
  })
})

describe('decideAfterPromptChoice', () => {
  it('maps the Save button to "save"', () => {
    expect(decideAfterPromptChoice(CLOSE_PROMPT_CHOICE.SAVE)).toBe('save')
  })

  it('maps the Discard button to "discard"', () => {
    expect(decideAfterPromptChoice(CLOSE_PROMPT_CHOICE.DISCARD)).toBe('discard')
  })

  it('maps the Cancel button to "cancel"', () => {
    expect(decideAfterPromptChoice(CLOSE_PROMPT_CHOICE.CANCEL)).toBe('cancel')
  })

  it('treats a dismissed dialog (undefined/-1) the same as Cancel — never discard unanswered', () => {
    expect(decideAfterPromptChoice(undefined)).toBe('cancel')
    expect(decideAfterPromptChoice(-1)).toBe('cancel')
  })
})

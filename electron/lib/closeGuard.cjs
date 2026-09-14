'use strict';

/**
 * P2.5/SHELL-02/ELEC-06 — pure decision logic behind the main-process window
 * `close` handler, split out of `main.cjs` so it is testable without a real
 * `BrowserWindow`/`dialog` (Electron's own modules can't be constructed
 * outside a running app, so anything that touches them directly is
 * effectively untestable from Vitest).
 *
 * The renderer has no visible way to block a native window close on its own
 * — a `beforeunload` handler in Electron/Chromium does not surface a prompt
 * — so the renderer instead pushes its combined dirty state to main via a
 * lightweight IPC signal (`renderer:dirty-state`) whenever it changes, and
 * main decides what to do when the OS asks to close the window.
 */

/** What the main-process `close` handler should do, given the renderer's last-reported dirty state. */
const CLOSE_DECISION = Object.freeze({
  /** Not dirty (or dirty state was never reported) — let the close proceed untouched. */
  ALLOW: 'allow',
  /** Dirty — the caller must block the close and show the Save/Discard/Cancel prompt. */
  PROMPT: 'prompt',
});

/** `dialog.showMessageBoxSync`'s button-index outcomes for the close-confirmation prompt, in display order. */
const CLOSE_PROMPT_BUTTONS = Object.freeze(['Save', 'Discard', 'Cancel']);
const CLOSE_PROMPT_CHOICE = Object.freeze({ SAVE: 0, DISCARD: 1, CANCEL: 2 });

/**
 * @param {boolean} isDirty the renderer's last-reported combined dirty state
 * @returns {'allow' | 'prompt'}
 */
function decideOnClose(isDirty) {
  return isDirty ? CLOSE_DECISION.PROMPT : CLOSE_DECISION.ALLOW;
}

/**
 * Maps a `dialog.showMessageBoxSync` result (a button index, or `undefined`/
 * `-1` if the dialog itself was dismissed e.g. via the window's own close
 * button) to what should happen next. Treats an unrecognized/dismissed
 * result the same as Cancel — never treat "the user didn't answer" as
 * permission to discard their work.
 *
 * @param {number | undefined} choice
 * @returns {'save' | 'discard' | 'cancel'}
 */
function decideAfterPromptChoice(choice) {
  if (choice === CLOSE_PROMPT_CHOICE.SAVE) return 'save';
  if (choice === CLOSE_PROMPT_CHOICE.DISCARD) return 'discard';
  return 'cancel';
}

module.exports = {
  CLOSE_DECISION,
  CLOSE_PROMPT_BUTTONS,
  CLOSE_PROMPT_CHOICE,
  decideOnClose,
  decideAfterPromptChoice,
};

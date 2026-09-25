import { expect, type Page } from '@playwright/test'

/**
 * Waits until the element matching `selector` is the page's focused element.
 *
 * Not `toBeFocused()`: Playwright's version also requires
 * `document.hasFocus()` — the WINDOW having OS focus — which a freshly
 * launched Electron window on the CI runner sometimes doesn't get (it
 * reported "inactive" while the textarea was the active element; seen twice
 * in `spreadsheet-grid-fixes.spec.ts` on e2e-windows). Keystrokes from
 * Playwright go to the page's focused element either way, and every spec
 * using this goes on to assert the edit in the saved file.
 */
export async function expectDomFocused(page: Page, selector: string, timeout = 5_000): Promise<void> {
  await expect
    .poll(() => page.evaluate((sel) => document.activeElement === document.querySelector(sel), selector), {
      message: `${selector} should be the focused element`,
      timeout,
    })
    .toBe(true)
}

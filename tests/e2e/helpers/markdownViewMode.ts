import { expect, type Page } from '@playwright/test'

/**
 * Presses Ctrl+2 (split view) and waits for the markdown editor's textarea.
 *
 * REF-EFFECT-1 — on the CI runner (`76bee8d`) Ctrl+2 pressed right after the
 * markdown viewer appeared left the app in Preview, in two specs at once; it
 * never reproduced locally, even throttled and under CPU contention. If it
 * happens again, the failure reports what the app looked like instead of only
 * "element not found": where focus was (a focused field gates Ctrl+2), which
 * view the toolbar shows, and whether a later press works.
 */
export async function openMarkdownEditorWithShortcut(page: Page): Promise<void> {
  await page.keyboard.press('Control+2')
  const textarea = page.locator('.editor-panel__textarea')
  try {
    await expect(textarea).toBeVisible({ timeout: 10_000 })
  } catch (error) {
    const state = await page.evaluate(() => {
      const active = document.activeElement
      return {
        activeElement: active ? `${active.tagName}.${active.className}` : null,
        activeIsEditable: active instanceof HTMLElement && (active.isContentEditable || active.matches('input, textarea')),
        mainClass: document.querySelector('main')?.className ?? null,
        pressedViewButtons: [...document.querySelectorAll('button[aria-pressed="true"]')].map((b) => b.textContent),
        documentHasFocus: document.hasFocus(),
      }
    })
    await page.keyboard.press('Control+2')
    const secondPressWorked = await textarea.waitFor({ timeout: 3_000 }).then(
      () => true,
      () => false,
    )
    throw new Error(
      `Ctrl+2 did not open the markdown editor. App state: ${JSON.stringify(state)}; ` +
        `a second Ctrl+2 ${secondPressWorked ? 'DID' : 'did not'} open it.\n${String(error)}`,
    )
  }
}

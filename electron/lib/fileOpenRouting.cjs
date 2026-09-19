'use strict';

/**
 * P5.2/ELEC-07 — pure decision logic behind "should an incoming file-open
 * request (a second-instance argv path, or the macOS `open-file` event) be
 * pushed straight to the renderer, or queued for it to pick up later?".
 * Split out of `main.cjs` for the same reason as closeGuard.cjs/windowState.cjs:
 * it has no Electron dependency, so it's directly testable from Vitest.
 *
 * The race this closes: `mainWindow` (the `BrowserWindow` instance) can
 * exist well before the renderer has actually mounted and registered its
 * `ipcRenderer.on('file-opened-path', ...)` listener (see
 * `useFileHandler.ts`'s boot-subscriptions effect) — `webContents.send` to a
 * channel nothing is listening on yet is silently dropped, not queued. The
 * original ELEC-07 fix only covered the narrower case of `mainWindow` being
 * `null` outright (a second launch arriving before the window object even
 * exists); a second launch arriving in the gap between window creation and
 * full renderer mount was still lost.
 *
 * `rendererReady` is not a timing guess — main.cjs sets it the moment the
 * renderer's first `get-initial-file` invoke lands, and by construction the
 * renderer always registers its `file-opened-path` listener synchronously
 * *before* making that call (both happen in the same effect, listener
 * subscription first), so its arrival is a precise "the listener now
 * exists" signal rather than an arbitrary delay.
 *
 * @param {{ hasWindow: boolean, rendererReady: boolean }} state
 * @returns {'send' | 'queue'}
 */
function decideFileOpenAction({ hasWindow, rendererReady }) {
  return hasWindow && rendererReady ? 'send' : 'queue';
}

module.exports = { decideFileOpenAction };

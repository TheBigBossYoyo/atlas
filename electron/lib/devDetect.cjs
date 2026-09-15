'use strict';

/**
 * Pure dev/prod detection logic (RUN-12), split out of `main.cjs` the same
 * way closeGuard/csp/pathAllowlist/etc. are so it's unit-testable without a
 * real `app`/`BrowserWindow` (Electron's own modules can't be constructed
 * outside a running app).
 *
 * `app.isPackaged` alone used to decide this, but it only reflects whether
 * Electron was launched from an asar/packaged app — it's still `false` for
 * `npm run electron:preview` (`npm run build && electron .`), which runs the
 * unpackaged CLI against a freshly-built `dist/`. That made "preview" load
 * `http://localhost:5173` instead, silently falling back to a dev server
 * that isn't even guaranteed to be running. Keyed instead on whether a
 * production build actually exists on disk, with an explicit env override
 * for the rare case a caller wants to force one mode or the other.
 *
 * **Known nuance** (also recorded in `docs/KNOWN_LIMITATIONS.md`): because
 * this checks the filesystem rather than *how* the process was launched, a
 * `dist/` left over from an earlier `npm run build`/`electron:build` makes
 * every subsequent plain `electron .` (including a local `npx playwright
 * test` run, whose spec files launch Electron directly with no explicit
 * mode) resolve to prod against that possibly-stale build, even when the
 * intent was to exercise the live dev server. `ATLAS_DEV=1` is the explicit
 * escape hatch — CI's `e2e-windows` job never hits this because it always
 * builds fresh immediately before running the e2e suite.
 */

/**
 * @param {{ atlasDevEnv: string | undefined, distIndexExists: () => boolean }} deps
 * @returns {boolean}
 */
function resolveIsDev({ atlasDevEnv, distIndexExists }) {
  if (atlasDevEnv === '1') return true;
  if (atlasDevEnv === '0') return false;
  return !distIndexExists();
}

module.exports = { resolveIsDev };

# Release Notes

Practical notes for building and shipping the Windows installer, kept
separate from `.sisyphus/plans/atlas-phase3-improvement.md` (the planning
doc) so a future release doesn't have to dig through phase-plan history to
find them. Written after Task P5.1/P5.2 (release hardening — metadata and
Electron polish); the code-signing decision below is explicitly the owner's
to make, not something this task implemented.

## Release checklist

This is the repeatable version of what actually happened for the `3.2.0`
(`4244bdd`) and `3.3.0` (`09cbc9c`) releases — both were a single `chore(release):
<version>` commit on `main` after every wave targeted for that release had
already merged, built by the lead (not a worktree agent) directly on their
machine. Atlas does not use git tags for releases today — `main` at the
release commit *is* the release; "which commit shipped 3.3.0" is answered
by `git log --oneline -S'"version": "3.3.0"' -- package.json`.

### 1. Confirm everything targeted for this release is on `main`

- `git log --oneline -20` and cross-check against
  `.sisyphus/plans/atlas-phase3-improvement.md`'s Execution Status section
  (§13) and `.sisyphus/plans/atlas-phase3-findings-register.md` — every wave
  branch that's supposed to be in this release should already show as
  merged (`git rev-list --count main..<branch>` = 0).
- Nothing on a worktree agent's branch that hasn't been reviewed/merged.

### 2. Full verification, in order (never run these three concurrently — see `CONTRIBUTING.md`)

This is CI's own job order (`.github/workflows/ci.yml`), run locally first
so a failure is caught before it's someone else's problem on a shared
runner:

```bash
npx eslint src electron tests scripts   # 1. lint
npx tsc -b                               # 2. type-check (all 5 projects — see STRICT_MODE_TODO.md)
npm run coverage                         # 3. unit + characterization + corpus suites, with coverage
npx vite build                           # 4. renderer bundle
node scripts/check-bundle.mjs            # 5. bundle-regression gate (already run as `build`'s postbuild if
                                          #    you used `npm run build` instead of the two lines above)
npx playwright test                      # 6. Electron e2e — build first (see above); workers: 1
```

After step 6, **revert the e2e fixture corpus it rewrote** before
committing anything: `git checkout -- tests/e2e/fixtures/`.

CI also runs `e2e-windows` on every push to `main`; treat a red run there
as blocking even if everything passed locally — it's the only Windows-native
run of the real packaged-shape app path.

### 3. Bump the version

```bash
npm version <new-version> --no-git-tag-version
```

This updates `package.json` and `package-lock.json`'s top-level `version`
field only (no commit, no git tag — `--no-git-tag-version` is required
since Atlas doesn't tag releases). Follow semver by feel: a wave of new
user-facing features (new document types, editing surfaces) is a minor bump
(`3.2.0` → `3.3.0`); a wave of fixes/hardening with no new capability is a
patch bump.

### 4. Move `[Unreleased]` to a dated version in `CHANGELOG.md`

```diff
-## [Unreleased]
+## [<version>] — <YYYY-MM-DD> (<one-line summary of the wave(s) it closes>)
```

Everything already written under `[Unreleased]` becomes that version's
notes — don't rewrite it, the wave's own commits already documented what
shipped in detail. Start a fresh, empty `[Unreleased]` heading above it only
once something new lands after this release.

### 5. Update the planning docs

- `.sisyphus/plans/atlas-phase3-improvement.md`'s Execution Status (§13) —
  mark the phase/task rows this release closes as **DONE** with the
  version/commit, per its own "update it at each future phase gate" rule.
- `.sisyphus/plans/atlas-phase3-findings-register.md` — record the wave(s)
  in its per-wave resolution table (§14c-style) with commit SHAs.
- Commit all of the above (`package.json`, `package-lock.json`,
  `CHANGELOG.md`, the two plan docs) together as `chore(release): <version>`.

### 6. Build the installer

```bash
npm run electron:build   # = npm run build && electron-builder --win
```

This is a full, un-interruptible build — don't run it alongside Vitest or
Playwright (see `CONTRIBUTING.md`). Output lands at
`release/Atlas-Setup-<version>.exe` (the `artifactName` pattern in
`electron-builder.yml`'s `nsis` block). The installer is
**`perMachine: true` with `allowElevation: true`** (every install/update
needs a UAC prompt — see `DEFER-7` in the improvement plan for the
per-user-install alternative that hasn't been decided on) and **unsigned**
today (see "Code signing" below) — expect a Windows SmartScreen warning on
first run of a freshly built installer.

### 7. Compute and record the SHA256

```powershell
Get-FileHash release\Atlas-Setup-<version>.exe -Algorithm SHA256
```

Record the hash next to the release (today: `.sisyphus/HANDOFF.md`'s
"Current state" line is where this has been kept — there is no other
release-artifact ledger yet). This is the only integrity check available
for an unsigned installer: anyone who wants to confirm they received an
unmodified build has nothing else to check it against.

### 8. Smoke-test the packaged installer

Run the actual installer, not just `npm run electron:preview` (which skips
NSIS entirely) — this is the one step that catches an installer-specific
regression (a missing packaged asset, a broken file association, an icon
that didn't embed):

1. Install it (accept the UAC prompt / SmartScreen "Run anyway").
2. Launch Atlas from the Start Menu shortcut the installer created.
3. Open one file of each of a few representative formats (a `.docx`, an
   `.xlsx`, a `.pdf`) via `Ctrl+O` and via a Windows Explorer double-click
   (confirms the file association the installer registered actually routes
   to Atlas).
4. Make an edit and save it (confirms the packaged app's write path works
   outside a dev environment — different working directory, different user
   permissions than a worktree checkout).
5. Check **Properties → Details** on the installed `Atlas.exe` (default
   install path is `%ProgramFiles%\Atlas\` — `perMachine: true` in
   `electron-builder.yml`'s `nsis` block, which is why installing it needs
   the UAC elevation prompt from step 1 above; `allowToChangeInstallationDirectory:
   true` means the actual path may differ if it was changed during install)
   and confirm:
   - **File description**: `Atlas`
   - **File version** / **Product version**: matches the release (e.g. `3.3.0.0`)
   - **Product name**: `Atlas`
   - **Copyright**: `Copyright (c) 2026 Atlas`
   - **Icon**: the Atlas icon, not Electron's default
   (`src/__tests__/electron-lib/releaseMetadata.test.ts` asserts the
   *config* that produces these; this step confirms the packaged binary
   actually reflects it — see that file's own history for why the two can
   disagree: `signAndEditExecutable: false` used to silently skip this
   embedding entirely.)
6. Uninstall via **Settings → Apps** and confirm it removes cleanly (no
   leftover Start Menu entry, no orphaned registry `appId` entry beyond
   what Windows itself retains for "recently uninstalled").

### 9. Push

```bash
git push origin main
```

No tag to push — `main` at this commit is the release, per the note above.

### 10. Rolling back

Since there's no tag and no separate release branch, "rolling back" means
one of:

- **The release commit hasn't been used yet** (bad build, wrong version
  bumped): `git revert` the `chore(release): <version>` commit, or, if nothing
  has been pushed/shared yet, amend it — this repo's own git-safety rules
  still apply (never force-push `main`; never `git reset --hard` without
  confirming nothing valuable is on the discarded side).
- **A shipped release has a serious bug found after the fact**: there is no
  installed-base auto-updater (see `DEFER-7`/code-signing below), so
  "rolling back" for a user who already installed it means telling them to
  uninstall and reinstall the previous version's `.exe` — which means the
  previous release's installer must still be available somewhere (it is
  not currently archived anywhere outside whoever built it; consider
  keeping `release/Atlas-Setup-<version>.exe` for at least the last two
  versions until a proper artifact store exists). On `main`, the fix ships
  as a normal new patch release rather than editing history — reverting the
  bad commit(s) and cutting `<version>+1` is safer than trying to
  un-release something that may already be installed on the owner's
  machine.

## Executable / installer metadata

The packaged `.exe` and its NSIS installer pick up the following from
`package.json` and `electron-builder.yml`:

| Field | Source | Current value |
|---|---|---|
| ProductName / FileDescription | `electron-builder.yml`'s `productName` | `Atlas` |
| CompanyName | `package.json`'s `author.name` | `Atlas` |
| LegalCopyright | `electron-builder.yml`'s `copyright` | `Copyright (c) 2026 Atlas` |
| FileVersion / ProductVersion | `package.json`'s `version` | `3.3.0` |
| Icon | `electron-builder.yml`'s `win.icon` | `build/icon.ico` (7 sizes: 16–256px, generated by `scripts/build-icon.mjs`) |
| Uninstaller entry name (Control Panel) | `nsis.uninstallDisplayName` | `Atlas 3.3.0` |
| App id | `electron-builder.yml`'s `appId` | `com.mdreader.app` (legacy value, tracked as ELEC-22 — not changed here to avoid orphaning existing installs' registry/updater identity) |

**Verifying this without running `electron-builder`:** contributor worktrees
are asked not to run `npm run electron:build` (it fights other agents for
CPU/disk and only the lead builds the release installer), so this was
verified by:
1. Reading `node_modules/app-builder-lib/out/winPackager.js`'s `signApp()`/
   `signAndEditResources()` directly to confirm exactly which config flag
   controls resource (icon/version) embedding versus signing (see the
   `ELEC-10` comment in `electron-builder.yml`).
2. `src/__tests__/electron-lib/releaseMetadata.test.ts`, which asserts the
   yml/`package.json` fields above stay correct (e.g. that
   `signAndEditExecutable: false` — which silently disables icon/version
   embedding too, not just signing — never comes back, and that `author` is
   an object with a `.name`, not a bare string electron-builder's
   `companyName` getter can't read).
3. A full confirmation still requires the lead's next real
   `electron-builder --win` run and inspecting the resulting
   `Atlas-Setup-*.exe`'s Properties → Details tab in Windows Explorer.

### What used to be wrong (ELEC-10)

`electron-builder.yml` had `win.signAndEditExecutable: false`. That flag
doesn't just skip code signing — per `app-builder-lib`'s own `WinPackager`,
setting it to `false` returns out of `signApp()` before resource editing
(`resedit`) ever runs at all, so the exe would have shipped with **no**
custom icon and **no** version metadata, not just no signature. Fixed by
switching to `signExecutable: false` instead, which skips only the actual
signing step and leaves icon/version-metadata embedding enabled — this is
exactly what electron-builder's own log message for the old flag recommends
("To skip only code signing while keeping icon and metadata applied, use
`signExecutable: false` instead").

Separately, `package.json`'s `"author": "Atlas"` was a bare string;
electron-builder's `AppInfo.companyName` getter does `author.name`, which on
a string is `undefined` — so `CompanyName` would never have been set even
with resource editing enabled. Changed to `"author": { "name": "Atlas" }`.

## Code signing — what it would take (not done)

The owner has no code-signing certificate today, and P5.1 is explicitly
scoped to skip anything that needs one. This section is the reference for
if/when that changes.

**What's needed:**
- An **OV (Organization Validation) or EV (Extended Validation) code-signing
  certificate** from a CA (DigiCert, Sectigo, SSL.com, etc.), or a cloud
  signing service (Azure Trusted Signing, SignPath) that issues short-lived
  certs per signing operation instead of a file you hold yourself.
  - OV certs typically require proof the purchaser is a registered business
    entity (articles of incorporation, D-U-N-S number, or similar) — a solo
    developer publishing as an individual usually cannot get one issued to
    themselves as a person in the same way; most CAs only issue OV to
    organizations.
  - EV certs have the same organizational-identity requirement, are more
    expensive, and (unlike OV) are usually shipped on a hardware token
    (USB, or a cloud HSM) rather than a plain file, which also changes how
    CI would need to sign (interactive token vs. a `CSC_LINK`-style secret).
  - **Azure Trusted Signing** is the practical option for an individual/solo
    publisher today: it accepts individual identity verification (not just
    organizations) in supported countries, issues short-lived certs per
    signature automatically (no cert file to store/rotate), and is
    meaningfully cheaper than a traditional 1–3 year OV cert.
- **Cost, order of magnitude:** a traditional OV certificate from a CA runs
  roughly **$100–$400/year**; EV roughly **$300–$700/year** (both vary a lot
  by CA and multi-year discounts). Azure Trusted Signing is a
  pay-as-you-go/subscription Azure service, priced per signing
  certificate-profile rather than a flat annual cert fee — check current
  Azure pricing before committing, but it is generally the cheapest path for
  a low-volume solo release cadence.
- **What changes in `electron-builder.yml`:** `win.signExecutable: false`
  (added by this task) is removed/reverted to the default (`true`); one of
  `win.signtoolOptions` (a local/CI-provided `.pfx`/`.p12` file via
  `certificateFile` + `certificatePassword`, or `certificateSubjectName` for
  a token/HSM already installed on the signing machine) or
  `win.azureSignOptions` (for Azure Trusted Signing) is configured. CI would
  supply the certificate/credentials via secrets (electron-builder's
  conventional `CSC_LINK`/`CSC_KEY_PASSWORD` env vars for a `.pfx`, or
  Azure's own credential env vars for Trusted Signing) — never committed to
  the repo.
- **What SmartScreen does today (unsigned):** Windows SmartScreen shows a
  "Windows protected your PC" warning with an "unknown publisher" style
  message on first run of a freshly-downloaded, unsigned `Atlas-Setup-*.exe`
  until the specific file's SHA-256 has accumulated enough install/reputation
  telemetry across Microsoft's user base — for a low-download-volume
  personal tool that can take a long time or never really happen, so every
  release's installer effectively shows this warning to every new user.
  Users can bypass it via "More info" → "Run anyway", but it is a genuine
  trust/friction cost.
- **What SmartScreen does once signed:** an OV certificate does **not**
  eliminate the warning immediately — code-signed binaries still build
  reputation with SmartScreen over time (just faster and more reliably than
  unsigned ones, since the reputation is now tied to a persistent verified
  publisher identity across releases, not a single throwaway file hash). An
  **EV** certificate (or Azure Trusted Signing, which uses a similar
  identity-verification model) gets **immediate** SmartScreen reputation —
  no warning from the first release onward — which is the main practical
  reason EV/Trusted-Signing is worth the extra cost over OV for a tool meant
  to be downloaded directly by end users rather than distributed through a
  store.
- Also worth deciding at the same time (per `DEFER-3`/`DEFER-7` in the phase
  plan): whether to keep `nsis.perMachine: true` + `allowElevation: true`
  (every install/update needs a UAC prompt) or move to a per-user install,
  and whether/how an auto-updater gets wired up — both are coupled to
  whichever signing path gets chosen.

## Not covered by this document

The "Release checklist" section above covers the unsigned-build path Atlas
actually ships today (steps 1–10). Still out of scope, per Task P5.3 in the
phase plan, and requiring a signed build on real hardware once code signing
(see above) is set up: a full manual regression pass beyond the smoke test
in step 8, a dedicated signed-build checklist, and installer *upgrade-path*
verification (installing `<version>` over an existing `<version-1>` install
— `differentialPackage: false` in `electron-builder.yml` means this is
currently a full reinstall over the old location, not a binary diff, but
the actual upgrade UX — Start Menu shortcut survives, settings/recent-files
persist, no duplicate entry in Control Panel — has not been verified against
a real prior install).

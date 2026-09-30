/**
 * NEWDOC-DRAG-1 — verification with a REAL OS mouse click.
 *
 * RUN BY HAND:  npx vite build && node scripts/verify-drag-regions.mjs
 *
 * Not part of any suite and cannot be: it needs a real desktop session, it moves
 * the actual mouse pointer, and it leaves a native save dialog on screen. Run it
 * after touching anything in the Electron toolbar — `-webkit-app-region` is an
 * inherited property, so a new panel rendered in there is click-dead by default
 * and NOTHING else in this repo can tell you.
 *
 * Proven to discriminate, which is the only reason to keep it: with
 * `.app--electron .dropdown__menu { -webkit-app-region: no-drag }` removed from
 * `src/index.css` and the app rebuilt, it reports the menu still showing its six
 * items and no dialog — the exact bug as reported. With the rule present, the menu
 * closes and the save dialog appears.
 *
 * Playwright dispatches input over CDP straight into the renderer, which bypasses
 * the OS drag-region hit-testing that caused the bug — so no Playwright click can
 * tell a fixed build from a broken one. This synthesises a genuine mouse event
 * through `mouse_event` instead, which Windows routes through the same hit-test
 * path a human's mouse does.
 *
 * Coordinates are sent as absolute 0..65535 fractions of the primary display,
 * which cancels the DPI scale factor out of the arithmetic entirely: the fraction
 * is (contentOrigin + cssPoint) / displaySize in DIPs, and both sides scale alike.
 *
 * Self-calibrating: it first real-clicks the sidebar toggle, a `toolbar__nodrag`
 * control whose effect is unambiguous. If that does not toggle the sidebar the
 * coordinates are wrong, and any later "the menu item did nothing" would be
 * meaningless rather than a finding.
 */
import { execFileSync } from 'node:child_process'

import { _electron as electron } from '@playwright/test'

const CLICK_PS = (fx, fy) => `
$sig = @'
using System; using System.Runtime.InteropServices;
public class M {
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, IntPtr e);
  public const uint MOVE = 0x0001, ABS = 0x8000, LDOWN = 0x0002, LUP = 0x0004;
  public static void ClickAt(uint x, uint y) {
    mouse_event(MOVE | ABS, x, y, 0, IntPtr.Zero);
    System.Threading.Thread.Sleep(120);
    mouse_event(LDOWN | ABS, x, y, 0, IntPtr.Zero);
    System.Threading.Thread.Sleep(60);
    mouse_event(LUP | ABS, x, y, 0, IntPtr.Zero);
  }
}
'@
Add-Type -TypeDefinition $sig -ErrorAction SilentlyContinue
[M]::ClickAt(${fx}, ${fy})
`

const DIALOG_PS = `
$sig = @'
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public class DLG {
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  delegate bool EnumProc(IntPtr h, IntPtr p);
  public static List<string> D() {
    var o = new List<string>();
    EnumWindows((h, p) => {
      var c = new StringBuilder(256); GetClassNameW(h, c, 256);
      if (c.ToString() == "#32770") { var t = new StringBuilder(512); GetWindowTextW(h, t, 512); o.Add(t.ToString()); }
      return true;
    }, IntPtr.Zero);
    return o;
  }
}
'@
Add-Type -TypeDefinition $sig -ErrorAction SilentlyContinue
[DLG]::D() | ForEach-Object { $_ }
`

const ps = (script) => execFileSync('powershell', ['-NoProfile', '-Command', script], { encoding: 'utf8' }).trim()
const dialogs = () => ps(DIALOG_PS)

const app = await electron.launch({ args: ['.'], cwd: process.cwd(), env: { ...process.env, CI: '1', PLAYWRIGHT: '1' } })
const page = await app.firstWindow()
await page.waitForTimeout(3000)

const geom = await app.evaluate(({ BrowserWindow, screen }) => {
  const win = BrowserWindow.getAllWindows()[0]
  win.setBounds({ x: 60, y: 60, width: 1200, height: 800 })
  const b = win.getContentBounds()
  const d = screen.getPrimaryDisplay()
  return { content: b, display: d.size }
})
await page.waitForTimeout(800)
console.log('content bounds', geom.content, 'display', geom.display)

async function realClick(box) {
  const cx = geom.content.x + box.x + box.width / 2
  const cy = geom.content.y + box.y + box.height / 2
  const fx = Math.round((cx / geom.display.width) * 65535)
  const fy = Math.round((cy / geom.display.height) * 65535)
  ps(CLICK_PS(fx, fy))
  await page.waitForTimeout(900)
}

// --- calibration: a no-drag toolbar control with an unambiguous effect
const toggle = page.locator('.toolbar__btn--icon').first()
const before = await page.locator('.sidebar, [class*="sidebar"]').count()
await realClick((await toggle.boundingBox()))
const after = await page.locator('.sidebar, [class*="sidebar"]').count()
console.log(`CALIBRATION: sidebar elements ${before} -> ${after} (a change means real clicks are landing)`)

// --- the actual test: open the New menu, then real-click an item
await page.getByRole('button', { name: /New/ }).first().click() // CDP click just to OPEN it
await page.waitForTimeout(600)
console.log('menu open, items:', await page.locator('.dropdown__item').count())

const item = page.locator('.dropdown__item').first()
console.log('dialogs before:', dialogs() || 'NONE')
await realClick((await item.boundingBox()))
await page.waitForTimeout(2500)

const stillOpen = await page.locator('.dropdown__item').count()
const dlg = dialogs()
console.log(`RESULT: menu items still showing = ${stillOpen} (0 means the click landed)`)
console.log(`RESULT: save dialog = ${dlg || 'NONE'}`)
console.log(stillOpen === 0 && dlg !== '' ? 'PASS — a real mouse click reaches the menu item' : 'FAIL — the real click did not get through')

await app.close().catch(() => {})
process.exit(0)

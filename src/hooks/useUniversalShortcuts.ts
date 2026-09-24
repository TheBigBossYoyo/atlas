import { useCallback } from 'react';
import type { ViewMode } from '../types';
import { useShellShortcut } from './useShortcutManager';
import type { ShortcutHandler } from './shortcutManagerContext';

interface UseUniversalShortcutsOptions {
  openFile: () => void | Promise<void>;
  saveFile: () => void | Promise<unknown>;
  saveFileAs: () => void | Promise<unknown>;
  exportPrimary: () => void | Promise<unknown>;
  openExportMenu: () => void;
  /** NEW-01 — Ctrl+N opens the toolbar's "New document" menu, mirroring Ctrl+E/openExportMenu. */
  openNewMenu: () => void;
  cycleTheme: () => void;
  toggleSidebar: () => void;
  setViewMode: (mode: ViewMode) => void;
  toggleShortcuts: () => void;
  closeFile: () => void;
  isMarkdown: boolean;
}

/**
 * The app-chrome shell-global shortcuts (P2.1) — the lowest-precedence tier
 * of the centralized dispatcher (`useShortcutManager.ts`). Ctrl+S/Ctrl+Shift+S
 * (save), Ctrl+P (print/export) and Ctrl+W (close file) intentionally fire
 * even while a plain text field has focus (mirroring the pre-dispatcher
 * behavior — a user mid-edit in RawEditor should still be able to
 * save/print/close), everything else is skipped while `ctx.inPlainField` is
 * true so a modifier combo a *field itself* gives meaning to (e.g. Ctrl+B
 * toggling bold in a contentEditable rich-text surface) doesn't also trigger
 * the shell action (SHELL-08/09's root cause, now handled structurally
 * instead of per-viewer).
 *
 * FIELD-01 — Ctrl+W was missing from that carve-out list even though it has
 * the same shape as Ctrl+S/Ctrl+P: it's a shell/window-level convention (in
 * a real browser tab, Ctrl+W closes the tab regardless of focus) with no
 * competing in-field meaning, not a per-field editing command. Gating it
 * behind `inPlainField` meant typing in the markdown editor's textarea and
 * pressing Ctrl+W did nothing at all — no close, no unsaved-changes prompt —
 * since `closeFile` was never reached. SHORTCUT-FIELD-1 later found the same
 * for Ctrl+O/Ctrl+N (dead while typing in a Word document, the markdown
 * editor or the code editor), so they moved above the gate too. The rest
 * (`e`/`t`/`b`/`1`/`2`/`3`/`/`) stay gated: `b` in
 * particular *must* stay gated — Ctrl+B is the exact combo a contentEditable
 * rich-text surface (e.g. DocxViewer) natively treats as "toggle bold",
 * which is precisely what this gate exists to protect.
 */
export function useUniversalShortcuts({
  openFile,
  saveFile,
  saveFileAs,
  exportPrimary,
  openExportMenu,
  openNewMenu,
  cycleTheme,
  toggleSidebar,
  setViewMode,
  toggleShortcuts,
  closeFile,
  isMarkdown,
}: UseUniversalShortcutsOptions): void {
  const handler = useCallback<ShortcutHandler>(
    (event, ctx) => {
      const meta = event.ctrlKey || event.metaKey;
      if (!meta) return false;

      const lowerKey = event.key.toLowerCase();

      if (lowerKey === 's') {
        event.preventDefault();
        if (event.shiftKey) {
          void saveFileAs();
        } else {
          void saveFile();
        }
        return true;
      }

      if (lowerKey === 'p') {
        event.preventDefault();
        void exportPrimary();
        return true;
      }

      // FIELD-01 — see the file-header comment: Ctrl+W closes the document
      // (or raises the unsaved-changes prompt) the same way Ctrl+S/Ctrl+P
      // already do, even with a plain field/contentEditable focused.
      if (lowerKey === 'w') {
        event.preventDefault();
        closeFile();
        return true;
      }

      // SHORTCUT-FIELD-1 — open/new are file-level commands with the same
      // shape as Ctrl+S/Ctrl+W: no field gives them a meaning of its own, and
      // gating them left Ctrl+O dead while typing in any editor.
      if (lowerKey === 'o') {
        event.preventDefault();
        void openFile();
        return true;
      }

      if (lowerKey === 'n') {
        event.preventDefault();
        openNewMenu();
        return true;
      }

      if (ctx.inPlainField) return false;

      switch (lowerKey) {
        case 'e':
          event.preventDefault();
          openExportMenu();
          return true;
        case 't':
          // SHELL-17 — Ctrl+Shift+T reopens the last closed document; only the
          // unshifted combo cycles the theme.
          if (event.shiftKey) return false;
          event.preventDefault();
          cycleTheme();
          return true;
        case 'b':
          event.preventDefault();
          toggleSidebar();
          return true;
        case '1':
          if (!isMarkdown) return false;
          event.preventDefault();
          setViewMode('preview');
          return true;
        case '2':
          if (!isMarkdown) return false;
          event.preventDefault();
          setViewMode('split');
          return true;
        case '3':
          if (!isMarkdown) return false;
          event.preventDefault();
          setViewMode('editor');
          return true;
        case '/':
          event.preventDefault();
          toggleShortcuts();
          return true;
        default:
          return false;
      }
    },
    [
      closeFile,
      cycleTheme,
      exportPrimary,
      isMarkdown,
      openExportMenu,
      openNewMenu,
      openFile,
      saveFile,
      saveFileAs,
      setViewMode,
      toggleShortcuts,
      toggleSidebar,
    ],
  );

  useShellShortcut(handler);
}

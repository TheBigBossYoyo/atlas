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
 * (save) and Ctrl+P (print/export) intentionally fire even while a plain
 * text field has focus (mirroring the pre-dispatcher behavior — a user
 * mid-edit in RawEditor should still be able to save/print), everything else
 * is skipped while `ctx.inPlainField` is true so typing a shortcut's letter
 * into a text field doesn't also trigger the shell action (SHELL-08/09's
 * root cause, now handled structurally instead of per-viewer).
 */
export function useUniversalShortcuts({
  openFile,
  saveFile,
  saveFileAs,
  exportPrimary,
  openExportMenu,
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

      if (ctx.inPlainField) return false;

      switch (lowerKey) {
        case 'o':
          event.preventDefault();
          void openFile();
          return true;
        case 'w':
          event.preventDefault();
          closeFile();
          return true;
        case 'e':
          event.preventDefault();
          openExportMenu();
          return true;
        case 't':
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

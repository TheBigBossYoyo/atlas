import { useEffect } from 'react';
import type { ViewMode } from '../types';

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
  isMarkdown: boolean;
}

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
  isMarkdown,
}: UseUniversalShortcutsOptions): void {
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const meta = e.ctrlKey || e.metaKey;
      if (!meta) return;

      const inField =
        e.target instanceof HTMLInputElement ||
        e.target instanceof HTMLTextAreaElement;

      if (e.key.toLowerCase() === 's') {
        if (e.shiftKey) {
          e.preventDefault();
          void saveFileAs();
        } else {
          e.preventDefault();
          void saveFile();
        }
        return;
      }

      if (e.key.toLowerCase() === 'p') {
        e.preventDefault();
        void exportPrimary();
        return;
      }

      if (inField) return;

      switch (e.key.toLowerCase()) {
        case 'o':
          e.preventDefault();
          void openFile();
          break;
        case 'e':
          e.preventDefault();
          openExportMenu();
          break;
        case 't':
          e.preventDefault();
          cycleTheme();
          break;
        case 'b':
          e.preventDefault();
          toggleSidebar();
          break;
        case '1':
          if (!isMarkdown) break;
          e.preventDefault();
          setViewMode('preview');
          break;
        case '2':
          if (!isMarkdown) break;
          e.preventDefault();
          setViewMode('split');
          break;
        case '3':
          if (!isMarkdown) break;
          e.preventDefault();
          setViewMode('editor');
          break;
        case '/':
          e.preventDefault();
          toggleShortcuts();
          break;
        default:
          break;
      }
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [
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
  ]);
}

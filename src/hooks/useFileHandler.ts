import { useState, useCallback, useEffect, useRef } from 'react';
import type { LoadedFile, FormatId } from '../formats/types';
import { detectByExtension, detectByMagic, detectFormat } from '../formats/detect';
import type { RecentFile } from '../types';

// ---------------------------------------------------------------------------
// Format class sets — single source of truth for routing decisions
// ---------------------------------------------------------------------------

const TEXT_CLASS_FORMATS = new Set<FormatId>([
  'markdown', 'csv', 'tsv', 'text', 'code',
]);

const BINARY_CLASS_FORMATS = new Set<FormatId>([
  'pdf', 'docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'rtf',
]);

const BROWSER_MODE_ERROR =
  'This feature requires the Atlas desktop app — file access is unavailable in a plain browser tab.';

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export interface FileHandlerState {
  file: LoadedFile | null;
  loading: boolean;
  error: string | null;
  /**
   * Increments every time a load successfully commits a new `file` (P2.10).
   * Combined with `file.path`, it gives callers a stable "which load is this"
   * identity even when the same path is reopened, so a dirty-reset guard can
   * key off file *identity* instead of comparing derived content strings.
   */
  loadGeneration: number;
}

export interface FileHandlerActions {
  openDialog: () => Promise<void>;
  loadFromPath: (absPath: string) => Promise<void>;
  clear: () => void;
  /** Dismiss the current error without discarding the currently-open file (P2.3). */
  clearError: () => void;
}

export type UseFileHandlerReturn = FileHandlerState & FileHandlerActions & {
  // ---------------------------------------------------------------------------
  // Compatibility shim — exposes legacy markdown-centric fields so that
  // existing callers (App.tsx etc.) continue to compile until W3.5 migrates
  // them to consume `file` directly.
  // ---------------------------------------------------------------------------
  markdown: string;
  fileName: string | null;
  filePath: string;
};

export interface UseFileHandlerOptions {
  /**
   * P2.4 — the single `useRecentFiles()` instance lives in App.tsx; this hook
   * no longer keeps its own, so a removed recent entry can't be silently
   * resurrected by a stale internal copy.
   */
  addRecent: (file: Omit<RecentFile, 'openedAt'>) => void;
  /**
   * P1.1 — called before every open (dialog result, Recent click, drag-drop,
   * or the OS "Open with" IPC event — all of which funnel through
   * `loadFromPath`) so the caller can show a Save/Discard/Cancel prompt when
   * there are unsaved changes. Resolving `false` aborts the open with no
   * state change. Omit to allow every open unconditionally (e.g. in tests
   * that don't exercise the guard).
   */
  confirmDiscardChanges?: () => Promise<boolean>;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useFileHandler({
  addRecent,
  confirmDiscardChanges,
}: UseFileHandlerOptions): UseFileHandlerReturn {
  const [file, setFile] = useState<LoadedFile | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [loadGeneration, setLoadGeneration] = useState(0);

  // Guard so getInitialFile only fires once even in StrictMode double-invoke
  const initialFileFetched = useRef(false);

  // Monotonically increasing request id (P2.10) — lets a slower, earlier
  // load discover it's been superseded and discard its own result instead of
  // overwriting a faster, newer one.
  const requestIdRef = useRef(0);

  // -------------------------------------------------------------------------
  // Core loader
  // -------------------------------------------------------------------------

  const loadFromPath = useCallback(async (absPath: string): Promise<void> => {
    // P2.12/SHELL-20/ELEC-19/QA-26 — running outside Electron (e.g. `npm run
    // dev` in a plain browser tab) must fail with a friendly, surfaced error
    // instead of throwing on a non-null assertion.
    if (typeof window === 'undefined' || !window.electronAPI) {
      setError(BROWSER_MODE_ERROR);
      return;
    }
    const electronAPI = window.electronAPI;

    // P1.1 — guard every open path behind the combined dirty check. This is
    // the single funnel every entry point (dialog, Recent click, drag-drop,
    // and the OS "Open with" IPC subscription below) already calls through,
    // so gating here covers all of them at once.
    if (confirmDiscardChanges) {
      const canProceed = await confirmDiscardChanges();
      if (!canProceed) {
        return;
      }
    }

    const requestId = ++requestIdRef.current;
    setLoading(true);
    setError(null);

    try {
      const extFormat = detectByExtension(absPath);

      let loaded: LoadedFile;

      if (TEXT_CLASS_FORMATS.has(extFormat)) {
        const data = await electronAPI.openFileByPath(absPath);
        if (!data) throw new Error(`Failed to read file: ${absPath}`);
        loaded = {
          kind: 'text',
          content: data.content,
          path: absPath,
          format: extFormat,
        };
      } else if (BINARY_CLASS_FORMATS.has(extFormat)) {
        const data = await electronAPI.readBinaryByPath(absPath);
        loaded = {
          kind: 'binary',
          content: data.buffer,
          path: absPath,
          format: extFormat,
        };
      } else {
        // Unknown extension — read binary then use magic-byte detection
        const data = await electronAPI.readBinaryByPath(absPath);
        const magicFormat = detectByMagic(data.buffer);
        const finalFormat = detectFormat(absPath, data.buffer);

        if (TEXT_CLASS_FORMATS.has(finalFormat)) {
          const content = new TextDecoder('utf-8').decode(data.buffer);
          loaded = {
            kind: 'text',
            content,
            path: absPath,
            format: finalFormat,
          };
        } else {
          loaded = {
            kind: 'binary',
            content: data.buffer,
            path: absPath,
            // finalFormat may still be 'unknown' if magic also failed
            format: magicFormat !== 'unknown' ? magicFormat : finalFormat,
          };
        }
      }

      // A newer load already started while we were awaiting IPC — discard
      // this now-stale result instead of clobbering the newer one (P2.10).
      if (requestIdRef.current !== requestId) {
        return;
      }

      setFile(loaded);
      setLoadGeneration(g => g + 1);

      // Push to recents — name derived from path tail
      const name = absPath.split(/[\\/]/).pop() ?? absPath;
      addRecent({ path: absPath, name });
    } catch (err) {
      if (requestIdRef.current !== requestId) {
        return;
      }
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      if (requestIdRef.current === requestId) {
        setLoading(false);
      }
    }
  }, [addRecent, confirmDiscardChanges]);

  // -------------------------------------------------------------------------
  // Dialog opener
  // -------------------------------------------------------------------------

  const openDialog = useCallback(async (): Promise<void> => {
    if (typeof window === 'undefined' || !window.electronAPI) {
      setError(BROWSER_MODE_ERROR);
      return;
    }

    // NOTE: openFileBinary() returns the buffer alongside the path, but we
    // deliberately route through loadFromPath so all detection logic stays in
    // one place. The second IPC read is acceptable overhead for now.
    const result = await window.electronAPI.openFileBinary();
    if (result.canceled) return;
    await loadFromPath(result.path);
  }, [loadFromPath]);

  // -------------------------------------------------------------------------
  // Clear
  // -------------------------------------------------------------------------

  const clear = useCallback(() => {
    setFile(null);
    setError(null);
    setLoading(false);
  }, []);

  const clearError = useCallback(() => {
    setError(null);
  }, []);

  // -------------------------------------------------------------------------
  // Boot subscriptions (mount once)
  // -------------------------------------------------------------------------

  useEffect(() => {
    if (typeof window === 'undefined' || !window.electronAPI) return;

    // Subscribe to OS "Open with" / menu-bar open events
    const unsub = window.electronAPI.onFileOpenedPath((path) => {
      void loadFromPath(path);
    });

    // Load the file that was passed as a CLI argument (argv path)
    if (!initialFileFetched.current) {
      initialFileFetched.current = true;
      void window.electronAPI.getInitialFile().then((res) => {
        if (res && 'path' in res && res.path) {
          void loadFromPath(res.path);
        }
      });
    }

    return unsub;
  }, [loadFromPath]);

  // -------------------------------------------------------------------------
  // Compatibility shim — derived from `file` for legacy callers
  // -------------------------------------------------------------------------

  const markdown =
    file?.kind === 'text' && file.format === 'markdown' ? file.content : '';
  const fileName =
    file ? (file.path.split(/[\\/]/).pop() ?? null) : null;
  const filePath = file?.path ?? '';

  return {
    file,
    loading,
    error,
    loadGeneration,
    openDialog,
    loadFromPath,
    clear,
    clearError,
    // Legacy shim
    markdown,
    fileName,
    filePath,
  };
}

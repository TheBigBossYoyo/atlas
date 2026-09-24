import { useState, useCallback, useEffect, useRef } from 'react';
import type { LoadedFile, FormatId } from '../formats/types';
import { detectByExtension, detectByMagic, looksLikeText } from '../formats/detect';
import { decodeTextBufferWithMeta, setTextFileMeta } from '../utils/textDecoding';
import type { RecentFile } from '../types';
import { t } from '../i18n';

// ---------------------------------------------------------------------------
// Format class sets — single source of truth for routing decisions
// ---------------------------------------------------------------------------

const TEXT_CLASS_FORMATS = new Set<FormatId>([
  'markdown', 'csv', 'tsv', 'text', 'code',
]);

const BINARY_CLASS_FORMATS = new Set<FormatId>([
  'pdf', 'docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'rtf', 'doc', 'ppt',
]);

// i18n — functions (not string constants) so each call reflects whatever the
// UI language is AT THE TIME the error actually occurs, not whatever it was
// when this module first evaluated.
function browserModeError(): string {
  return t('errors.browserModeUnavailable');
}

// UX — `open-file-by-path` (main process) resolves to `null`, rather than
// throwing, when the target no longer exists on disk (a stale "Recent"
// entry, or the file being deleted/moved out from under Atlas while it was
// open elsewhere). Mirrors `FILE_NOT_FOUND_MESSAGE` in `electron/main.cjs`'s
// `file:readBinaryByPath` handler (the binary-class equivalent of this same
// failure) instead of the old `Failed to read file: ${absPath}`, which read
// as a raw dev-facing string and leaked the full filesystem path into the UI.
function fileNotFoundError(): string {
  return t('errors.fileNotFound');
}

// P2.14/LOAD-04 — human-readable labels for the "extension vs. actual
// contents disagree" confirm prompt below. Falls back to the bare FormatId
// for anything not worth a friendlier label (translated the same way as
// `browserModeError`/`fileNotFoundError` above — resolved at call time, not
// module load, so it always reflects the UI language at the moment the
// prompt actually appears).
const FORMAT_LABEL_KEYS: Partial<Record<FormatId, string>> = {
  docx: 'fileHandler.formatLabel.docx',
  xlsx: 'fileHandler.formatLabel.xlsx',
  pptx: 'fileHandler.formatLabel.pptx',
  pdf: 'fileHandler.formatLabel.pdf',
  odt: 'fileHandler.formatLabel.odt',
  ods: 'fileHandler.formatLabel.ods',
  odp: 'fileHandler.formatLabel.odp',
  rtf: 'fileHandler.formatLabel.rtf',
  doc: 'fileHandler.formatLabel.doc',
  ppt: 'fileHandler.formatLabel.ppt',
};

function formatLabel(format: FormatId): string {
  const key = FORMAT_LABEL_KEYS[format];
  return key !== undefined ? t(key) : format;
}

/**
 * P2.14/LOAD-04 — the recognized-extension "fast path" used to trust the
 * extension blindly and hand the bytes straight to a format-specific
 * parser, which could throw an unfriendly low-level error on a renamed or
 * corrupted file. Verifies the magic bytes actually agree with the
 * extension for binary-class formats; on a confident mismatch, asks before
 * proceeding instead of silently mis-parsing (or failing outright).
 *
 * Returns `true` when the load should proceed (no mismatch, or the user
 * chose "open anyway"), `false` when it should be silently aborted.
 */
function confirmMagicMatchesExtension(absPath: string, extFormat: FormatId, buffer: ArrayBuffer): boolean {
  const magicFormat = detectByMagic(buffer);

  // No confident magic result, or it agrees with the extension — nothing to warn about.
  if (magicFormat === 'unknown' || magicFormat === extFormat) {
    return true;
  }

  if (typeof window === 'undefined' || typeof window.confirm !== 'function') {
    // No way to ask the user — fail safe by refusing rather than silently mis-parsing.
    return false;
  }

  const name = absPath.split(/[\\/]/).pop() ?? absPath;
  return window.confirm(
    t('fileHandler.extensionMismatchConfirm', {
      name,
      extFormat: formatLabel(extFormat),
      magicFormat: formatLabel(magicFormat),
    }),
  );
}

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
  /**
   * SHELL-17 — shows a file whose bytes the caller already holds (switching
   * to another open document), with no disk read and no unsaved-changes
   * prompt of its own: the caller owns that decision when it switches.
   */
  adopt: (loaded: LoadedFile) => void;
  /**
   * @param prefetchedBuffer P4.10/LOAD-13 — when the caller already has the
   * file's bytes (e.g. from the Open dialog's own read), pass them here to
   * skip a redundant second disk read/IPC round trip.
   */
  loadFromPath: (absPath: string, prefetchedBuffer?: ArrayBuffer) => Promise<void>;
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

  const loadFromPath = useCallback(async (absPath: string, prefetchedBuffer?: ArrayBuffer): Promise<void> => {
    // P2.12/SHELL-20/ELEC-19/QA-26 — running outside Electron (e.g. `npm run
    // dev` in a plain browser tab) must fail with a friendly, surfaced error
    // instead of throwing on a non-null assertion.
    if (typeof window === 'undefined' || !window.electronAPI) {
      setError(browserModeError());
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
        if (prefetchedBuffer) {
          // P4.10/LOAD-13 — a dialog-based open already has the bytes (the
          // Open dialog's own IPC handler reads the file to return it
          // alongside the chosen path); decode locally instead of asking
          // main to read the same file from disk a second time.
          //
          // NIGHT/text-roundtrip — `decodeTextBufferWithMeta` also detects
          // the file's encoding/BOM/newline convention so a later save can
          // reproduce it (SHELL-1/SHELL-2); recorded in the path-keyed
          // registry since `LoadedFile` itself carries no such field.
          const { content, meta } = decodeTextBufferWithMeta(prefetchedBuffer);
          setTextFileMeta(absPath, meta);
          loaded = {
            kind: 'text',
            content,
            path: absPath,
            format: extFormat,
          };
        } else {
          const data = await electronAPI.openFileByPath(absPath);
          if (!data) throw new Error(fileNotFoundError());
          // NIGHT/text-roundtrip — main's `readMarkdownFile` already detected
          // `meta` (see `electron/main.cjs`); record it the same way as the
          // prefetched-buffer branch above.
          if (data.meta) setTextFileMeta(absPath, data.meta);
          loaded = {
            kind: 'text',
            content: data.content,
            path: absPath,
            format: extFormat,
          };
        }
      } else if (BINARY_CLASS_FORMATS.has(extFormat)) {
        const buffer = prefetchedBuffer ?? (await electronAPI.readBinaryByPath(absPath)).buffer;

        // A newer load may have started while we were awaiting the read —
        // don't bother prompting for a load the user has already superseded.
        if (requestIdRef.current !== requestId) {
          return;
        }

        if (!confirmMagicMatchesExtension(absPath, extFormat, buffer)) {
          return;
        }

        loaded = {
          kind: 'binary',
          content: buffer,
          path: absPath,
          format: extFormat,
        };
      } else {
        // Unknown extension — read the bytes, then try magic-byte detection
        // and finally a plain-text sniff before giving up as genuinely
        // unrecognized.
        const buffer = prefetchedBuffer ?? (await electronAPI.readBinaryByPath(absPath)).buffer;
        const magicFormat = detectByMagic(buffer);

        if (magicFormat !== 'unknown') {
          loaded = {
            kind: 'binary',
            content: buffer,
            path: absPath,
            format: magicFormat,
          };
        } else if (looksLikeText(buffer)) {
          // P2.11/LOAD-10 — an extensionless plain-text file (README,
          // LICENSE, Dockerfile, .gitignore) gets a real text viewer
          // instead of the generic "unknown format" empty state.
          // NIGHT/text-roundtrip — same metadata tracking as the TEXT_CLASS
          // branch above, so this file's Ctrl+S round-trips too.
          const { content, meta } = decodeTextBufferWithMeta(buffer);
          setTextFileMeta(absPath, meta);
          loaded = {
            kind: 'text',
            content,
            path: absPath,
            format: 'text',
          };
        } else {
          loaded = {
            kind: 'binary',
            content: buffer,
            path: absPath,
            format: 'unknown',
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
      setError(browserModeError());
      return;
    }
    const electronAPI = window.electronAPI;

    // P4.10/LOAD-13 — openFileBinary() itself performs the (potentially
    // slow) file read; surface the same loading/error state around it that
    // loadFromPath manages for every other open path, instead of a silent
    // pause followed by handing an already-read buffer to loadFromPath (so
    // detection/decoding still happens in one place, without a second IPC
    // round trip re-reading the same file from disk).
    const requestId = ++requestIdRef.current;
    setLoading(true);
    setError(null);

    let result: Awaited<ReturnType<typeof electronAPI.openFileBinary>>;
    try {
      result = await electronAPI.openFileBinary();
    } catch (err) {
      if (requestIdRef.current === requestId) {
        setError(err instanceof Error ? err.message : String(err));
        setLoading(false);
      }
      return;
    }

    // A newer request (another openDialog/loadFromPath call) started while
    // the dialog/read was in flight — let it own the loading/error state.
    if (requestIdRef.current !== requestId) {
      return;
    }

    if (result.canceled) {
      setLoading(false);
      return;
    }

    // Hand off to loadFromPath, which manages its own loading/error/request-id
    // lifecycle from here — including the confirmDiscardChanges guard, which
    // can decline without ever touching loading, so reset it first rather
    // than risk leaving the indicator stuck on.
    setLoading(false);
    await loadFromPath(result.path, result.buffer);
  }, [loadFromPath]);

  // -------------------------------------------------------------------------
  // Clear
  // -------------------------------------------------------------------------

  const adopt = useCallback((loaded: LoadedFile) => {
    setError(null);
    setLoading(false);
    setFile(loaded);
    setLoadGeneration((generation) => generation + 1);
  }, []);

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
    adopt,
    clear,
    clearError,
    // Legacy shim
    markdown,
    fileName,
    filePath,
  };
}

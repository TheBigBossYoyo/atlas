import { useState, useCallback, useEffect, useRef } from 'react';
import type { LoadedFile, FormatId } from '../formats/types';
import { detectByExtension, detectByMagic, detectFormat } from '../formats/detect';
import { useRecentFiles } from './useRecentFiles';

// ---------------------------------------------------------------------------
// Format class sets — single source of truth for routing decisions
// ---------------------------------------------------------------------------

const TEXT_CLASS_FORMATS = new Set<FormatId>([
  'markdown', 'csv', 'tsv', 'text', 'code',
]);

const BINARY_CLASS_FORMATS = new Set<FormatId>([
  'pdf', 'docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'rtf',
]);

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export interface FileHandlerState {
  file: LoadedFile | null;
  loading: boolean;
  error: string | null;
}

export interface FileHandlerActions {
  openDialog: () => Promise<void>;
  loadFromPath: (absPath: string) => Promise<void>;
  clear: () => void;
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

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useFileHandler(): UseFileHandlerReturn {
  const [file, setFile] = useState<LoadedFile | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Guard so getInitialFile only fires once even in StrictMode double-invoke
  const initialFileFetched = useRef(false);

  const { addRecent } = useRecentFiles();

  // -------------------------------------------------------------------------
  // Core loader
  // -------------------------------------------------------------------------

  const loadFromPath = useCallback(async (absPath: string): Promise<void> => {
    setLoading(true);
    setError(null);

    try {
      const extFormat = detectByExtension(absPath);

      let loaded: LoadedFile;

      if (TEXT_CLASS_FORMATS.has(extFormat)) {
        const data = await window.electronAPI!.openFileByPath(absPath);
        if (!data) throw new Error(`Failed to read file: ${absPath}`);
        loaded = {
          kind: 'text',
          content: data.content,
          path: absPath,
          format: extFormat,
        };
      } else if (BINARY_CLASS_FORMATS.has(extFormat)) {
        const data = await window.electronAPI!.readBinaryByPath(absPath);
        loaded = {
          kind: 'binary',
          content: data.buffer,
          path: absPath,
          format: extFormat,
        };
      } else {
        // Unknown extension — read binary then use magic-byte detection
        const data = await window.electronAPI!.readBinaryByPath(absPath);
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

      setFile(loaded);

      // Push to recents — name derived from path tail
      const name = absPath.split(/[\\/]/).pop() ?? absPath;
      addRecent({ path: absPath, name });
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, [addRecent]);

  // -------------------------------------------------------------------------
  // Dialog opener
  // -------------------------------------------------------------------------

  const openDialog = useCallback(async (): Promise<void> => {
    // NOTE: openFileBinary() returns the buffer alongside the path, but we
    // deliberately route through loadFromPath so all detection logic stays in
    // one place. The second IPC read is acceptable overhead for now.
    const result = await window.electronAPI!.openFileBinary();
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
    openDialog,
    loadFromPath,
    clear,
    // Legacy shim
    markdown,
    fileName,
    filePath,
  };
}

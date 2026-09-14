import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import { ViewerRouter } from './components/ViewerRouter';
import { useTheme } from './hooks/useTheme';
import { useFileHandler } from './hooks/useFileHandler';
import { useSearch } from './hooks/useSearch';
import { useToc } from './hooks/useToc';
import { useRecentFiles } from './hooks/useRecentFiles';
import { useFontSize } from './hooks/useFontSize';
import { useAutosave } from './hooks/useAutosave';
import { useUniversalShortcuts } from './hooks/useUniversalShortcuts';
import { Toolbar } from './components/Toolbar';
import { Sidebar } from './components/Sidebar';
import { MarkdownRenderer } from './components/MarkdownRenderer';
import { RawEditor } from './components/RawEditor';
import { SearchOverlay } from './components/SearchOverlay';
import { DropZone } from './components/DropZone';
import { FileStatusBanner } from './components/FileStatusBanner';
import { UnsavedChangesDialog } from './components/UnsavedChangesDialog';
import { WelcomeScreen } from './components/WelcomeScreen';
import { StatusBar } from './components/StatusBar';
import { ShortcutsModal } from './components/ShortcutsModal';
import { SAMPLE_MARKDOWN } from './constants';
import { THEMES, type ViewMode, type ExportFormat, type RecentFile } from './types';
import { exportMarkdown, exportHtml, exportPdf, exportDocx } from './utils/export';
import type { LoadedFile, NavItem } from './formats/types';
import { resolveDroppedFilePath } from './utils/dragDropPath';
import { ViewerProvider } from './viewers/shared/ViewerContext';
import { useSetNavItems, useSetViewerStats, useViewerIsDirty, useViewerSave } from './viewers/shared/useViewerContext';

const isElectron = typeof window !== 'undefined' && !!window.electronAPI;

const ENABLE_VIEWER_ROUTER = true;

interface MarkdownChromeBridgeProps {
  navItems: readonly NavItem[];
  words: number;
  headings: number;
}

function MarkdownChromeBridge({ navItems, words, headings }: MarkdownChromeBridgeProps) {
  const setNavItems = useSetNavItems();
  const setViewerStats = useSetViewerStats();

  useEffect(() => {
    setNavItems(navItems);
  }, [navItems, setNavItems]);

  useEffect(() => {
    setViewerStats({ kind: 'markdown', words, headings });
  }, [headings, setViewerStats, words]);

  return null;
}

interface ViewerSessionBridgeProps {
  onDirtyChange: (dirty: boolean) => void;
  saveRef: React.MutableRefObject<() => Promise<boolean>>;
}

/**
 * P1.1 — lifts the active viewer's document-session state (from ViewerContext,
 * only reachable inside <ViewerProvider>) up to App.tsx: `viewerDirty` state
 * for Toolbar/StatusBar and the unsaved-changes guard, and a stable ref to the
 * registered `save()` so App's global Ctrl+S/Save and the guard's own "Save"
 * button can reach whichever viewer is actually active.
 */
function ViewerSessionBridge({ onDirtyChange, saveRef }: ViewerSessionBridgeProps) {
  const dirty = useViewerIsDirty();
  const save = useViewerSave();

  useEffect(() => {
    onDirtyChange(dirty);
  }, [dirty, onDirtyChange]);

  useEffect(() => {
    saveRef.current = save;
  }, [save, saveRef]);

  return null;
}

function triggerBinaryDownload(content: ArrayBuffer, fileName: string): void {
  const url = URL.createObjectURL(new Blob([content]));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  window.setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function App() {
  const { theme, setTheme, cycleTheme } = useTheme();
  const { recent, addRecent, removeRecent } = useRecentFiles();
  const { increase, decrease, reset } = useFontSize();

  // P1.1 — document-session guard state. `confirmDiscardChanges` must exist
  // before `useFileHandler()` is called (it's passed in as an option), but
  // its own dirty check depends on state that isn't computed until later in
  // this function (isMarkdownDocument / isDirty / viewerDirty). It reads that
  // state from a ref instead of closing over it directly, so its identity can
  // stay stable (empty deps) while still always seeing the latest values —
  // the ref is kept in sync a little further down, once those values exist.
  const dirtyGuardStateRef = useRef({ isMarkdownDocument: false, isDirty: false, viewerDirty: false });
  const pendingConfirmResolveRef = useRef<((proceed: boolean) => void) | null>(null);
  const viewerSaveRef = useRef<() => Promise<boolean>>(async () => false);
  const [viewerDirty, setViewerDirty] = useState(false);
  const [unsavedDialogOpen, setUnsavedDialogOpen] = useState(false);
  const [unsavedDialogSaving, setUnsavedDialogSaving] = useState(false);
  const [unsavedDialogError, setUnsavedDialogError] = useState<string | null>(null);

  const confirmDiscardChanges = useCallback((): Promise<boolean> => {
    const { isMarkdownDocument, isDirty, viewerDirty } = dirtyGuardStateRef.current;
    const currentlyDirty = isMarkdownDocument ? isDirty : viewerDirty;
    if (!currentlyDirty) {
      return Promise.resolve(true);
    }

    // A confirmation is already showing (e.g. a second open request — the OS
    // "file-opened-path" event can fire independently of anything the guard's
    // own modal is blocking) — refuse the new request rather than silently
    // overwriting `pendingConfirmResolveRef` and orphaning the first caller's
    // Promise forever (it would otherwise never resolve, since only one
    // resolver can be stored at a time).
    if (pendingConfirmResolveRef.current) {
      return Promise.resolve(false);
    }

    return new Promise<boolean>((resolve) => {
      pendingConfirmResolveRef.current = resolve;
      setUnsavedDialogError(null);
      setUnsavedDialogOpen(true);
    });
  }, []);

  const {
    file,
    markdown,
    fileName,
    filePath,
    loading,
    error,
    loadGeneration,
    openDialog: openFile,
    loadFromPath: openFileFromPath,
    clearError,
  } = useFileHandler({ addRecent, confirmDiscardChanges });

  // Derive isDirty / setMarkdown / saveFile / saveFileAs / drag handlers from
  // legacy autosave hook until W3.5 migrates App.tsx fully to LoadedFile.
  // For now markdown content comes from the shim (file.kind==='text' && format==='markdown').
  const [localMarkdown, setLocalMarkdown] = useState<string>(() => markdown ?? '');
  const [isDirty, setIsDirty] = useState(false);
  const [isDragging, setIsDragging] = useState(false);

  // P2.10/SHELL-05/LOAD-07 — reset keyed on file *identity* (path + a
  // load-generation token bumped by every successful load), not on comparing
  // the derived markdown string. String-equality comparison silently failed
  // to reset dirty/content whenever the newly-opened file's content happened
  // to match whatever was last synced (e.g. switching away from a dirty
  // Sample session to a freshly-opened binary file — SHELL-05's stale dirty
  // dot). The token also means reopening the exact same path still counts as
  // a fresh load.
  const fileIdentityKey = file ? `${file.path}#${loadGeneration}` : null;
  const [lastSyncedFileKey, setLastSyncedFileKey] = useState<string | null>(null);

  if (fileIdentityKey !== lastSyncedFileKey) {
    setLastSyncedFileKey(fileIdentityKey);
    setLocalMarkdown(markdown);
    setIsDirty(false);
  }

  const setMarkdown = useCallback((value: string) => {
    setLocalMarkdown(value);
    setIsDirty(true);
  }, []);

  const isMarkdownDocument = useMemo(
    () => (file ? file.format === 'markdown' : localMarkdown.length > 0),
    [file, localMarkdown.length]
  );

  // S16 — PptxViewer/OdpViewer's SlideDeck already renders its own thumbnail
  // rail; showing the app-level Sidebar too would duplicate slide navigation.
  const isSlidesDocument = useMemo(
    () => file?.format === 'pptx' || file?.format === 'odp',
    [file]
  );

  // P1.1 — the combined dirty signal Toolbar/StatusBar render and the
  // unsaved-changes guard checks: markdown's own isDirty when a markdown
  // document is active, otherwise whatever the active viewer has reported
  // through the shared document-session contract.
  const combinedDirty = isMarkdownDocument ? isDirty : viewerDirty;

  // Keep confirmDiscardChanges's ref-read state fresh. An effect (not a
  // render-time assignment) so this never mutates a ref during render.
  useEffect(() => {
    dirtyGuardStateRef.current = { isMarkdownDocument, isDirty, viewerDirty };
  }, [isMarkdownDocument, isDirty, viewerDirty]);

  const saveFile = useCallback(async (): Promise<boolean> => {
    if (!isMarkdownDocument) {
      // P1.1/SHELL-10/DXE-07/RUN-03 — route the global Save/Ctrl+S to
      // whichever viewer is actually active via the shared capability
      // contract, instead of silently no-op-ing for every non-markdown format.
      return viewerSaveRef.current();
    }
    if (!window.electronAPI) return false;
    const result = await window.electronAPI.saveFile({
      content: localMarkdown,
      suggestedName: fileName ?? 'document.md',
      existingPath: filePath || undefined,
    });
    if (result.saved) { setIsDirty(false); return true; }
    return false;
  }, [fileName, filePath, isMarkdownDocument, localMarkdown]);

  const saveFileAs = useCallback(async (): Promise<boolean> => {
    if (!isMarkdownDocument) return false;
    if (!window.electronAPI) return false;
    const result = await window.electronAPI.saveFile({
      content: localMarkdown,
      suggestedName: fileName ?? 'document.md',
    });
    if (result.saved) { setIsDirty(false); return true; }
    return false;
  }, [fileName, isMarkdownDocument, localMarkdown]);

  // Counts nested dragenter/dragleave pairs so the overlay only hides once
  // the drag has actually left the window, not merely a child element
  // (SHELL-24 — the classic target-identity check goes false-negative on
  // nested children).
  const dragCounterRef = useRef(0);

  const handleUnsavedDialogCancel = useCallback(() => {
    pendingConfirmResolveRef.current?.(false);
    pendingConfirmResolveRef.current = null;
    setUnsavedDialogOpen(false);
    setUnsavedDialogError(null);
  }, []);

  const handleUnsavedDialogDiscard = useCallback(() => {
    pendingConfirmResolveRef.current?.(true);
    pendingConfirmResolveRef.current = null;
    setUnsavedDialogOpen(false);
    setUnsavedDialogError(null);
  }, []);

  const handleUnsavedDialogSave = useCallback(async () => {
    setUnsavedDialogSaving(true);
    setUnsavedDialogError(null);
    try {
      const saved = isMarkdownDocument ? await saveFile() : await viewerSaveRef.current();
      if (saved) {
        pendingConfirmResolveRef.current?.(true);
        pendingConfirmResolveRef.current = null;
        setUnsavedDialogOpen(false);
      } else {
        setUnsavedDialogError('Save failed. Discard your changes or cancel and try saving again.');
      }
    } finally {
      setUnsavedDialogSaving(false);
    }
  }, [isMarkdownDocument, saveFile]);

  const handleDragEnter = useCallback((e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
    dragCounterRef.current += 1;
    setIsDragging(true);
  }, []);
  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
    dragCounterRef.current = Math.max(0, dragCounterRef.current - 1);
    if (dragCounterRef.current === 0) setIsDragging(false);
  }, []);
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
  }, []);
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
    dragCounterRef.current = 0;
    setIsDragging(false);
    void resolveDroppedFilePath(window.electronAPI, e.dataTransfer.files[0]).then((path) => {
      if (path) void openFileFromPath(path);
    });
  }, [openFileFromPath]);

  useAutosave(localMarkdown, fileName);

  const [viewMode, setViewMode] = useState<ViewMode>('preview');
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [exportMenuOpen, setExportMenuOpen] = useState(false);

  const contentRef = useRef<HTMLDivElement>(null);
  const tocItems = useToc(localMarkdown);
  const markdownWordCount = useMemo(
    () => localMarkdown.trim().split(/\s+/).filter(Boolean).length,
    [localMarkdown]
  );
  const markdownNavItems = useMemo<readonly NavItem[]>(
    () =>
      tocItems.map((item) => ({
        id: item.id,
        label: item.text,
        level: item.level,
        onSelect: () => {
          document.getElementById(item.id)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
        },
      })),
    [tocItems]
  );
  const hasContent = useMemo(
    () => (isMarkdownDocument ? localMarkdown.length > 0 : file !== null),
    [file, isMarkdownDocument, localMarkdown.length]
  );
  const currentFormat = useMemo(() => {
    if (file) {
      return file.format;
    }

    return 'markdown' as const;
  }, [file]);
  const viewerFile = useMemo<LoadedFile | null>(
    () => (!isMarkdownDocument && file ? file : null),
    [file, isMarkdownDocument]
  );
  const canSave = useMemo(() => isMarkdownDocument && hasContent, [hasContent, isMarkdownDocument]);

  const {
    isOpen: searchOpen,
    query: searchQuery,
    setQuery: setSearchQuery,
    matchCount,
    currentMatch,
    goToMatch,
    open: openSearch,
    close: closeSearch,
  } = useSearch(contentRef, isMarkdownDocument);

  const loadSample = useCallback(() => {
    setMarkdown(SAMPLE_MARKDOWN);
  }, [setMarkdown]);

  const handleOpenRecent = useCallback(
    (file: RecentFile) => {
      if (file.path && isElectron) {
        // Recent-file paths are renderer-persisted (localStorage), so they
        // must be re-validated and re-registered with the main process's
        // read allowlist (P1.2) before loadFromPath's IPC reads will accept
        // them — a file removed or moved since it was remembered is simply
        // not reopened rather than surfacing a raw IPC rejection.
        const requestOpen = window.electronAPI?.requestOpenRecent;
        if (requestOpen) {
          void requestOpen(file.path).then((result) => {
            if (result.ok) void openFileFromPath(file.path);
          });
        } else {
          void openFileFromPath(file.path);
        }
      } else {
        void openFile();
      }
    },
    [openFile, openFileFromPath]
  );

  const handleRemoveRecent = useCallback(
    (key: string) => {
      removeRecent(key);
    },
    [removeRecent]
  );

  const handleExport = useCallback(
    async (format: ExportFormat) => {
      const baseName = fileName ?? 'document.md';
      try {
        if (!isMarkdownDocument && file?.kind === 'binary' && file.format === 'pdf' && format === 'pdf') {
          triggerBinaryDownload(file.content, baseName);
          return;
        }

        if (!isMarkdownDocument) {
          await exportPdf('viewer-content', baseName);
          return;
        }

        switch (format) {
          case 'md':
            await exportMarkdown(localMarkdown, baseName);
            break;
          case 'html':
            await exportHtml(localMarkdown, baseName, theme);
            break;
          case 'pdf':
            await exportPdf('markdown-content', baseName);
            break;
          case 'docx':
            await exportDocx(localMarkdown, baseName);
            break;
        }
      } catch (err) {
        console.error('Export failed:', err);
        alert(`Export failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    },
    [file, fileName, isMarkdownDocument, localMarkdown, theme]
  );

  useUniversalShortcuts({
    openFile,
    saveFile,
    saveFileAs,
    exportPrimary: () => handleExport('pdf'),
    openExportMenu: () => setExportMenuOpen(true),
    cycleTheme,
    toggleSidebar: () => setSidebarOpen(prev => !prev),
    setViewMode,
    toggleShortcuts: () => setShortcutsOpen(prev => !prev),
    isMarkdown: isMarkdownDocument,
  });

  // Warn before unloading with unsaved changes
  useEffect(() => {
    if (!combinedDirty) return;
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [combinedDirty]);

  return (
    <div
      className={`app ${isElectron ? 'app--electron' : ''}`}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      onDragOver={handleDragOver}
      onDrop={handleDrop}
    >
      <Toolbar
        theme={theme}
        themes={THEMES}
        viewMode={viewMode}
        sidebarOpen={sidebarOpen}
        fileName={fileName}
        isDirty={combinedDirty}
        hasContent={hasContent}
        canSave={canSave}
        canSearch={isMarkdownDocument}
        canChangeFontSize={isMarkdownDocument}
        isMarkdown={isMarkdownDocument}
        isElectron={isElectron}
        exportFormat={currentFormat}
        exportMenuOpen={exportMenuOpen}
        onSelectTheme={setTheme}
        onViewModeChange={setViewMode}
        onToggleSidebar={() => setSidebarOpen(prev => !prev)}
        onOpenFile={openFile}
        onSave={() => void saveFile()}
        onExport={(fmt) => void handleExport(fmt)}
        onOpenSearch={openSearch}
        onShowShortcuts={() => setShortcutsOpen(true)}
        onIncreaseFont={increase}
        onDecreaseFont={decrease}
        onResetFont={reset}
        onExportMenuOpenChange={setExportMenuOpen}
      />

      <FileStatusBanner loading={loading} error={error} onDismissError={clearError} />

      {isMarkdownDocument ? (
        <SearchOverlay
          isOpen={searchOpen}
          query={searchQuery}
          matchCount={matchCount}
          currentMatch={currentMatch}
          onQueryChange={setSearchQuery}
          onNext={() => goToMatch('next')}
          onPrev={() => goToMatch('prev')}
          onClose={closeSearch}
        />
      ) : null}

      <ViewerProvider filePath={filePath || null}>
        <ViewerSessionBridge onDirtyChange={setViewerDirty} saveRef={viewerSaveRef} />
        <div className="app__body">
          {hasContent && !isSlidesDocument && <Sidebar isOpen={sidebarOpen} />}

          {!hasContent ? (
            <WelcomeScreen
              onOpenFile={openFile}
              onLoadSample={loadSample}
              recent={recent}
              onOpenRecent={handleOpenRecent}
              onRemoveRecent={handleRemoveRecent}
            />
          ) : isMarkdownDocument ? (
            <>
              <MarkdownChromeBridge
                navItems={markdownNavItems}
                words={markdownWordCount}
                headings={tocItems.length}
              />
              <main className={`content content--${viewMode}`}>
                {(viewMode === 'editor' || viewMode === 'split') && (
                  <RawEditor markdown={localMarkdown} onChange={setMarkdown} />
                )}
                {(viewMode === 'preview' || viewMode === 'split') && (
                  <div className="preview-panel" data-viewer={file?.format === 'markdown' ? 'markdown' : undefined}>
                    <MarkdownRenderer
                      ref={contentRef}
                      markdown={localMarkdown}
                      searchQuery={searchOpen ? searchQuery : undefined}
                    />
                  </div>
                )}
              </main>
            </>
          ) : viewerFile && ENABLE_VIEWER_ROUTER ? (
            <main className="content content--viewer">
              <div className="preview-panel" id="viewer-content">
                <ViewerRouter file={viewerFile} />
              </div>
            </main>
          ) : null}
        </div>

        {hasContent && <StatusBar fileName={fileName} isDirty={combinedDirty} />}
      </ViewerProvider>

      <ShortcutsModal isOpen={shortcutsOpen} onClose={() => setShortcutsOpen(false)} />

      <UnsavedChangesDialog
        isOpen={unsavedDialogOpen}
        isSaving={unsavedDialogSaving}
        errorMessage={unsavedDialogError}
        onSave={() => void handleUnsavedDialogSave()}
        onDiscard={handleUnsavedDialogDiscard}
        onCancel={handleUnsavedDialogCancel}
      />

      <DropZone isVisible={isDragging} />

      {/* Suppress unused-var warnings for filePath while keeping it part of the contract */}
      {filePath && <span style={{ display: 'none' }} aria-hidden="true">{filePath}</span>}
    </div>
  );
}

export default App;

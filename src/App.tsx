import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import { ViewerRouter } from './components/ViewerRouter';
import { useTheme } from './hooks/useTheme';
import { useFileHandler } from './hooks/useFileHandler';
import { useSearch } from './hooks/useSearch';
import { useToc } from './hooks/useToc';
import { useRecentFiles } from './hooks/useRecentFiles';
import { useFontSize } from './hooks/useFontSize';
import { useAutosave, loadDraft, clearDraft, type Draft } from './hooks/useAutosave';
import { useUniversalShortcuts } from './hooks/useUniversalShortcuts';
import { ShortcutManagerProvider } from './hooks/ShortcutManagerProvider';
import { DraftRecoveryBanner } from './components/DraftRecoveryBanner';
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
import type { FormatId, LoadedFile, NavItem } from './formats/types';
import { resolveDroppedFilePath } from './utils/dragDropPath';
import { ViewerProvider } from './viewers/shared/ViewerContext';
import { useSetNavItems, useSetViewerStats, useViewerIsDirty, useViewerSave } from './viewers/shared/useViewerContext';

const isElectron = typeof window !== 'undefined' && !!window.electronAPI;

const ENABLE_VIEWER_ROUTER = true;

// P2.1/X4/DAT-15/RUN-08 — non-markdown formats whose viewer renders into the
// DOM as real text (not a canvas or a virtualized grid), so the existing
// TreeWalker-based `useSearch` can walk it. PDF/spreadsheets/slide decks are
// deliberately excluded — see `useSearch`'s own doc comment. Known limit:
// TextViewer virtualizes very large files (`react-window`), so a search can
// only find matches in currently-mounted rows near the viewport, the same
// inherent limitation any DOM-walking search has against a virtualized list.
const SEARCHABLE_VIEWER_FORMATS = new Set<FormatId>(['text', 'code', 'rtf', 'odt']);

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

/**
 * P2.1 — the actual app shell. Split out from the default-exported `App`
 * below purely so it can sit *inside* `<ShortcutManagerProvider>`: every
 * `useShellShortcut`/`useViewerShortcuts` call (here, in child hooks like
 * `useUniversalShortcuts`/`useFontSize`/`useSearch`, and in components like
 * `ThemeMenu`/`ExportMenu`/`ShortcutsModal`/`UnsavedChangesDialog`/
 * `DocxViewer`) needs to be a descendant of the provider, and a component
 * cannot consume the context it itself creates in its own returned JSX.
 */
function AppShell() {
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
  // Surfaces a friendly reason via FileStatusBanner when a standalone
  // Save/Save As (Toolbar button, Ctrl+S, Ctrl+Shift+S — as opposed to the
  // one already shown inline in UnsavedChangesDialog) fails, instead of the
  // save silently doing nothing.
  const [saveError, setSaveError] = useState<string | null>(null);

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
    clear,
    clearError,
  } = useFileHandler({ addRecent, confirmDiscardChanges });

  // Derive isDirty / setMarkdown / saveFile / saveFileAs / drag handlers from
  // legacy autosave hook until W3.5 migrates App.tsx fully to LoadedFile.
  // For now markdown content comes from the shim (file.kind==='text' && format==='markdown').
  const [localMarkdown, setLocalMarkdown] = useState<string>(() => markdown ?? '');
  const [isDirty, setIsDirty] = useState(false);
  const [isDragging, setIsDragging] = useState(false);

  // P2.6/SHELL-11/LOAD-20 — a leftover markdown draft from a crash/force-quit
  // (autosave writes one every 800ms but nothing ever read it back). Checked
  // once on mount; the user decides to restore it into an untitled markdown
  // session or discard it. Declared here (ahead of the file-identity reset
  // block just below, which dismisses it once a real file loads) rather than
  // its more natural home near the other draft-recovery handlers further
  // down, to avoid a temporal-dead-zone reference to `setPendingDraft`.
  const [pendingDraft, setPendingDraft] = useState<Draft | null>(() => loadDraft());

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
    setSaveError(null);
    // A real file just loaded (not the initial mount, since fileIdentityKey
    // and lastSyncedFileKey both start `null` and this branch never runs for
    // that first render) — dismiss any pending draft-recovery prompt rather
    // than showing it alongside newly-opened content; the draft itself stays
    // in storage in case the user wants it back another time.
    if (fileIdentityKey !== null) {
      setPendingDraft(null);
    }
  }

  const setMarkdown = useCallback((value: string) => {
    setLocalMarkdown(value);
    setIsDirty(true);
  }, []);

  const isMarkdownDocument = useMemo(
    () => (file ? file.format === 'markdown' : localMarkdown.length > 0),
    [file, localMarkdown.length]
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
    // No `existingPath` means `save-file` shows a native Save dialog — a
    // `{ saved: false }` result with no `error` there is an ordinary user
    // Cancel, not a failure worth surfacing. With an `existingPath`, no
    // dialog is shown at all, so the same result always means a genuine
    // write failure (surfaced with a generic fallback when main didn't
    // supply a more specific reason).
    const hadExistingPath = Boolean(filePath);
    const result = await window.electronAPI.saveFile({
      content: localMarkdown,
      suggestedName: fileName ?? 'document.md',
      existingPath: filePath || undefined,
    });
    if (result.saved) {
      setIsDirty(false);
      setSaveError(null);
      clearDraft();
      return true;
    }
    if (result.error) {
      setSaveError(result.error);
    } else if (hadExistingPath) {
      setSaveError('Failed to save the file. Please try again.');
    }
    return false;
  }, [fileName, filePath, isMarkdownDocument, localMarkdown]);

  const saveFileAs = useCallback(async (): Promise<boolean> => {
    if (!isMarkdownDocument) return false;
    if (!window.electronAPI) return false;
    const result = await window.electronAPI.saveFile({
      content: localMarkdown,
      suggestedName: fileName ?? 'document.md',
    });
    if (result.saved) {
      setIsDirty(false);
      setSaveError(null);
      clearDraft();
      return true;
    }
    // A dialog-based save with no specific `error` is an ordinary user
    // Cancel — only a specifically-reported reason is worth surfacing here.
    if (result.error) {
      setSaveError(result.error);
    }
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

  // P2.8/SHELL-16 — `clear()` was already fully implemented but never wired
  // to anything. Gated behind the same unsaved-changes guard every other
  // open/close path uses.
  const closeFile = useCallback(async (): Promise<void> => {
    const canProceed = await confirmDiscardChanges();
    if (!canProceed) return;
    clear();
  }, [clear, confirmDiscardChanges]);

  const handleRestoreDraft = useCallback(() => {
    if (!pendingDraft) return;
    setMarkdown(pendingDraft.markdown);
    setPendingDraft(null);
  }, [pendingDraft, setMarkdown]);

  const handleDiscardDraft = useCallback(() => {
    clearDraft();
    setPendingDraft(null);
  }, []);

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

  // P2.6/SHELL-12 — gated on `isMarkdownDocument` so switching to a binary
  // file can never keep writing markdown drafts under its name (a stale
  // `localMarkdown`/`fileName` pairing was possible for one render during a
  // file-identity transition before this gate existed).
  useAutosave(isMarkdownDocument ? localMarkdown : '', fileName, isMarkdownDocument);

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

  // P2.1/X4/DAT-15/RUN-08 — in-app search now also covers Text/Code/RTF/ODT
  // (fully-rendered DOM, compatible with the existing TreeWalker search),
  // not just markdown. PDF/spreadsheets/slide decks are deliberately left
  // out (canvas/virtualized-grid content the TreeWalker can't see) — Ctrl+F
  // simply isn't registered for them, so it's never swallowed with no
  // feature behind it.
  const canSearchFormat = useMemo(
    () => isMarkdownDocument || (file !== null && SEARCHABLE_VIEWER_FORMATS.has(file.format)),
    [file, isMarkdownDocument],
  );

  const {
    isOpen: searchOpen,
    query: searchQuery,
    setQuery: setSearchQuery,
    matchCount,
    currentMatch,
    goToMatch,
    open: openSearch,
    close: closeSearch,
  } = useSearch(contentRef, canSearchFormat);

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
    closeFile: () => void closeFile(),
    isMarkdown: isMarkdownDocument,
  });

  // P2.8/SHELL-18 — the title bar never reflected the open file or its dirty
  // state, unlike every comparable desktop editor.
  useEffect(() => {
    document.title = fileName ? `${combinedDirty ? '● ' : ''}${fileName} — Atlas` : 'Atlas';
  }, [combinedDirty, fileName]);

  // P2.5/SHELL-02/ELEC-06 — push the combined dirty signal to main so its
  // window `close` handler knows whether to block the close behind a native
  // Save/Discard/Cancel prompt (Electron shows no visible confirmation for a
  // renderer-only `beforeunload` handler, kept below only as a browser-tab-mode
  // fallback).
  useEffect(() => {
    window.electronAPI?.notifyDirtyState?.(combinedDirty);
  }, [combinedDirty]);

  // Main asks the renderer to save (the user chose "Save" in that native
  // prompt) and waits for the result before deciding whether to actually
  // close the window.
  useEffect(() => {
    return window.electronAPI?.onRequestSaveBeforeClose?.(() => {
      void saveFile().then((saved) => {
        window.electronAPI?.reportSaveBeforeCloseResult?.({ saved });
      });
    });
  }, [saveFile]);

  // Warn before unloading with unsaved changes (browser-tab-mode fallback —
  // the packaged Electron app is guarded by the main-process `close` handler
  // above instead, since Chromium/Electron never surfaces a visible prompt
  // for this handler alone).
  useEffect(() => {
    if (!combinedDirty) return;
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [combinedDirty]);

  const handleDismissStatusError = useCallback(() => {
    if (error) {
      clearError();
      return;
    }
    setSaveError(null);
  }, [clearError, error]);

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
        canSearch={canSearchFormat}
        canChangeFontSize={isMarkdownDocument}
        canClose={hasContent}
        isMarkdown={isMarkdownDocument}
        isElectron={isElectron}
        exportFormat={currentFormat}
        exportMenuOpen={exportMenuOpen}
        onSelectTheme={setTheme}
        onViewModeChange={setViewMode}
        onToggleSidebar={() => setSidebarOpen(prev => !prev)}
        onOpenFile={openFile}
        onSave={() => void saveFile()}
        onCloseFile={() => void closeFile()}
        onExport={(fmt) => void handleExport(fmt)}
        onOpenSearch={openSearch}
        onShowShortcuts={() => setShortcutsOpen(true)}
        onIncreaseFont={increase}
        onDecreaseFont={decrease}
        onResetFont={reset}
        onExportMenuOpenChange={setExportMenuOpen}
      />

      <FileStatusBanner loading={loading} error={error ?? saveError} onDismissError={handleDismissStatusError} />

      <DraftRecoveryBanner draft={pendingDraft} onRestore={handleRestoreDraft} onDiscard={handleDiscardDraft} />

      {canSearchFormat ? (
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
          {hasContent && <Sidebar isOpen={sidebarOpen} />}

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
              {/* P2.1/X4 — the same `contentRef` the markdown preview uses for
                  its TreeWalker search is attached here too (only the branch
                  that's actually mounted ever owns it) so Text/Code/RTF/ODT
                  get the identical search behavior markdown already has. */}
              <div className="preview-panel" id="viewer-content" ref={canSearchFormat ? contentRef : undefined}>
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

/**
 * P2.1 — wraps `AppShell` in the centralized shortcut dispatcher's provider.
 * See `AppShell`'s own doc comment for why this can't just be one component.
 */
function App() {
  return (
    <ShortcutManagerProvider>
      <AppShell />
    </ShortcutManagerProvider>
  );
}

export default App;

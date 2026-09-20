import { useState, useRef, useCallback, useEffect, useMemo, lazy, Suspense } from 'react';
import { ViewerRouter } from './components/ViewerRouter';
import { ViewerLoading } from './components/ViewerLoading';
import { useTheme } from './hooks/useTheme';
import { useFileHandler } from './hooks/useFileHandler';
import { useShellShortcut } from './hooks/useShortcutManager';
import { TabBar } from './components/TabBar';
import {
  EMPTY_SESSIONS,
  activateSession,
  activeSession,
  closeSession,
  moveSession,
  neighbourSession,
  openSession,
  renameSession,
  takeLastClosedPath,
  type DocumentSessionsState,
  reloadSessionFile,
} from './session/documentSessions';
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
import { assertNever, type FormatId, type LoadedFile, type NavItem } from './formats/types';
import type { NewDocumentFormat } from './electron';
import { resolveDroppedFilePath } from './utils/dragDropPath';
import { EmptyFileNotice } from './components/EmptyFileNotice';
import { LARGE_MARKDOWN_PREVIEW_BYTES, LargeMarkdownNotice } from './components/LargeMarkdownNotice';
import { ViewerProvider } from './viewers/shared/ViewerContext';
import {
  useSetNavItems,
  useSetViewerStats,
  useViewerIsDirty,
  useViewerSave,
  useViewerSaveAs,
  useGetExportableContent,
} from './viewers/shared/useViewerContext';
import type { ExportableContent } from './viewers/shared/viewerContextValue';
import { ToastProvider } from './components/ToastProvider';
import { LocaleProvider, useTranslate } from './i18n';
import { translateWriteError } from './i18n/translateWriteError';
import { useToast } from './hooks/useToast';
import { scrollIntoViewRespectingMotionPreference } from './utils/motionPreference';

const isElectron = typeof window !== 'undefined' && !!window.electronAPI;

// PERF-01 — react-markdown's rehype/remark pipeline (rehype-katex ->
// katex, rehype-highlight -> highlight.js + its bundled languages,
// rehype-raw -> parse5/hast-util-raw, dompurify, marked's own runtime
// weight, etc.) previously sat behind a plain top-level `import` here, so
// it was pulled into the SAME eagerly-loaded entry chunk as `main.tsx`
// itself, on cold start, for every format Atlas opens (measured: this
// dependency graph alone was >1MB of the ~2.55MB entry chunk). Every other
// per-format viewer is already code-split through `formats/registry.ts`'s
// `lazy()` loaders — this gives the markdown editor/preview split the same
// treatment instead of special-casing it as the one eager exception. The
// `<Suspense>` fallback below reuses `ViewerLoading`, the same "Loading
// markdown…" spinner every other viewer already shows while its own chunk
// fetches, so this is not a new UI state, just the existing one applied
// consistently.
const MarkdownRenderer = lazy(() =>
  import('./components/MarkdownRenderer').then((m) => ({ default: m.MarkdownRenderer }))
);

// PERF-01 — `./utils/export` transitively pulls in `docx` (real export
// build, not just its types), `xlsx` (via spreadsheetPdf's shared parser),
// `html2canvas-pro`, and `react-dom/server` (slidesPdf's
// `renderToStaticMarkup`) — together over 1.5MB of the entry chunk this
// used to sit in via a plain top-level `import { exportXxx, ... }`, even
// though every one of these functions only ever runs from inside the
// user-triggered Export menu handlers below. Loaded on demand instead;
// the browser/ESM module cache means the dynamic import only actually
// fetches once, so repeat exports pay no extra cost after the first. (In
// Vitest specifically, that first call is transformed on demand rather than
// pre-bundled — see App.export.test.tsx's own comment on why its `waitFor`s
// use a longer timeout, not a shorter one from some new slowness in the app
// itself.)
const loadExportUtils = () => import('./utils/export');

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
  saveAsRef: React.MutableRefObject<() => Promise<boolean>>;
  exportContentRef: React.MutableRefObject<() => ExportableContent | null>;
  /** Re-derive the exportable-content format whenever the active file
   * changes — see the effect below for why this can't just key off
   * `getExportableContent`'s own identity. */
  filePath: string | null;
  onExportableFormatChange: (format: string | null) => void;
}

/**
 * P1.1 — lifts the active viewer's document-session state (from ViewerContext,
 * only reachable inside <ViewerProvider>) up to App.tsx: `viewerDirty` state
 * for Toolbar/StatusBar and the unsaved-changes guard, a stable ref to the
 * registered `save()` so App's global Ctrl+S/Save and the guard's own "Save"
 * button can reach whichever viewer is actually active, and (UX-12) the same
 * for `getExportableContent` so a real per-format CSV export can be offered
 * once a viewer registers one.
 */
function ViewerSessionBridge({
  onDirtyChange,
  saveRef,
  saveAsRef,
  exportContentRef,
  filePath,
  onExportableFormatChange,
}: ViewerSessionBridgeProps) {
  const dirty = useViewerIsDirty();
  const save = useViewerSave();
  const saveAs = useViewerSaveAs();
  const getExportableContent = useGetExportableContent();

  useEffect(() => {
    onDirtyChange(dirty);
  }, [dirty, onDirtyChange]);

  useEffect(() => {
    saveRef.current = save;
  }, [save, saveRef]);

  useEffect(() => {
    saveAsRef.current = saveAs;
  }, [saveAs, saveAsRef]);

  useEffect(() => {
    exportContentRef.current = getExportableContent;
    // `getExportableContent`'s own identity is stable even when a newly
    // mounted viewer's registration changes what it would return (mirroring
    // how `save` above reads a ref rather than changing identity itself), so
    // this also re-derives whenever the open file changes — the point at
    // which a different viewer would actually be registering something new.
    onExportableFormatChange(getExportableContent()?.format ?? null);
  }, [exportContentRef, filePath, getExportableContent, onExportableFormatChange]);

  return null;
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
  const t = useTranslate();
  const { theme, setTheme, cycleTheme } = useTheme();
  const { recent, addRecent, removeRecent } = useRecentFiles();
  const { increase, decrease, reset } = useFontSize();
  const showToast = useToast();

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
  // Mirrors `viewerSaveRef`, for the "Save As" side of the document-session
  // contract — lets the global Ctrl+Shift+S / `saveFileAs()` reach whichever
  // non-markdown viewer is active, the same way `viewerSaveRef` already lets
  // plain Ctrl+S reach it.
  const viewerSaveAsRef = useRef<() => Promise<boolean>>(async () => false);
  // UX-12 — the active viewer's registered `getExportableContent` (a
  // Phase-3 placeholder today; see viewerContextValue.ts), lifted the same
  // way `viewerSaveRef` lifts `save`. `exportableContentFormat` mirrors just
  // its `format` field into reactive state so the Export menu can decide
  // whether to offer a real "Export to CSV" item.
  const exportContentRef = useRef<() => ExportableContent | null>(() => null);
  const [exportableContentFormat, setExportableContentFormat] = useState<string | null>(null);
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
    adopt: adoptFile,
    clear,
    clearError,
  } = useFileHandler({ addRecent, confirmDiscardChanges });

  // SHELL-17 — every open document, and which one is showing. The loaded
  // bytes live here, so switching documents never re-reads the disk; only the
  // active one is mounted, and the shell's existing Save/Discard/Cancel
  // prompt runs before switching away from unsaved changes.
  const [sessions, setSessions] = useState<DocumentSessionsState>(EMPTY_SESSIONS);

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
  // Markdown parsing is superlinear and runs on the main thread, so a very
  // large document would freeze the app for minutes. Its preview waits for an
  // explicit request; the editor pane is unaffected. Opting in is per
  // document — a new file starts guarded again.
  const [renderLargePreview, setRenderLargePreview] = useState(false);
  const previewTooLarge = !renderLargePreview && localMarkdown.length > LARGE_MARKDOWN_PREVIEW_BYTES;

  const fileIdentityKey = file ? `${file.path}#${loadGeneration}` : null;
  const [lastSyncedFileKey, setLastSyncedFileKey] = useState<string | null>(null);

  // Review fix — `saveFile`/`saveFileAs` close over the tab that was active
  // when the save *started*; `window.electronAPI.saveFile(...)` can take a
  // while, and the user is free to switch tabs (Discard through the unsaved-
  // changes guard) before it resolves. The continuation needs to know, at
  // resolution time, whether the tab it saved is still the one showing — a
  // ref (not the `fileIdentityKey` closed over at call time) so it always
  // reads the *latest* identity rather than the one from whichever render
  // kicked off the save.
  const fileIdentityKeyRef = useRef<string | null>(fileIdentityKey);
  useEffect(() => {
    fileIdentityKeyRef.current = fileIdentityKey;
  }, [fileIdentityKey]);

  if (fileIdentityKey !== lastSyncedFileKey) {
    setLastSyncedFileKey(fileIdentityKey);
    setLocalMarkdown(markdown);
    setIsDirty(false);
    setSaveError(null);
    setRenderLargePreview(false);
    // SHELL-17 — a newly loaded (or re-activated) file becomes the showing tab.
    if (file) setSessions((current) => openSession(current, file));
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

  // SHELL-17 — after a save to a new path the document lives there: its tab,
  // the title bar and every later Ctrl+S follow it instead of still naming
  // (and silently overwriting) the old file.
  const followSavedPath = useCallback(
    (savedPath: string | undefined): void => {
      if (!savedPath || file?.kind !== 'text' || savedPath === file.path) return;
      const renamed: LoadedFile = { ...file, path: savedPath, content: localMarkdown };
      setSessions((current) => renameSession(current, file.path, renamed));
      adoptFile(renamed);
    },
    [adoptFile, file, localMarkdown],
  );

  // Review fix — the inactive-tab counterpart of `followSavedPath`. A save
  // that finishes after the user switched away must still update *that
  // document's* tab (so reactivating it later shows the right path), but
  // must never touch the tab that's showing now: no `adoptFile` (which would
  // yank the view back to the just-saved document) and no dirty/draft state
  // belonging to whatever the user switched to.
  const applySavedPathToInactiveTab = useCallback(
    (savedFile: LoadedFile, savedPath: string | undefined, savedContent: string): void => {
      if (!savedPath || savedFile.kind !== 'text' || savedPath === savedFile.path) return;
      const renamed: LoadedFile = { ...savedFile, path: savedPath, content: savedContent };
      setSessions((current) => renameSession(current, savedFile.path, renamed));
    },
    [],
  );

  const saveFile = useCallback(async (): Promise<boolean> => {
    if (!isMarkdownDocument) {
      // P1.1/SHELL-10/DXE-07/RUN-03 — route the global Save/Ctrl+S to
      // whichever viewer is actually active via the shared capability
      // contract, instead of silently no-op-ing for every non-markdown format.
      return viewerSaveRef.current();
    }
    if (!window.electronAPI) return false;
    // Snapshot which document and content this save is actually for — the
    // await below can outlive the tab that started it (review fix, see
    // `fileIdentityKeyRef`).
    const savingFile = file;
    const savingIdentityKey = fileIdentityKey;
    const savingContent = localMarkdown;
    // No `existingPath` means `save-file` shows a native Save dialog — a
    // `{ saved: false }` result with no `error` there is an ordinary user
    // Cancel, not a failure worth surfacing. With an `existingPath`, no
    // dialog is shown at all, so the same result always means a genuine
    // write failure (surfaced with a generic fallback when main didn't
    // supply a more specific reason).
    const hadExistingPath = Boolean(filePath);
    const result = await window.electronAPI.saveFile({
      content: savingContent,
      suggestedName: fileName ?? 'document.md',
      existingPath: filePath || undefined,
    });
    const stillShowing = fileIdentityKeyRef.current === savingIdentityKey;
    if (result.saved) {
      if (stillShowing) {
        setIsDirty(false);
        setSaveError(null);
        clearDraft();
        followSavedPath(result.path);
      } else if (savingFile) {
        applySavedPathToInactiveTab(savingFile, result.path, savingContent);
      }
      return true;
    }
    if (stillShowing) {
      if (result.error) {
        // X5 — main.cjs now classifies every save failure class (permission
        // denied, disk full, missing/renamed parent folder, locked file, ...)
        // into a friendly `error` string (previously only a file lock did),
        // surfaced here via the same FileStatusBanner every other save/load
        // error already uses (a second, redundant toast for the identical
        // event would leave two `role="alert"` regions on screen for one
        // failure — worse accessibility, not better).
        setSaveError(translateWriteError(t, result));
      } else if (hadExistingPath) {
        setSaveError(t('errors.write.saveFailed'));
      }
    }
    return false;
  }, [applySavedPathToInactiveTab, file, fileIdentityKey, fileName, filePath, followSavedPath, isMarkdownDocument, localMarkdown, t]);

  const saveFileAs = useCallback(async (): Promise<boolean> => {
    if (!isMarkdownDocument) {
      // P1.1/SHELL-10/DXE-07/RUN-03's fix for Ctrl+S ("route the global
      // Save to whichever viewer is actually active" — see `saveFile`
      // above) never got a Save-As counterpart, so the documented global
      // Ctrl+Shift+S silently did nothing for every non-markdown format.
      // Mirrors `saveFile`'s own non-markdown branch.
      return viewerSaveAsRef.current();
    }
    if (!window.electronAPI) return false;
    const savingFile = file;
    const savingIdentityKey = fileIdentityKey;
    const savingContent = localMarkdown;
    const result = await window.electronAPI.saveFile({
      content: savingContent,
      suggestedName: fileName ?? 'document.md',
    });
    const stillShowing = fileIdentityKeyRef.current === savingIdentityKey;
    if (result.saved) {
      if (stillShowing) {
        setIsDirty(false);
        setSaveError(null);
        clearDraft();
        followSavedPath(result.path);
      } else if (savingFile) {
        applySavedPathToInactiveTab(savingFile, result.path, savingContent);
      }
      return true;
    }
    // A dialog-based save with no specific `error` is an ordinary user
    // Cancel — only a specifically-reported reason is worth surfacing here.
    if (stillShowing && result.error) {
      setSaveError(translateWriteError(t, result));
    }
    return false;
  }, [applySavedPathToInactiveTab, file, fileIdentityKey, fileName, followSavedPath, isMarkdownDocument, localMarkdown, t]);
  // (viewerSaveAsRef is a stable ref identity, so it's intentionally left
  // out of the dependency array above, matching viewerSaveRef's usage in
  // saveFile.)

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

  // SHELL-17 — showing another open document re-reads it from disk: saves only
  // ever write to disk, so the bytes kept in its session can predate its last
  // save. Only the latest request lands when switches overlap.
  const showRequestRef = useRef(0);
  const showSessionFile = useCallback(
    async (stored: LoadedFile): Promise<void> => {
      const request = ++showRequestRef.current;
      const current = await reloadSessionFile(stored, window.electronAPI);
      if (request === showRequestRef.current) adoptFile(current);
    },
    [adoptFile],
  );

  // A non-markdown viewer (DOCX, spreadsheet, slides) that saved to a new
  // path reports it here: its tab, the title bar and every later save follow
  // the document to the file it was actually written to. Without this the tab
  // kept pointing at the file it was opened from, so coming back to it re-read
  // that original file over the user's work and the next Ctrl+S wrote to
  // neither file.
  //
  // SAVE-1 — the write this is reporting can resolve well after the user
  // switched to a different tab (nothing awaits it), so trusting the ambient
  // `file`/`filePath` closure to mean "the document that was saved" is wrong:
  // it means "whatever is showing right now". `startedFromPath` (the path
  // the viewer had open when its own save began — see `useReportSavedPath`'s
  // doc comment) is used instead to look up the actual session that was
  // saved. This mirrors the markdown save path's own
  // `followSavedPath`/`applySavedPathToInactiveTab` split: the showing tab
  // only gets `adoptFile`'d (via `showSessionFile`) when it is genuinely the
  // document that was saved; any other tab is just renamed in place — no
  // `adoptFile`, no re-read, no dirty/content change to whatever the user
  // has since switched to.
  const handleViewerSavedPath = useCallback(
    (startedFromPath: string, savedPath: string): void => {
      if (!savedPath || savedPath === startedFromPath) return;
      const savedSession = sessions.sessions.find((session) => session.id === startedFromPath);
      if (!savedSession) return; // its tab was closed while the save was in flight
      if (file && file.path === startedFromPath) {
        const moved: LoadedFile = { ...file, path: savedPath };
        setSessions((current) => renameSession(current, startedFromPath, moved));
        void showSessionFile(moved);
      } else {
        const moved: LoadedFile = { ...savedSession.file, path: savedPath };
        setSessions((current) => renameSession(current, startedFromPath, moved));
      }
    },
    [file, sessions, showSessionFile],
  );

  // SHELL-17 — switching documents. Unsaved changes in
  // the one being left behind go through the same Save/Discard/Cancel prompt
  // as opening or closing a file (only the showing document can be dirty,
  // which is exactly why the prompt happens here).
  const selectSession = useCallback(
    async (id: string): Promise<void> => {
      const target = sessions.sessions.find((session) => session.id === id);
      if (!target || target.id === sessions.activeId) return;
      const canProceed = await confirmDiscardChanges();
      if (!canProceed) return;
      setSessions((current) => activateSession(current, id));
      await showSessionFile(target.file);
    },
    [confirmDiscardChanges, sessions, showSessionFile],
  );

  const closeSessionById = useCallback(
    async (id: string): Promise<void> => {
      if (id === sessions.activeId) {
        const canProceed = await confirmDiscardChanges();
        if (!canProceed) return;
      }
      const next = closeSession(sessions, id);
      setSessions(next);
      const nowActive = activeSession(next);
      if (nowActive) {
        if (nowActive.id !== sessions.activeId) await showSessionFile(nowActive.file);
      } else {
        clear();
      }
    },
    [clear, confirmDiscardChanges, sessions, showSessionFile],
  );

  const closeActiveSession = useCallback(async (): Promise<void> => {
    if (sessions.activeId === null) {
      await closeFile();
      return;
    }
    await closeSessionById(sessions.activeId);
  }, [closeFile, closeSessionById, sessions.activeId]);

  const stepSession = useCallback(
    (delta: 1 | -1): void => {
      const target = neighbourSession(sessions, delta);
      if (target) void selectSession(target.id);
    },
    [selectSession, sessions],
  );

  const reopenClosedSession = useCallback((): void => {
    const taken = takeLastClosedPath(sessions);
    if (!taken) return;
    setSessions(taken.state);
    // A real load, not a replay of kept bytes: a closed document keeps no
    // content, and this way a file that has since been moved or deleted says
    // so through the usual banner instead of reopening as an empty document.
    void openFileFromPath(taken.path);
  }, [openFileFromPath, sessions]);

  // SHELL-17 — Ctrl+Tab / Ctrl+Shift+Tab walk the open documents, and
  // Ctrl+Shift+T brings back the last one that was closed.
  useShellShortcut(
    useCallback(
      (event: KeyboardEvent): boolean => {
        if (!(event.ctrlKey || event.metaKey)) return false;
        if (event.key === 'Tab') {
          event.preventDefault();
          stepSession(event.shiftKey ? -1 : 1);
          return true;
        }
        if (event.shiftKey && event.key.toLowerCase() === 't') {
          event.preventDefault();
          reopenClosedSession();
          return true;
        }
        return false;
      },
      [reopenClosedSession, stepSession],
    ),
  );

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
    // Dropping several files at once used to silently open only
    // `files[0]` and discard the rest with no feedback at all — found by
    // driving the real app. Atlas already supports many open tabs (SHELL-17),
    // so every dropped file opens, one at a time (each fully awaited, so a
    // dirty-document prompt or an extension-mismatch confirm from one drop
    // resolves before the next file's own open begins).
    const files = Array.from(e.dataTransfer.files);
    void (async () => {
      for (const droppedFile of files) {
        const path = await resolveDroppedFilePath(window.electronAPI, droppedFile);
        if (path) await openFileFromPath(path);
      }
    })();
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
  const [newMenuOpen, setNewMenuOpen] = useState(false);

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
          const target = document.getElementById(item.id);
          if (target) scrollIntoViewRespectingMotionPreference(target, { behavior: 'smooth', block: 'start' });
        },
      })),
    [tocItems]
  );
  // NEW-01 — this used to key off `isMarkdownDocument ? localMarkdown.length
  // > 0 : ...`, so a genuinely loaded but EMPTY markdown file (a 0-byte
  // `.md`, or one just emptied by the user) fell through to the Welcome
  // screen instead of showing an (empty) editor bound to that path — the
  // opposite of every other format, where merely having a `file` loaded is
  // enough. A real `file` always means content, even zero-length content;
  // only the no-file "typed/loaded sample markdown into a blank untitled
  // session" case still depends on the draft's own length.
  const hasContent = useMemo(
    () => (file ? true : localMarkdown.length > 0),
    [file, localMarkdown.length]
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
  // UX-12 — csv/tsv files always carry real delimited-text content directly;
  // X1 adds a real per-sheet CSV export for xlsx/ods (parsed directly from
  // the file's own bytes — see `exportSpreadsheetCsvPerSheet` — no viewer
  // involvement needed); any other format only gets one once its viewer
  // registers it via getExportableContent (a Phase-3 placeholder today).
  const canExportCsv = useMemo(
    () =>
      currentFormat === 'csv' ||
      currentFormat === 'tsv' ||
      currentFormat === 'xlsx' ||
      currentFormat === 'ods' ||
      exportableContentFormat === 'csv',
    [currentFormat, exportableContentFormat]
  );

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
    // P2.6 — dismiss any pending draft-recovery prompt the same way opening a
    // real file already does. Without this, loading the sample while the
    // banner is still showing would let autosave (now enabled, since
    // localMarkdown is non-empty) start overwriting the crashed session's
    // draft under the same storage key within its 800ms debounce — while the
    // banner still offers a now-stale "Restore" for content already evicted
    // underneath it. The in-memory pendingDraft snapshot still restores
    // correctly if the user clicks it in that narrow window, but a second
    // crash before they do would then lose the original draft for good.
    setPendingDraft(null);
  }, [setMarkdown]); // setPendingDraft is a stable state setter, omitted per convention elsewhere in this file

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

  // NEW-01 — main shows a native Save dialog and writes a blank template
  // atomically (see electron/main.cjs's `document:new`); the renderer never
  // picks the destination path itself. `openFileFromPath` then opens the
  // freshly created file through the exact same guarded path Open/Recent/
  // drag-drop already use — its own `confirmDiscardChanges` check (inside
  // `useFileHandler`) still runs before any current unsaved work is
  // replaced, exactly as if the user had picked this brand-new file from
  // the Open dialog.
  const handleNewDocument = useCallback(
    async (format: NewDocumentFormat): Promise<void> => {
      if (!window.electronAPI?.newDocument) {
        // Mirrors useFileHandler's own BROWSER_MODE_ERROR messaging — "New"
        // needs the desktop app's native Save dialog + main-owned write
        // path, same as Open/Save already do.
        showToast(t('errors.browserModeUnavailable'), 'error');
        return;
      }
      const result = await window.electronAPI.newDocument(format);
      if (result.created) {
        await openFileFromPath(result.path);
      } else if (result.error) {
        const message = translateWriteError(t, result);
        if (message) showToast(message, 'error');
      }
    },
    [openFileFromPath, showToast, t]
  );

  const handleNonMarkdownExport = useCallback(
    async (format: ExportFormat, baseName: string): Promise<void> => {
      if (!file) return;

      const exportUtils = await loadExportUtils();

      if (format === 'copy') {
        // X1/UX-11 — a "Save a copy" passthrough of the file's own original
        // bytes, always through the native save dialog (never a raw browser
        // download — see `exportPdfCopy`/`exportWorkbookCopy`).
        if (file.kind === 'binary' && file.format === 'pdf') {
          await exportUtils.exportPdfCopy(file.content, baseName);
        } else if (file.kind === 'binary' && (file.format === 'xlsx' || file.format === 'ods')) {
          await exportUtils.exportWorkbookCopy(file.content, baseName, file.format);
        }
        return;
      }

      if (format === 'csv') {
        // UX-12 — prefer a viewer-registered real export payload
        // (getExportableContent is a Phase-3 placeholder today — see
        // viewerContextValue.ts).
        const exportable = exportContentRef.current();
        if (exportable && exportable.format === 'csv' && typeof exportable.data === 'string') {
          await exportUtils.exportCsv(exportable.data, exportable.suggestedName || baseName, 'csv');
          return;
        }
        // csv/tsv files already carry their own real delimited-text content
        // directly, with no viewer needed.
        if (file.kind === 'text' && (file.format === 'csv' || file.format === 'tsv')) {
          await exportUtils.exportCsv(file.content, baseName, file.format);
          return;
        }
        // X1 — xlsx/ods are parsed directly from their own bytes (the exact
        // same shared parser the viewer uses), one CSV per visible sheet.
        if (file.kind === 'binary' && (file.format === 'xlsx' || file.format === 'ods')) {
          await exportUtils.exportSpreadsheetCsvPerSheet(file.content, baseName);
        }
        return;
      }

      if (format === 'html') {
        // X1 — text/code only; every other non-markdown format's ExportMenu
        // never offers 'html'.
        if (file.kind === 'text' && (file.format === 'text' || file.format === 'code')) {
          await exportUtils.exportTextHtml(file.content, baseName);
        }
        return;
      }

      // format === 'pdf' — X1's real per-format export. Each branch reuses
      // whatever representation of the document already has its FULL
      // content (the live DOM for formats that render everything up front,
      // the file's own parsed bytes for virtualized viewers) — see each
      // module's header comment in `src/utils/export/` for why.
      switch (file.format) {
        case 'docx':
          await exportUtils.exportDocxPdf('viewer-content', baseName);
          break;
        case 'rtf':
          await exportUtils.exportRtfPdf('viewer-content', baseName);
          break;
        case 'odt':
          await exportUtils.exportOdtPdf('viewer-content', baseName);
          break;
        case 'pptx':
          if (file.kind === 'binary') await exportUtils.exportSlidesPdf(file.content, baseName, 'pptx');
          break;
        case 'odp':
          if (file.kind === 'binary') await exportUtils.exportSlidesPdf(file.content, baseName, 'odp');
          break;
        case 'xlsx':
        case 'ods':
          if (file.kind === 'binary') await exportUtils.exportSpreadsheetPdf(file.content, baseName);
          break;
        case 'csv':
        case 'tsv':
          if (file.kind === 'text') await exportUtils.exportDelimitedTablePdf(file.content, baseName, file.format);
          break;
        case 'text':
        case 'code':
          if (file.kind === 'text') await exportUtils.exportTextPdf(file.content, baseName);
          break;
        case 'pdf':
          if (file.kind === 'binary') await exportUtils.exportPdfCopy(file.content, baseName);
          break;
        case 'doc':
        case 'ppt':
          // wave-4 legacy-office — read-only, text-only viewers with no
          // render surface these export helpers know how to print; export
          // is explicitly out of scope for these two formats (see
          // `legacy/doc`/`legacy/ppt` module headers).
          throw new Error('PDF export failed: this file type is not supported.');
        case 'unknown':
          throw new Error('PDF export failed: this file type is not supported.');
        case 'markdown':
          // Unreachable — the caller only reaches `handleNonMarkdownExport`
          // when `!isMarkdownDocument`. Guarded explicitly (review fix) so
          // the `default` below stays a true `assertNever` exhaustiveness
          // check: adding a new FormatId without a case here now fails to
          // compile instead of silently exporting nothing with no error.
          throw new Error('PDF export failed: this file type is not supported.');
        default:
          assertNever(file.format);
      }
    },
    [exportContentRef, file],
  );

  const handleExport = useCallback(
    async (format: ExportFormat) => {
      const baseName = fileName ?? 'document.md';
      try {
        if (!isMarkdownDocument) {
          await handleNonMarkdownExport(format, baseName);
          return;
        }

        const exportUtils = await loadExportUtils();

        switch (format) {
          case 'md':
            await exportUtils.exportMarkdown(localMarkdown, baseName);
            break;
          case 'html':
            // X2 — serializes the live-rendered #markdown-content DOM
            // (KaTeX/Mermaid/highlighted code) instead of re-parsing the raw
            // markdown source.
            await exportUtils.exportHtml('markdown-content', baseName, theme);
            break;
          case 'pdf':
            // X1 — vector printToPDF of the live DOM (selectable text)
            // instead of the old html2canvas-pro raster screenshot.
            await exportUtils.exportMarkdownPdf('markdown-content', baseName, theme);
            break;
          case 'docx':
            await exportUtils.exportDocx(localMarkdown, baseName);
            break;
          case 'csv':
          case 'copy':
            // Never offered for markdown documents — ExportMenu gates these
            // items on the active (non-markdown) format.
            break;
        }
      } catch (err) {
        // UX-18 — a themed, non-blocking toast instead of a blocking,
        // theme-ignoring native alert(). Each exportX() already wraps its own
        // thrown errors in a friendly, format-specific message (RUN-14), so
        // this only needs to surface it, not wrap it again.
        console.error('Export failed:', err);
        showToast(err instanceof Error ? err.message : String(err), 'error');
      }
    },
    [fileName, handleNonMarkdownExport, isMarkdownDocument, localMarkdown, showToast, theme]
  );

  useUniversalShortcuts({
    openFile,
    saveFile,
    saveFileAs,
    exportPrimary: () => handleExport('pdf'),
    openExportMenu: () => setExportMenuOpen(true),
    openNewMenu: () => setNewMenuOpen(true),
    cycleTheme,
    toggleSidebar: () => setSidebarOpen(prev => !prev),
    setViewMode,
    toggleShortcuts: () => setShortcutsOpen(prev => !prev),
    closeFile: () => void closeActiveSession(),
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
        canExportCsv={canExportCsv}
        newMenuOpen={newMenuOpen}
        onSelectTheme={setTheme}
        onViewModeChange={setViewMode}
        onToggleSidebar={() => setSidebarOpen(prev => !prev)}
        onOpenFile={openFile}
        onNewDocument={(format) => void handleNewDocument(format)}
        onNewMenuOpenChange={setNewMenuOpen}
        onSave={() => void saveFile()}
        onCloseFile={() => void closeActiveSession()}
        onExport={(fmt) => void handleExport(fmt)}
        onOpenSearch={openSearch}
        onShowShortcuts={() => setShortcutsOpen(true)}
        onIncreaseFont={increase}
        onDecreaseFont={decrease}
        onResetFont={reset}
        onExportMenuOpenChange={setExportMenuOpen}
      />

      <TabBar
        sessions={sessions.sessions}
        activeId={sessions.activeId}
        isActiveDirty={combinedDirty}
        onSelect={(id) => void selectSession(id)}
        onClose={(id) => void closeSessionById(id)}
        onReorder={(fromIndex, toIndex) => setSessions((current) => moveSession(current, fromIndex, toIndex))}
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

      <ViewerProvider filePath={filePath || null} onSavedPath={handleViewerSavedPath}>
        <ViewerSessionBridge
          onDirtyChange={setViewerDirty}
          saveRef={viewerSaveRef}
          saveAsRef={viewerSaveAsRef}
          exportContentRef={exportContentRef}
          filePath={filePath || null}
          onExportableFormatChange={setExportableContentFormat}
        />
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
                    {previewTooLarge ? (
                      <LargeMarkdownNotice
                        characters={localMarkdown.length}
                        onRenderAnyway={() => setRenderLargePreview(true)}
                      />
                    ) : (
                      <Suspense fallback={<ViewerLoading format="markdown" />}>
                        <MarkdownRenderer
                          ref={contentRef}
                          markdown={localMarkdown}
                          searchQuery={searchOpen ? searchQuery : undefined}
                        />
                      </Suspense>
                    )}
                  </div>
                )}
              </main>
            </>
          ) : viewerFile ? (
            <main className="content content--viewer">
              {/* P2.1/X4 — the same `contentRef` the markdown preview uses for
                  its TreeWalker search is attached here too (only the branch
                  that's actually mounted ever owns it) so Text/Code/RTF/ODT
                  get the identical search behavior markdown already has. */}
              <div className="preview-panel" id="viewer-content" ref={canSearchFormat ? contentRef : undefined}>
                {/* NEW-01 — a genuinely 0-byte file of a format Atlas has no
                    blank template for (main.cjs only substitutes one for
                    docx/xlsx/ods/pptx/odp — see `substituteBlankTemplateIfEmpty`)
                    gets a friendly notice instead of reaching that format's
                    parser as an empty buffer and throwing a low-level error. */}
                {viewerFile.kind === 'binary' && viewerFile.content.byteLength === 0 ? (
                  <EmptyFileNotice fileName={fileName ?? viewerFile.path} />
                ) : (
                  <ViewerRouter file={viewerFile} />
                )}
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
    </div>
  );
}

/**
 * UX-18 — wraps the shell in its own <ToastProvider> so `<App />` is a
 * complete, self-contained tree (usable as-is from main.tsx or a test's
 * `render(<App />)`) without every caller needing to remember to supply one.
 *
 * P2.1 — also wraps `AppShell` in the centralized shortcut dispatcher's
 * provider. See `AppShell`'s own doc comment for why this can't just be one
 * component.
 */
function App() {
  return (
    <LocaleProvider>
      <ToastProvider>
        <ShortcutManagerProvider>
          <AppShell />
        </ShortcutManagerProvider>
      </ToastProvider>
    </LocaleProvider>
  );
}

export default App;

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
import { WelcomeScreen } from './components/WelcomeScreen';
import { StatusBar } from './components/StatusBar';
import { ShortcutsModal } from './components/ShortcutsModal';
import { SAMPLE_MARKDOWN } from './constants';
import { THEMES, type ViewMode, type ExportFormat, type RecentFile } from './types';
import { exportMarkdown, exportHtml, exportPdf, exportDocx } from './utils/export';
import type { LoadedFile, NavItem } from './formats/types';
import { ViewerProvider } from './viewers/shared/ViewerContext';
import { useSetNavItems, useSetViewerStats } from './viewers/shared/useViewerContext';

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

  const {
    file,
    markdown,
    fileName,
    filePath,
    openDialog: openFile,
    loadFromPath: openFileFromPath,
  } = useFileHandler();

  // Derive isDirty / setMarkdown / saveFile / saveFileAs / drag handlers from
  // legacy autosave hook until W3.5 migrates App.tsx fully to LoadedFile.
  // For now markdown content comes from the shim (file.kind==='text' && format==='markdown').
  const [localMarkdown, setLocalMarkdown] = useState<string>(() => markdown ?? '');
  const [isDirty, setIsDirty] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [lastSyncedMarkdown, setLastSyncedMarkdown] = useState<string>(() => markdown ?? '');

  if (markdown !== undefined && markdown !== null && markdown !== lastSyncedMarkdown) {
    setLastSyncedMarkdown(markdown);
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

  const saveFile = useCallback(async (): Promise<boolean> => {
    if (!isMarkdownDocument) return false;
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

  const handleDragEnter = useCallback((e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation(); setIsDragging(true);
  }, []);
  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
    if (e.currentTarget === e.target) setIsDragging(false);
  }, []);
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
  }, []);
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); e.stopPropagation();
    setIsDragging(false);
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile && 'path' in droppedFile) {
      void openFileFromPath((droppedFile as File & { path: string }).path);
    }
  }, [openFileFromPath]);

  // Track recents when a file is loaded
  useEffect(() => {
    if (file) {
      const name = file.path.split(/[\\/]/).pop() ?? file.path;
      addRecent({ name, path: file.path });
    }
  }, [file, addRecent]);

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
        void openFileFromPath(file.path);
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
    if (!isDirty) return;
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [isDirty]);

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
        isDirty={isDirty}
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
              <div className="preview-panel" id="viewer-content">
                <ViewerRouter file={viewerFile} />
              </div>
            </main>
          ) : null}
        </div>

        {hasContent && <StatusBar fileName={fileName} isDirty={isDirty} />}
      </ViewerProvider>

      <ShortcutsModal isOpen={shortcutsOpen} onClose={() => setShortcutsOpen(false)} />

      <DropZone isVisible={isDragging} />

      {/* Suppress unused-var warnings for filePath while keeping it part of the contract */}
      {filePath && <span style={{ display: 'none' }} aria-hidden="true">{filePath}</span>}
    </div>
  );
}

export default App;

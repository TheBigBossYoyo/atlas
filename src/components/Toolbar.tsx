import {
  FileText,
  Save,
  X,
  Columns2,
  Eye,
  Code2,
  Search,
  PanelLeftClose,
  PanelLeft,
  Keyboard,
  Plus,
  Minus,
  RotateCcw,
} from 'lucide-react';
import type { FormatId } from '../formats/types';
import type { Theme, ThemeMeta, ViewMode, ExportFormat } from '../types';
import type { NewDocumentFormat } from '../electron';
import { ExportMenu } from './ExportMenu';
import { NewDocumentMenu } from './NewDocumentMenu';
import { ThemeMenu } from './ThemeMenu';

interface ToolbarProps {
  theme: Theme;
  themes: readonly ThemeMeta[];
  viewMode: ViewMode;
  sidebarOpen: boolean;
  fileName: string | null;
  isDirty: boolean;
  hasContent: boolean;
  canSave: boolean;
  canSearch: boolean;
  canChangeFontSize: boolean;
  /** P2.8/SHELL-16 — whether a document is open to close (Ctrl+W is enabled independently of this button's own visibility). */
  canClose: boolean;
  isMarkdown: boolean;
  isElectron: boolean;
  exportFormat: FormatId;
  exportMenuOpen: boolean;
  /** UX-12 — whether a real CSV export is available for the current document. */
  canExportCsv: boolean;
  /** NEW-01 — whether Atlas is running inside Electron; "New" needs `window.electronAPI.newDocument`, unavailable in a plain browser tab. */
  newMenuOpen: boolean;
  onSelectTheme: (t: Theme) => void;
  onViewModeChange: (mode: ViewMode) => void;
  onToggleSidebar: () => void;
  onOpenFile: () => void;
  onNewDocument: (format: NewDocumentFormat) => void;
  onNewMenuOpenChange: (open: boolean) => void;
  onSave: () => void;
  onCloseFile: () => void;
  onExport: (format: ExportFormat) => void;
  onOpenSearch: () => void;
  onShowShortcuts: () => void;
  onIncreaseFont: () => void;
  onDecreaseFont: () => void;
  onResetFont: () => void;
  onExportMenuOpenChange: (open: boolean) => void;
}

const VIEW_MODES: { mode: ViewMode; icon: typeof Eye; label: string }[] = [
  { mode: 'preview', icon: Eye, label: 'Preview' },
  { mode: 'split', icon: Columns2, label: 'Split' },
  { mode: 'editor', icon: Code2, label: 'Editor' },
];

export function Toolbar({
  theme,
  themes,
  viewMode,
  sidebarOpen,
  fileName,
  isDirty,
  hasContent,
  canSave,
  canSearch,
  canChangeFontSize,
  canClose,
  isMarkdown,
  isElectron,
  exportFormat,
  exportMenuOpen,
  canExportCsv,
  newMenuOpen,
  onSelectTheme,
  onViewModeChange,
  onToggleSidebar,
  onOpenFile,
  onNewDocument,
  onNewMenuOpenChange,
  onSave,
  onCloseFile,
  onExport,
  onOpenSearch,
  onShowShortcuts,
  onIncreaseFont,
  onDecreaseFont,
  onResetFont,
  onExportMenuOpenChange,
}: ToolbarProps) {
  return (
    <header className={`toolbar ${isElectron ? 'toolbar--electron' : ''}`}>
      <div className="toolbar__left">
        <button
          className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
          onClick={onToggleSidebar}
          title={sidebarOpen ? 'Hide sidebar' : 'Show sidebar'}
          aria-label="Toggle sidebar"
        >
          {sidebarOpen ? <PanelLeftClose size={18} /> : <PanelLeft size={18} />}
        </button>

        {/* Rendered unconditionally, like Open — `onNewDocument` itself
            surfaces a friendly error in plain-browser-tab mode (no
            `window.electronAPI`), the same way `openFile` already does. */}
        <NewDocumentMenu onCreate={onNewDocument} open={newMenuOpen} onOpenChange={onNewMenuOpenChange} />

        <button className="toolbar__btn toolbar__nodrag" onClick={onOpenFile} title="Open file (Ctrl+O)">
          <FileText size={16} />
          <span>Open</span>
        </button>

        {canSave ? (
          <button
            className="toolbar__btn toolbar__nodrag"
            onClick={onSave}
            title="Save (Ctrl+S)"
          >
            <Save size={16} />
            <span>Save</span>
          </button>
        ) : null}

        {fileName && (
          <span className="toolbar__filename">
            {fileName}
            {isDirty && <span className="toolbar__dirty" aria-label="Unsaved changes"> ●</span>}
          </span>
        )}

        {canClose ? (
          <button
            className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
            onClick={onCloseFile}
            title="Close file (Ctrl+W)"
            aria-label="Close file"
          >
            <X size={16} />
          </button>
        ) : null}
      </div>

      <div className="toolbar__center">
        {isMarkdown ? (
          <div className="toolbar__view-toggle toolbar__nodrag">
            {VIEW_MODES.map(({ mode, icon: Icon, label }) => (
              <button
                key={mode}
                className={`toolbar__view-btn ${viewMode === mode ? 'toolbar__view-btn--active' : ''}`}
                onClick={() => onViewModeChange(mode)}
                title={label}
                aria-label={label}
              >
                <Icon size={15} />
                <span>{label}</span>
              </button>
            ))}
          </div>
        ) : null}
      </div>

      <div className="toolbar__right">
        {canChangeFontSize ? (
          <div className="toolbar__font-group toolbar__nodrag" role="group" aria-label="Font size">
            <button
              className="toolbar__btn toolbar__btn--icon"
              onClick={onDecreaseFont}
              title="Decrease font size (Ctrl+-)"
              aria-label="Decrease font size"
            >
              <Minus size={16} />
            </button>
            <button
              className="toolbar__btn toolbar__btn--icon"
              onClick={onResetFont}
              title="Reset font size (Ctrl+0)"
              aria-label="Reset font size"
            >
              <RotateCcw size={14} />
            </button>
            <button
              className="toolbar__btn toolbar__btn--icon"
              onClick={onIncreaseFont}
              title="Increase font size (Ctrl+=)"
              aria-label="Increase font size"
            >
              <Plus size={16} />
            </button>
          </div>
        ) : null}

        {canSearch ? (
          <button
            className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
            onClick={onOpenSearch}
            title="Search (Ctrl+F)"
            aria-label="Search"
          >
            <Search size={18} />
          </button>
        ) : null}

        <ExportMenu
          onExport={onExport}
          format={exportFormat}
          open={exportMenuOpen}
          onOpenChange={onExportMenuOpenChange}
          disabled={!hasContent}
          canExportCsv={canExportCsv}
        />

        <ThemeMenu current={theme} themes={themes} onSelect={onSelectTheme} />

        <button
          className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
          onClick={onShowShortcuts}
          title="Keyboard shortcuts (Ctrl+/)"
          aria-label="Keyboard shortcuts"
        >
          <Keyboard size={18} />
        </button>
      </div>
    </header>
  );
}

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
import { LanguageMenu } from './LanguageMenu';
import { useTranslate } from '../i18n';

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

const VIEW_MODES: { mode: ViewMode; icon: typeof Eye; labelKey: string }[] = [
  { mode: 'preview', icon: Eye, labelKey: 'toolbar.viewMode.preview' },
  { mode: 'split', icon: Columns2, labelKey: 'toolbar.viewMode.split' },
  { mode: 'editor', icon: Code2, labelKey: 'toolbar.viewMode.editor' },
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
  const t = useTranslate();
  return (
    <header className={`toolbar ${isElectron ? 'toolbar--electron' : ''}`}>
      <div className="toolbar__left">
        <button
          className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
          onClick={onToggleSidebar}
          title={sidebarOpen ? t('toolbar.hideSidebar') : t('toolbar.showSidebar')}
          aria-label={t('toolbar.toggleSidebarAria')}
        >
          {sidebarOpen ? <PanelLeftClose size={18} /> : <PanelLeft size={18} />}
        </button>

        {/* Rendered unconditionally, like Open — `onNewDocument` itself
            surfaces a friendly error in plain-browser-tab mode (no
            `window.electronAPI`), the same way `openFile` already does. */}
        <NewDocumentMenu onCreate={onNewDocument} open={newMenuOpen} onOpenChange={onNewMenuOpenChange} />

        <button className="toolbar__btn toolbar__nodrag" onClick={onOpenFile} title={t('toolbar.openTitle')}>
          <FileText size={16} />
          <span>{t('toolbar.open')}</span>
        </button>

        {canSave ? (
          <button
            className="toolbar__btn toolbar__nodrag"
            onClick={onSave}
            title={t('toolbar.saveTitle')}
          >
            <Save size={16} />
            <span>{t('toolbar.save')}</span>
          </button>
        ) : null}

        {fileName && (
          <span className="toolbar__filename">
            {fileName}
            {isDirty && <span className="toolbar__dirty" aria-label={t('common.unsavedChanges')}> ●</span>}
          </span>
        )}

        {canClose ? (
          <button
            className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
            onClick={onCloseFile}
            title={t('toolbar.closeFileTitle')}
            aria-label={t('toolbar.closeFileAria')}
          >
            <X size={16} />
          </button>
        ) : null}
      </div>

      <div className="toolbar__center">
        {isMarkdown ? (
          <div className="toolbar__view-toggle toolbar__nodrag">
            {VIEW_MODES.map(({ mode, icon: Icon, labelKey }) => {
              const label = t(labelKey);
              return (
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
              );
            })}
          </div>
        ) : null}
      </div>

      <div className="toolbar__right">
        {canChangeFontSize ? (
          <div className="toolbar__font-group toolbar__nodrag" role="group" aria-label={t('toolbar.fontSizeGroupAria')}>
            <button
              className="toolbar__btn toolbar__btn--icon"
              onClick={onDecreaseFont}
              title={t('toolbar.decreaseFontTitle')}
              aria-label={t('toolbar.decreaseFontAria')}
            >
              <Minus size={16} />
            </button>
            <button
              className="toolbar__btn toolbar__btn--icon"
              onClick={onResetFont}
              title={t('toolbar.resetFontTitle')}
              aria-label={t('toolbar.resetFontTitle')}
            >
              <RotateCcw size={14} />
            </button>
            <button
              className="toolbar__btn toolbar__btn--icon"
              onClick={onIncreaseFont}
              title={t('toolbar.increaseFontTitle')}
              aria-label={t('toolbar.increaseFontAria')}
            >
              <Plus size={16} />
            </button>
          </div>
        ) : null}

        {canSearch ? (
          <button
            className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
            onClick={onOpenSearch}
            title={t('toolbar.searchTitle')}
            aria-label={t('toolbar.searchAria')}
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

        <LanguageMenu />

        <button
          className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
          onClick={onShowShortcuts}
          title={t('toolbar.shortcutsTitle')}
          aria-label={t('toolbar.shortcutsAria')}
        >
          <Keyboard size={18} />
        </button>
      </div>
    </header>
  );
}

import React, { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import {
  Bold,
  Italic,
  Underline,
  Strikethrough,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  List,
  ListOrdered,
  IndentIncrease,
  IndentDecrease,
  Table as TableIcon,
  Image as ImageIcon,
  Link as LinkIcon,
  Heading,
  FileText,
  Columns,
  Undo2,
  Redo2,
  Search,
  MessageSquare,
  Eye,
  Check,
  X,
  Subscript,
  Superscript
} from 'lucide-react';
import type { ToolbarCommand, ToolbarState } from './toolbarTypes';
import { FontPicker } from './FontPicker';
import { TableEditMenuItems } from './TableEditMenuItems';
import { TablePropertiesDialog } from './TablePropertiesDialog';
import { useTranslate } from '../../../i18n';
import './__styles__/toolbar.css';

function useClickOutside(ref: React.RefObject<HTMLElement | null>, handler: () => void) {
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (ref.current && !ref.current.contains(event.target as Node)) {
        handler();
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [ref, handler]);
}

const DEFAULT_FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Georgia', 'Tahoma', 'Cambria'];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
const DEFAULT_COLORS = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff',
  '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff'
];

/** D18/DXE-06/DXE-12 — controls with no real implementation behind them yet
 * (unlike lists/indent/hyperlink/page-break/table, which this wave
 * implemented) get a disabled state and an explanatory tooltip instead of
 * silently no-opping when clicked. Translated with the free `t()` (not a
 * hook) since it's read at module scope by components below that also
 * call `useTranslate()` themselves — this constant only needs the string
 * once per render pass, not a live subscription. */

interface ToolbarProps {
  state: ToolbarState;
  onCommand: (cmd: ToolbarCommand) => void;
  activeTab?: 'home' | 'insert' | 'layout' | 'review';
  onTabChange?: (tab: 'home' | 'insert' | 'layout' | 'review') => void;
  availableStyles?: ReadonlyArray<{ id: string; name: string }>;
  availableFonts?: ReadonlyArray<string>;
  /** Fonts the open document already uses — listed first in the font picker. */
  documentFonts?: ReadonlyArray<string>;
  /** USR-12 — document-level actions (save, print, zoom…) shown compactly at the right of the tab row. */
  trailing?: React.ReactNode;
}

export const Toolbar: React.FC<ToolbarProps> = ({
  state,
  onCommand,
  activeTab: controlledTab,
  onTabChange,
  availableFonts = DEFAULT_FONTS,
  documentFonts = [],
  trailing,
}) => {
  const t = useTranslate();
  const [internalTab, setInternalTab] = useState<'home' | 'insert' | 'layout' | 'review'>('home');
  const currentTab = controlledTab !== undefined ? controlledTab : internalTab;

  const handleTabClick = useCallback((tab: 'home' | 'insert' | 'layout' | 'review') => {
    if (onTabChange) {
      onTabChange(tab);
    } else {
      setInternalTab(tab);
    }
  }, [onTabChange]);

  const activeFormats = useMemo(() => state.activeFormats, [state.activeFormats]);

  const handleCommand = useCallback((cmd: ToolbarCommand) => {
    onCommand(cmd);
  }, [onCommand]);

  const TAB_LABELS = {
    home: t('docx.toolbar.tab.home'),
    insert: t('docx.toolbar.tab.insert'),
    layout: t('docx.toolbar.tab.layout'),
    review: t('docx.toolbar.tab.review'),
  } as const;

  return (
    <div className="docx-toolbar" role="toolbar" aria-label={t('docx.toolbar.rootAria')} aria-orientation="horizontal">
      <div className="docx-toolbar__header">
      <div className="docx-toolbar__tabs" role="tablist" aria-label={t('docx.toolbar.tabsAria')}>
        {(['home', 'insert', 'layout', 'review'] as const).map(tab => (
          <button
            key={tab}
            type="button"
            role="tab"
            aria-selected={currentTab === tab}
            aria-controls={`docx-toolbar-panel-${tab}`}
            id={`docx-toolbar-tab-${tab}`}
            className={`docx-toolbar__tab ${currentTab === tab ? 'docx-toolbar__tab--active' : ''}`}
            onClick={() => handleTabClick(tab)}
          >
            {TAB_LABELS[tab]}
          </button>
        ))}
      </div>
      {trailing !== undefined && <div className="docx-toolbar__trailing">{trailing}</div>}
      </div>
      <div
        className="docx-toolbar__panel"
        role="tabpanel"
        id={`docx-toolbar-panel-${currentTab}`}
        aria-labelledby={`docx-toolbar-tab-${currentTab}`}
      >
        {currentTab === 'home' && (
          <HomeTab
            state={state}
            activeFormats={activeFormats}
            availableFonts={availableFonts}
            documentFonts={documentFonts}
            onCommand={handleCommand}
          />
        )}
        {currentTab === 'insert' && (
          <InsertTab onCommand={handleCommand} state={state} />
        )}
        {currentTab === 'layout' && (
          <LayoutTab onCommand={handleCommand} />
        )}
        {currentTab === 'review' && (
          <ReviewTab state={state} onCommand={handleCommand} />
        )}
      </div>
    </div>
  );
};

// --- TABS ---

interface TabProps {
  onCommand: (cmd: ToolbarCommand) => void;
  state?: ToolbarState;
  activeFormats?: ReadonlySet<string>;
  availableFonts?: ReadonlyArray<string>;
  documentFonts?: ReadonlyArray<string>;
}

const HomeTab: React.FC<TabProps> = ({ onCommand, state, activeFormats, availableFonts, documentFonts }) => {
  const t = useTranslate();
  return (
    <>
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.undo')} onClick={() => onCommand({ kind: 'undo' })}><Undo2 /></IconButton>
        <IconButton label={t('docx.toolbar.redo')} onClick={() => onCommand({ kind: 'redo' })}><Redo2 /></IconButton>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <FontPicker
          value={state?.fontFamily ?? null}
          fonts={availableFonts ?? DEFAULT_FONTS}
          documentFonts={documentFonts}
          onSelect={(family) => onCommand({ kind: 'set-font-family', family })}
        />
        <select
          className="docx-toolbar__size-select"
          aria-label={t('docx.toolbar.fontSizeAria')}
          value={state?.fontSizePt ?? ''}
          onChange={(e) => onCommand({ kind: 'set-font-size', sizePt: Number(e.target.value) })}
        >
          <option value="" disabled>{t('docx.toolbar.sizePlaceholder')}</option>
          {(state?.fontSizePt != null && !FONT_SIZES.includes(state.fontSizePt) ? [...FONT_SIZES, state.fontSizePt].sort((a, b) => a - b) : FONT_SIZES).map(s => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <FormatButton kind="toggle-bold" label={t('docx.toolbar.bold')} icon={<Bold />} active={activeFormats?.has('bold')} onCommand={onCommand} />
        <FormatButton kind="toggle-italic" label={t('docx.toolbar.italic')} icon={<Italic />} active={activeFormats?.has('italic')} onCommand={onCommand} />
        <FormatButton kind="toggle-underline" label={t('docx.toolbar.underline')} icon={<Underline />} active={activeFormats?.has('underline')} onCommand={onCommand} />
        <FormatButton kind="toggle-strike" label={t('docx.toolbar.strikethrough')} icon={<Strikethrough />} active={activeFormats?.has('strike')} onCommand={onCommand} />
        <FormatButton kind="toggle-subscript" label={t('docx.toolbar.subscript')} icon={<Subscript />} active={activeFormats?.has('subscript')} onCommand={onCommand} />
        <FormatButton kind="toggle-superscript" label={t('docx.toolbar.superscript')} icon={<Superscript />} active={activeFormats?.has('superscript')} onCommand={onCommand} />
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <ColorPickerPopover icon={<span style={{ color: 'var(--toolbar-text)', fontWeight: 'bold' }}>A</span>} onSelect={(hex) => onCommand({ kind: 'set-font-color', colorHex: hex })} title={t('docx.toolbar.fontColor')} />
        <ColorPickerPopover icon={<span style={{ background: 'yellow', display: 'inline-block', width: '12px', height: '12px' }} />} onSelect={(hex) => onCommand({ kind: 'set-highlight-color', colorHex: hex })} title={t('docx.toolbar.highlightColor')} />
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <FormatButton kind="set-alignment" label={t('docx.toolbar.alignLeft')} payload={{ kind: 'set-alignment', align: 'left' }} icon={<AlignLeft />} active={state?.alignment === 'left'} onCommand={onCommand} />
        <FormatButton kind="set-alignment" label={t('docx.toolbar.alignCenter')} payload={{ kind: 'set-alignment', align: 'center' }} icon={<AlignCenter />} active={state?.alignment === 'center'} onCommand={onCommand} />
        <FormatButton kind="set-alignment" label={t('docx.toolbar.alignRight')} payload={{ kind: 'set-alignment', align: 'right' }} icon={<AlignRight />} active={state?.alignment === 'right'} onCommand={onCommand} />
        <FormatButton kind="set-alignment" label={t('docx.toolbar.justify')} payload={{ kind: 'set-alignment', align: 'justify' }} icon={<AlignJustify />} active={state?.alignment === 'justify'} onCommand={onCommand} />
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.bulletList')} onClick={() => onCommand({ kind: 'toggle-bullet-list' })}><List /></IconButton>
        <IconButton label={t('docx.toolbar.numberedList')} onClick={() => onCommand({ kind: 'toggle-numbered-list' })}><ListOrdered /></IconButton>
        <IconButton label={t('docx.toolbar.decreaseIndent')} onClick={() => onCommand({ kind: 'change-indent', delta: -1 })}><IndentDecrease /></IconButton>
        <IconButton label={t('docx.toolbar.increaseIndent')} onClick={() => onCommand({ kind: 'change-indent', delta: 1 })}><IndentIncrease /></IconButton>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.findAndReplace')} onClick={() => onCommand({ kind: 'open-find-replace' })}><Search /></IconButton>
      </div>
    </>
  );
};

const InsertTab: React.FC<TabProps> = ({ onCommand, state }) => {
  const t = useTranslate();
  const notYetSupported = t('docx.toolbar.notYetSupported');
  return (
    <>
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.pageBreak')} onClick={() => onCommand({ kind: 'insert-page-break' })}><FileText /></IconButton>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <TablePickerPopover onCommand={onCommand} />
        <TableEditPopover onCommand={onCommand} state={state} />
        <IconButton label={t('docx.toolbar.image')} onClick={() => onCommand({ kind: 'insert-image' })}><ImageIcon /></IconButton>
        <IconButton label={t('docx.toolbar.hyperlink')} onClick={() => onCommand({ kind: 'insert-hyperlink' })}><LinkIcon /></IconButton>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.header')} disabled disabledReason={notYetSupported} onClick={() => onCommand({ kind: 'insert-header' })}><Heading /></IconButton>
        <IconButton label={t('docx.toolbar.footer')} disabled disabledReason={notYetSupported} onClick={() => onCommand({ kind: 'insert-footer' })}><FileText /></IconButton>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.comment')} onClick={() => onCommand({ kind: 'insert-comment' })}><MessageSquare /></IconButton>
      </div>
    </>
  );
};

const LayoutTab: React.FC<TabProps> = ({ onCommand }) => {
  const t = useTranslate();
  const notYetSupported = t('docx.toolbar.notYetSupported');
  return (
    <>
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.margins')} disabled disabledReason={notYetSupported} onClick={() => onCommand({ kind: 'set-margins', preset: 'normal' })}><FileText /></IconButton>
        <IconButton label={t('docx.toolbar.orientation')} disabled disabledReason={notYetSupported} onClick={() => onCommand({ kind: 'set-orientation', orientation: 'portrait' })}><FileText /></IconButton>
        <IconButton label={t('docx.toolbar.pageSize')} disabled disabledReason={notYetSupported} onClick={() => onCommand({ kind: 'set-page-size', preset: 'a4' })}><FileText /></IconButton>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.columns')} disabled disabledReason={notYetSupported} onClick={() => onCommand({ kind: 'set-columns', count: 2 })}><Columns /></IconButton>
      </div>
    </>
  );
};

const ReviewTab: React.FC<TabProps> = ({ onCommand, state }) => {
  const t = useTranslate();
  return (
    <>
      <div className="docx-toolbar__group">
        <FormatButton kind="toggle-spell-check" label={t('docx.toolbar.toggleSpellCheck')} payload={{ kind: 'toggle-spell-check' }} icon={<Check />} active={state?.spellCheck} onCommand={onCommand} />
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <FormatButton kind="toggle-track-changes" label={t('docx.toolbar.toggleTrackChanges')} payload={{ kind: 'toggle-track-changes' }} icon={<Eye />} active={state?.trackChanges} onCommand={onCommand} />
        <IconButton label={t('docx.toolbar.accept')} onClick={() => onCommand({ kind: 'accept-change' })}><Check /></IconButton>
        <IconButton label={t('docx.toolbar.reject')} onClick={() => onCommand({ kind: 'reject-change' })}><X /></IconButton>
        <IconButton label={t('docx.toolbar.acceptAll')} onClick={() => onCommand({ kind: 'accept-all-changes' })}><Check /></IconButton>
        <IconButton label={t('docx.toolbar.rejectAll')} onClick={() => onCommand({ kind: 'reject-all-changes' })}><X /></IconButton>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <IconButton label={t('docx.toolbar.comments')} onClick={() => onCommand({ kind: 'open-comments-pane' })}><MessageSquare /></IconButton>
      </div>
    </>
  );
};

// --- COMPONENTS ---

/** A plain icon-only toolbar button: `title` (tooltip) and `aria-label`
 * (screen reader) always mirror the same human-readable `label` (UX-09).
 * `disabled`+`disabledReason` renders a real disabled control with an
 * explanatory tooltip instead of a silently no-op click handler (D18/DXE-12). */
const IconButton = ({
  label,
  onClick,
  disabled,
  disabledReason,
  children,
}: {
  label: string;
  onClick: () => void;
  disabled?: boolean;
  disabledReason?: string;
  children: React.ReactNode;
}) => {
  return (
    <button
      type="button"
      className="docx-toolbar__button"
      onClick={onClick}
      disabled={disabled}
      title={disabled ? (disabledReason ?? label) : label}
      aria-label={label}
    >
      {children}
    </button>
  );
};

const FormatButton = ({ kind, label, payload, icon, active, onCommand }: { kind: string; label?: string; payload?: ToolbarCommand; icon: React.ReactNode; active?: boolean; onCommand: (cmd: ToolbarCommand) => void }) => {
  const text = label ?? kind;
  return (
    <button
      type="button"
      className={`docx-toolbar__button ${active ? 'docx-toolbar__button--active' : ''}`}
      onClick={() => onCommand(payload || { kind } as ToolbarCommand)}
      title={text}
      aria-label={text}
      aria-pressed={active ?? false}
    >
      {icon}
    </button>
  );
};

const ColorPickerPopover = ({ icon, onSelect, title }: { icon: React.ReactNode; onSelect: (color: string) => void; title: string }) => {
  const t = useTranslate();
  const [open, setOpen] = useState(false);
  const [hex, setHex] = useState('');
  const ref = useRef<HTMLDivElement>(null);
  useClickOutside(ref, () => setOpen(false));

  return (
    <div className="docx-toolbar__popover-container" ref={ref}>
      <button
        type="button"
        className="docx-toolbar__button"
        onClick={() => setOpen(!open)}
        title={title}
        aria-label={title}
        aria-haspopup="true"
        aria-expanded={open}
      >
        {icon}
      </button>
      {open && (
        <div className="docx-toolbar__popover" role="dialog" aria-label={title}>
          <div className="docx-toolbar__color-grid" role="group" aria-label={t('docx.toolbar.colorGridAria')}>
            {DEFAULT_COLORS.map(c => (
              // DXE-23 — a real, keyboard-focusable/activatable button (not a
              // bare div) so the color grid is operable without a mouse.
              <button
                key={c}
                type="button"
                className="docx-toolbar__color-swatch"
                style={{ backgroundColor: c }}
                aria-label={t('docx.toolbar.colorSwatchAria', { hex: c })}
                onClick={() => { onSelect(c); setOpen(false); }}
              />
            ))}
          </div>
          <div className="docx-toolbar__color-input-row">
            <label className="docx-toolbar__visually-hidden" htmlFor="docx-toolbar-custom-color">
              {t('docx.toolbar.customColorLabel')}
            </label>
            <input
              id="docx-toolbar-custom-color"
              type="text"
              value={hex}
              onChange={e => setHex(e.target.value)}
              // Not a translatable label — an example hex value, same in
              // every locale. `{'...'}` (not a bare string literal) so the
              // `noHardcodedStrings` test's blanket literal-attribute check
              // doesn't mistake it for an untranslated string.
              placeholder={'#000000'}
              aria-label={t('docx.toolbar.customColorLabel')}
            />
            <button type="button" onClick={() => { if (hex) { onSelect(hex); setOpen(false); } }}>{t('docx.toolbar.ok')}</button>
          </div>
        </div>
      )}
    </div>
  );
};

const TABLE_PICKER_ROWS = 8;
const TABLE_PICKER_COLS = 10;

const TablePickerPopover = ({ onCommand }: { onCommand: (cmd: ToolbarCommand) => void }) => {
  const t = useTranslate();
  const [open, setOpen] = useState(false);
  const [hoverRow, setHoverRow] = useState(0);
  const [hoverCol, setHoverCol] = useState(0);
  const ref = useRef<HTMLDivElement>(null);
  useClickOutside(ref, () => setOpen(false));

  return (
    <div className="docx-toolbar__popover-container" ref={ref}>
      <button
        type="button"
        className="docx-toolbar__button"
        onClick={() => setOpen(!open)}
        title={t('docx.toolbar.insertTable')}
        aria-label={t('docx.toolbar.insertTable')}
        aria-haspopup="true"
        aria-expanded={open}
      >
        <TableIcon />
      </button>
      {open && (
        <div className="docx-toolbar__popover" role="dialog" aria-label={t('docx.toolbar.insertTableDialogAria')}>
          <div className="docx-toolbar__table-grid" role="group" aria-label={t('docx.toolbar.tableSizeAria')}>
            {Array.from({ length: TABLE_PICKER_ROWS }).map((_, r) => (
              Array.from({ length: TABLE_PICKER_COLS }).map((_, c) => {
                const isHovered = r <= hoverRow && c <= hoverCol;
                const setHover = () => { setHoverRow(r); setHoverCol(c); };
                return (
                  // DXE-23 — a real button so Tab/Shift+Tab and Enter/Space
                  // work; onFocus mirrors onMouseEnter so keyboard users see
                  // the same live "RxC Table" preview as mouse users.
                  <button
                    key={`${r}-${c}`}
                    type="button"
                    className={`docx-toolbar__table-cell ${isHovered ? 'docx-toolbar__table-cell--hover' : ''}`}
                    aria-label={t('docx.toolbar.tableCellAria', { rows: r + 1, cols: c + 1 })}
                    onMouseEnter={setHover}
                    onFocus={setHover}
                    onClick={() => {
                      onCommand({ kind: 'insert-table', rows: hoverRow + 1, cols: hoverCol + 1 });
                      setOpen(false);
                    }}
                  />
                );
              })
            ))}
          </div>
          <div className="docx-toolbar__table-label" aria-live="polite">
            {t('docx.toolbar.tableSizeLabel', { rows: hoverRow + 1, cols: hoverCol + 1 })}
          </div>
        </div>
      )}
    </div>
  );
};

/**
 * DXE-14 — structural table-editing commands (insert/delete row & column,
 * merge, split) plus the table properties dialog, gated on
 * `ToolbarState.insideTable` the same way the rest of this file gates
 * unimplemented controls (D18/DXE-12): a real disabled button with an
 * explanatory tooltip rather than a button that silently does nothing.
 */
const TableEditPopover = ({ onCommand, state }: { onCommand: (cmd: ToolbarCommand) => void; state?: ToolbarState }) => {
  const t = useTranslate();
  const [open, setOpen] = useState(false);
  const [showProperties, setShowProperties] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useClickOutside(ref, () => { setOpen(false); setShowProperties(false); });

  const insideTable = state?.insideTable ?? false;
  const close = () => { setOpen(false); setShowProperties(false); };

  return (
    <div className="docx-toolbar__popover-container" ref={ref}>
      <button
        type="button"
        className="docx-toolbar__button"
        onClick={() => setOpen(o => !o)}
        disabled={!insideTable}
        title={insideTable ? t('docx.toolbar.editTable') : t('docx.toolbar.editTablePlaceCursor')}
        aria-label={t('docx.toolbar.editTable')}
        aria-haspopup="true"
        aria-expanded={open}
      >
        <TableIcon />
      </button>
      {open && insideTable && (
        <div className="docx-toolbar__popover" role="dialog" aria-label={t('docx.toolbar.editTableDialogAria')}>
          {showProperties ? (
            <TablePropertiesDialog
              seed={{
                widthTwips: state?.tableWidthTwips ?? null,
                alignment: state?.tableAlignment ?? null,
                bordersOn: state?.tableBordersOn ?? true,
              }}
              onApply={cmd => { onCommand(cmd); close(); }}
              onCancel={() => setShowProperties(false)}
            />
          ) : (
            <>
              <TableEditMenuItems onCommand={onCommand} onAfterCommand={close} />
              <div className="docx-toolbar__menu-divider" role="separator" />
              <button type="button" className="docx-toolbar__menu-item" onClick={() => setShowProperties(true)}>
                {t('docx.toolbar.tableProperties')}
              </button>
            </>
          )}
        </div>
      )}
    </div>
  );
};

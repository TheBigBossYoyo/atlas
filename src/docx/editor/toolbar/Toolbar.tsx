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
  X
} from 'lucide-react';
import type { ToolbarCommand, ToolbarState } from './toolbarTypes';
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

interface ToolbarProps {
  state: ToolbarState;
  onCommand: (cmd: ToolbarCommand) => void;
  activeTab?: 'home' | 'insert' | 'layout' | 'review';
  onTabChange?: (tab: 'home' | 'insert' | 'layout' | 'review') => void;
  availableStyles?: ReadonlyArray<{ id: string; name: string }>;
  availableFonts?: ReadonlyArray<string>;
}

export const Toolbar: React.FC<ToolbarProps> = ({
  state,
  onCommand,
  activeTab: controlledTab,
  onTabChange,
  availableFonts = DEFAULT_FONTS
}) => {
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

  return (
    <div className="docx-toolbar" role="toolbar" aria-label="Document formatting toolbar" aria-orientation="horizontal">
      <div className="docx-toolbar__tabs" role="tablist" aria-label="Toolbar tabs">
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
            {tab.charAt(0).toUpperCase() + tab.slice(1)}
          </button>
        ))}
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
            onCommand={handleCommand}
          />
        )}
        {currentTab === 'insert' && (
          <InsertTab onCommand={handleCommand} />
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
}

const HomeTab: React.FC<TabProps> = ({ onCommand, state, activeFormats, availableFonts }) => {
  return (
    <>
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'undo' })} title="Undo"><Undo2 /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'redo' })} title="Redo"><Redo2 /></button>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <select
          className="docx-toolbar__font-select"
          value={state?.fontFamily || ''}
          onChange={(e) => onCommand({ kind: 'set-font-family', family: e.target.value })}
        >
          <option value="" disabled>Font</option>
          {availableFonts?.map(f => <option key={f} value={f}>{f}</option>)}
        </select>
        <select
          className="docx-toolbar__size-select"
          value={state?.fontSizePt || ''}
          onChange={(e) => onCommand({ kind: 'set-font-size', sizePt: Number(e.target.value) })}
        >
          <option value="" disabled>Size</option>
          {FONT_SIZES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <FormatButton kind="toggle-bold" icon={<Bold />} active={activeFormats?.has('bold')} onCommand={onCommand} />
        <FormatButton kind="toggle-italic" icon={<Italic />} active={activeFormats?.has('italic')} onCommand={onCommand} />
        <FormatButton kind="toggle-underline" icon={<Underline />} active={activeFormats?.has('underline')} onCommand={onCommand} />
        <FormatButton kind="toggle-strike" icon={<Strikethrough />} active={activeFormats?.has('strike')} onCommand={onCommand} />
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <ColorPickerPopover icon={<span style={{ color: 'var(--toolbar-text)', fontWeight: 'bold' }}>A</span>} onSelect={(hex) => onCommand({ kind: 'set-font-color', colorHex: hex })} title="Font Color" />
        <ColorPickerPopover icon={<span style={{ background: 'yellow', display: 'inline-block', width: '12px', height: '12px' }} />} onSelect={(hex) => onCommand({ kind: 'set-highlight-color', colorHex: hex })} title="Highlight Color" />
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <FormatButton kind="set-alignment" payload={{ kind: 'set-alignment', align: 'left' }} icon={<AlignLeft />} active={state?.alignment === 'left'} onCommand={onCommand} />
        <FormatButton kind="set-alignment" payload={{ kind: 'set-alignment', align: 'center' }} icon={<AlignCenter />} active={state?.alignment === 'center'} onCommand={onCommand} />
        <FormatButton kind="set-alignment" payload={{ kind: 'set-alignment', align: 'right' }} icon={<AlignRight />} active={state?.alignment === 'right'} onCommand={onCommand} />
        <FormatButton kind="set-alignment" payload={{ kind: 'set-alignment', align: 'justify' }} icon={<AlignJustify />} active={state?.alignment === 'justify'} onCommand={onCommand} />
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'toggle-bullet-list' })} title="Bullet List"><List /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'toggle-numbered-list' })} title="Numbered List"><ListOrdered /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'change-indent', delta: -1 })} title="Decrease Indent"><IndentDecrease /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'change-indent', delta: 1 })} title="Increase Indent"><IndentIncrease /></button>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'open-find-replace' })} title="Find and Replace"><Search /></button>
      </div>
    </>
  );
};

const InsertTab: React.FC<TabProps> = ({ onCommand }) => {
  return (
    <>
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'insert-page-break' })} title="Page Break"><FileText /></button>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <TablePickerPopover onCommand={onCommand} />
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'insert-image' })} title="Image"><ImageIcon /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'insert-hyperlink' })} title="Hyperlink"><LinkIcon /></button>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'insert-header' })} title="Header"><Heading /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'insert-footer' })} title="Footer"><FileText /></button>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'insert-comment' })} title="Comment"><MessageSquare /></button>
      </div>
    </>
  );
};

const LayoutTab: React.FC<TabProps> = ({ onCommand }) => {
  return (
    <>
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'set-margins', preset: 'normal' })} title="Margins"><FileText /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'set-orientation', orientation: 'portrait' })} title="Orientation"><FileText /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'set-page-size', preset: 'a4' })} title="Size"><FileText /></button>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'set-columns', count: 2 })} title="Columns"><Columns /></button>
      </div>
    </>
  );
};

const ReviewTab: React.FC<TabProps> = ({ onCommand, state }) => {
  return (
    <>
      <div className="docx-toolbar__group">
        <FormatButton kind="toggle-spell-check" payload={{ kind: 'toggle-spell-check' }} icon={<Check />} active={state?.spellCheck} onCommand={onCommand} />
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <FormatButton kind="toggle-track-changes" payload={{ kind: 'toggle-track-changes' }} icon={<Eye />} active={state?.trackChanges} onCommand={onCommand} />
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'accept-change' })} title="Accept"><Check /></button>
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'reject-change' })} title="Reject"><X /></button>
      </div>
      <div className="docx-toolbar__group-divider" />
      <div className="docx-toolbar__group">
        <button className="docx-toolbar__button" onClick={() => onCommand({ kind: 'open-comments-pane' })} title="Comments"><MessageSquare /></button>
      </div>
    </>
  );
};

// --- COMPONENTS ---

const FormatButton = ({ kind, payload, icon, active, onCommand }: { kind: string; payload?: ToolbarCommand; icon: React.ReactNode; active?: boolean; onCommand: (cmd: ToolbarCommand) => void }) => {
  return (
    <button
      className={`docx-toolbar__button ${active ? 'docx-toolbar__button--active' : ''}`}
      onClick={() => onCommand(payload || { kind } as ToolbarCommand)}
      title={kind}
    >
      {icon}
    </button>
  );
};

const ColorPickerPopover = ({ icon, onSelect, title }: { icon: React.ReactNode; onSelect: (color: string) => void; title: string }) => {
  const [open, setOpen] = useState(false);
  const [hex, setHex] = useState('');
  const ref = useRef<HTMLDivElement>(null);
  useClickOutside(ref, () => setOpen(false));

  return (
    <div className="docx-toolbar__popover-container" ref={ref}>
      <button className="docx-toolbar__button" onClick={() => setOpen(!open)} title={title}>
        {icon}
      </button>
      {open && (
        <div className="docx-toolbar__popover">
          <div className="docx-toolbar__color-grid">
            {DEFAULT_COLORS.map(c => (
              <div
                key={c}
                className="docx-toolbar__color-swatch"
                style={{ backgroundColor: c }}
                onClick={() => { onSelect(c); setOpen(false); }}
              />
            ))}
          </div>
          <div className="docx-toolbar__color-input-row">
            <input type="text" value={hex} onChange={e => setHex(e.target.value)} placeholder="#000000" />
            <button onClick={() => { if (hex) { onSelect(hex); setOpen(false); } }}>Ok</button>
          </div>
        </div>
      )}
    </div>
  );
};

const TablePickerPopover = ({ onCommand }: { onCommand: (cmd: ToolbarCommand) => void }) => {
  const [open, setOpen] = useState(false);
  const [hoverRow, setHoverRow] = useState(0);
  const [hoverCol, setHoverCol] = useState(0);
  const ref = useRef<HTMLDivElement>(null);
  useClickOutside(ref, () => setOpen(false));

  const rows = 8;
  const cols = 10;

  return (
    <div className="docx-toolbar__popover-container" ref={ref}>
      <button className="docx-toolbar__button" onClick={() => setOpen(!open)} title="Insert Table">
        <TableIcon />
      </button>
      {open && (
        <div className="docx-toolbar__popover">
          <div className="docx-toolbar__table-grid">
            {Array.from({ length: rows }).map((_, r) => (
              Array.from({ length: cols }).map((_, c) => {
                const isHovered = r <= hoverRow && c <= hoverCol;
                return (
                  <div
                    key={`${r}-${c}`}
                    className={`docx-toolbar__table-cell ${isHovered ? 'docx-toolbar__table-cell--hover' : ''}`}
                    onMouseEnter={() => { setHoverRow(r); setHoverCol(c); }}
                    onClick={() => {
                      onCommand({ kind: 'insert-table', rows: hoverRow + 1, cols: hoverCol + 1 });
                      setOpen(false);
                    }}
                  />
                );
              })
            ))}
          </div>
          <div className="docx-toolbar__table-label">
            {hoverRow + 1}x{hoverCol + 1} Table
          </div>
        </div>
      )}
    </div>
  );
};

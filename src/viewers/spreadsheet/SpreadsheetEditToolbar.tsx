/**
 * Shared editing toolbar segment for SpreadsheetViewer and CsvViewer (wave 3).
 *
 * Purely presentational + event plumbing — every action is a callback prop,
 * so this component has no opinion on what "insert a row" actually means to
 * the document model. Row/column actions operate on `selection` (the
 * currently active cell, tracked by the parent viewer — see that file's
 * comment on why selection tracking lives there and not in `useGridFind`);
 * they're disabled when nothing is selected rather than guessing a target.
 */
import { useEffect, useRef, useState } from 'react'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowLeftToLine,
  ArrowRightToLine,
  ArrowUpToLine,
  ArrowDownToLine,
  Bold,
  ClipboardPaste,
  Grid2x2,
  Italic,
  PaintBucket,
  Redo2,
  Save,
  Strikethrough,
  Trash2,
  Underline,
  Undo2,
} from 'lucide-react'

import { tsvToRows } from './spreadsheetClipboard'
import type { CellFormatPatch } from './spreadsheetDocument'
import type { CellBorder, ResolvedCellFormat } from './xlsxCellStyles'
import { useTranslate } from '../../i18n'

export type SaveFormatOption = {
  readonly id: string
  readonly label: string
}

export type SpreadsheetEditToolbarProps = {
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly onUndo: () => void
  readonly onRedo: () => void
  readonly selection: { readonly row: number; readonly col: number } | null
  readonly onInsertRowAbove: (row: number) => void
  readonly onDeleteRow: (row: number) => void
  readonly onInsertColumnLeft: (col: number) => void
  readonly onDeleteColumn: (col: number) => void
  readonly onPaste: (row: number, col: number, values: ReadonlyArray<ReadonlyArray<string>>) => void
  readonly onSave: () => void
  readonly saveFormats: ReadonlyArray<SaveFormatOption>
  readonly onSaveAs: (formatId: string) => void
  /**
   * SHEETFMT-2 — applies a formatting patch to the current selection.
   *
   * Omitted for a format with nowhere to store formatting (`.csv`/`.tsv` are
   * plain text), and the whole formatting group is then absent rather than
   * present and inert.
   */
  readonly onFormat?: (patch: CellFormatPatch) => void
  /**
   * The formatting the selected cell currently has, so a toggle can show its
   * state — a bold button that never looks pressed leaves the user guessing
   * whether it worked.
   */
  readonly selectionFormat?: ResolvedCellFormat
}

/**
 * The number formats offered, as the `formatCode` each one writes.
 *
 * Codes, not names: a `formatCode` is what `styles.xml` stores and what every
 * other spreadsheet program reads, so these round-trip to Excel exactly. The
 * empty code is General (no format), which is how a format is cleared.
 */
const NUMBER_FORMATS: ReadonlyArray<{ readonly code: string; readonly labelKey: string }> = [
  { code: '', labelKey: 'spreadsheetToolbar.numFmtGeneral' },
  { code: '0', labelKey: 'spreadsheetToolbar.numFmtInteger' },
  { code: '0.00', labelKey: 'spreadsheetToolbar.numFmtTwoDecimals' },
  { code: '#,##0.00', labelKey: 'spreadsheetToolbar.numFmtThousands' },
  { code: '0%', labelKey: 'spreadsheetToolbar.numFmtPercent' },
  { code: '0.00%', labelKey: 'spreadsheetToolbar.numFmtPercentTwo' },
  { code: 'yyyy-mm-dd', labelKey: 'spreadsheetToolbar.numFmtDate' },
  { code: 'yyyy-mm-dd hh:mm', labelKey: 'spreadsheetToolbar.numFmtDateTime' },
  { code: '@', labelKey: 'spreadsheetToolbar.numFmtText' },
]

/**
 * SHEETFMT-3 — the border presets offered.
 *
 * A short list of whole-cell shapes rather than Excel's full per-edge matrix:
 * these four cover what people reach for (box a block, underline a header,
 * clear it) and each is one unambiguous patch. Per-edge control needs a
 * different UI — a 3x3 grid of edge toggles — and is not here yet.
 *
 * The weight is `thin`, matching what Excel's own border buttons apply.
 */
const BORDER_PRESETS: ReadonlyArray<{
  readonly id: string
  readonly labelKey: string
  readonly border: CellBorder | null
}> = [
  {
    id: 'all',
    labelKey: 'spreadsheetToolbar.borderAll',
    border: {
      top: { weight: 'thin', color: undefined },
      right: { weight: 'thin', color: undefined },
      bottom: { weight: 'thin', color: undefined },
      left: { weight: 'thin', color: undefined },
    },
  },
  {
    id: 'bottom',
    labelKey: 'spreadsheetToolbar.borderBottom',
    border: { top: undefined, right: undefined, bottom: { weight: 'thin', color: undefined }, left: undefined },
  },
  {
    id: 'top',
    labelKey: 'spreadsheetToolbar.borderTop',
    border: { top: { weight: 'thin', color: undefined }, right: undefined, bottom: undefined, left: undefined },
  },
  { id: 'none', labelKey: 'spreadsheetToolbar.borderNone', border: null },
]

/** The swatches offered for a fill or a text colour. Office's own default palette row, which is what people expect to see. */
const COLOR_SWATCHES: ReadonlyArray<string> = [
  '#000000', '#FFFFFF', '#C00000', '#FF0000', '#FFC000', '#FFFF00',
  '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0',
]

/**
 * The border presets, in a popover.
 *
 * Shares `PopoverShell` with the colour pickers below so the dismissal
 * behaviour — outside click, Escape — is written once; a second copy of it is
 * exactly the sort of thing that drifts into one of them not closing.
 */
function BorderPopover({
  title,
  onSelect,
  disabled,
  t,
}: {
  readonly title: string
  readonly onSelect: (border: CellBorder | null) => void
  readonly disabled: boolean
  readonly t: (key: string) => string
}) {
  return (
    <PopoverShell icon={<Grid2x2 size={14} />} title={title} disabled={disabled}>
      {(close) => (
        <div className="spreadsheet-edit-toolbar__border-presets">
          {BORDER_PRESETS.map(preset => (
            <button
              key={preset.id}
              type="button"
              onClick={() => {
                onSelect(preset.border)
                close()
              }}
            >
              {t(preset.labelKey)}
            </button>
          ))}
        </div>
      )}
    </PopoverShell>
  )
}

/**
 * A popover trigger plus its dismissal behaviour, shared by the colour and
 * border pickers.
 */
function PopoverShell({
  icon,
  title,
  disabled,
  children,
}: {
  readonly icon: React.ReactNode
  readonly title: string
  readonly disabled: boolean
  readonly children: (close: () => void) => React.ReactNode
}) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    if (!open) return undefined
    const onDown = (event: MouseEvent): void => {
      if (ref.current && !ref.current.contains(event.target as Node)) setOpen(false)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false)
    }
    document.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])

  return (
    <div className="spreadsheet-edit-toolbar__popover-host" ref={ref}>
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        disabled={disabled}
        title={title}
        aria-label={title}
        aria-expanded={open}
        aria-haspopup="dialog"
      >
        {icon}
      </button>
      {open && (
        <div className="spreadsheet-edit-toolbar__popover" role="dialog" aria-label={title}>
          {children(() => setOpen(false))}
        </div>
      )}
    </div>
  )
}

/**
 * A colour swatch popover.
 *
 * Deliberately a plain `<button>` grid inside a `role="dialog"` rather than a
 * `role="menu"`: the A11Y work in this repo (pass 4) found that menu semantics
 * make a screen reader announce positional noise for a grid of swatches, and
 * real buttons are both simpler and better announced. Mirrors
 * `src/docx/editor/toolbar/Toolbar.tsx`'s own `ColorPickerPopover` for the same
 * reason.
 */
function ColorPopover({
  icon,
  title,
  clearLabel,
  onSelect,
  disabled,
}: {
  readonly icon: React.ReactNode
  readonly title: string
  readonly clearLabel: string
  readonly onSelect: (hex: string | null) => void
  readonly disabled: boolean
}) {
  return (
    <PopoverShell icon={icon} title={title} disabled={disabled}>
      {(close) => (
        <>
          <div className="spreadsheet-edit-toolbar__swatches">
            {COLOR_SWATCHES.map(hex => (
              <button
                key={hex}
                type="button"
                className="spreadsheet-edit-toolbar__swatch"
                style={{ background: hex }}
                aria-label={hex}
                title={hex}
                onClick={() => {
                  onSelect(hex)
                  close()
                }}
              />
            ))}
          </div>
          <button
            type="button"
            className="spreadsheet-edit-toolbar__clear"
            onClick={() => {
              onSelect(null)
              close()
            }}
          >
            {clearLabel}
          </button>
        </>
      )}
    </PopoverShell>
  )
}

/**
 * Reads the OS clipboard as TSV/plain text via the async Clipboard API
 * (user-gesture-gated — this is only ever called from a click handler), then
 * hands it to `spreadsheetClipboard.ts`'s own parser — the single source of
 * truth for TSV parsing, also used directly by its unit tests — rather than
 * re-implementing the same tab/newline splitting here.
 */
async function readClipboardRows(): Promise<string[][] | null> {
  try {
    const text = await navigator.clipboard.readText()
    // An empty clipboard is treated as "nothing to paste" (a no-op), not as
    // `tsvToRows`'s own "one empty cell" reading of an empty string — that
    // reading exists for a genuinely-empty-but-copied cell, not for a
    // clipboard with nothing in it at all.
    if (!text) return null
    return tsvToRows(text)
  } catch {
    // Clipboard read denied/unavailable — the grid's own built-in Ctrl+V
    // paste handling (glide-data-grid's `onPaste`) still works either way;
    // this button is a discoverable extra, not the only path.
    return null
  }
}

export function SpreadsheetEditToolbar({
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  selection,
  onInsertRowAbove,
  onDeleteRow,
  onInsertColumnLeft,
  onDeleteColumn,
  onPaste,
  onSave,
  saveFormats,
  onSaveAs,
  onFormat,
  selectionFormat,
}: SpreadsheetEditToolbarProps) {
  const t = useTranslate()
  const [saveFormat, setSaveFormat] = useState(saveFormats[0]?.id ?? '')

  const handlePasteClick = (): void => {
    void readClipboardRows().then((rows) => {
      if (rows) onPaste(selection?.row ?? 0, selection?.col ?? 0, rows)
    })
  }

  return (
    <div className="spreadsheet-edit-toolbar">
      <div className="spreadsheet-edit-toolbar__group">
        <button type="button" onClick={onUndo} disabled={!canUndo} title={t('spreadsheetToolbar.undoTitle')} aria-label={t('spreadsheetToolbar.undoAria')}>
          <Undo2 size={14} />
        </button>
        <button type="button" onClick={onRedo} disabled={!canRedo} title={t('spreadsheetToolbar.redoTitle')} aria-label={t('spreadsheetToolbar.redoAria')}>
          <Redo2 size={14} />
        </button>
      </div>

      {onFormat && (
        <div className="spreadsheet-edit-toolbar__group" role="group" aria-label={t('spreadsheetToolbar.formatGroupAria')}>
          {([
            ['bold', Bold, 'spreadsheetToolbar.bold'],
            ['italic', Italic, 'spreadsheetToolbar.italic'],
            ['underline', Underline, 'spreadsheetToolbar.underline'],
            ['strike', Strikethrough, 'spreadsheetToolbar.strike'],
          ] as const).map(([key, Icon, labelKey]) => {
            const on = selectionFormat?.[key] ?? false
            return (
              <button
                key={key}
                type="button"
                // A toggle, so its state is what `aria-pressed` is for — and
                // the click sends the OPPOSITE of the current state rather
                // than always `true`, which is what makes it a toggle rather
                // than a one-way switch.
                aria-pressed={on}
                className={on ? 'is-active' : undefined}
                onClick={() => onFormat({ [key]: !on })}
                disabled={!selection}
                title={t(labelKey)}
                aria-label={t(labelKey)}
              >
                <Icon size={14} />
              </button>
            )
          })}

          {([
            ['left', AlignLeft, 'spreadsheetToolbar.alignLeft'],
            ['center', AlignCenter, 'spreadsheetToolbar.alignCenter'],
            ['right', AlignRight, 'spreadsheetToolbar.alignRight'],
          ] as const).map(([align, Icon, labelKey]) => {
            const on = selectionFormat?.align === align
            return (
              <button
                key={align}
                type="button"
                aria-pressed={on}
                className={on ? 'is-active' : undefined}
                // Pressing the active alignment clears it, so there is a way
                // back to "whatever this column's default is".
                onClick={() => onFormat({ align: on ? null : align })}
                disabled={!selection}
                title={t(labelKey)}
                aria-label={t(labelKey)}
              >
                <Icon size={14} />
              </button>
            )
          })}

          <ColorPopover
            icon={<span className="spreadsheet-edit-toolbar__text-color-icon">A</span>}
            title={t('spreadsheetToolbar.textColor')}
            clearLabel={t('spreadsheetToolbar.clearColor')}
            disabled={!selection}
            onSelect={hex => onFormat({ color: hex })}
          />
          <ColorPopover
            icon={<PaintBucket size={14} />}
            title={t('spreadsheetToolbar.fillColor')}
            clearLabel={t('spreadsheetToolbar.clearFill')}
            disabled={!selection}
            onSelect={hex => onFormat({ fill: hex })}
          />
          <BorderPopover
            title={t('spreadsheetToolbar.borders')}
            disabled={!selection}
            onSelect={border => onFormat({ border })}
            t={t}
          />

          <select
            className="spreadsheet-edit-toolbar__numfmt"
            // Reads the SELECTION's own format, so it shows what the cell has
            // rather than a sticky last-chosen value.
            value={selectionFormat?.numberFormat ?? ''}
            onChange={event => onFormat({ numberFormat: event.target.value === '' ? null : event.target.value })}
            disabled={!selection}
            title={t('spreadsheetToolbar.numberFormat')}
            aria-label={t('spreadsheetToolbar.numberFormat')}
          >
            {/* An unrecognised format the FILE already uses has to be offered
                as an option, or picking anything else would be the only way to
                leave this control and the cell's own format would read as
                "General". */}
            {selectionFormat?.numberFormat !== undefined &&
              !NUMBER_FORMATS.some(f => f.code === selectionFormat.numberFormat) && (
                <option value={selectionFormat.numberFormat}>{selectionFormat.numberFormat}</option>
              )}
            {NUMBER_FORMATS.map(format => (
              <option key={format.code} value={format.code}>
                {t(format.labelKey)}
              </option>
            ))}
          </select>
        </div>
      )}

      <div className="spreadsheet-edit-toolbar__group">
        <button
          type="button"
          onClick={() => selection && onInsertRowAbove(selection.row)}
          disabled={!selection}
          title={t('spreadsheetToolbar.insertRowAboveTitle')}
          aria-label={t('spreadsheetToolbar.insertRowAboveAria')}
        >
          <ArrowUpToLine size={14} />
        </button>
        <button
          type="button"
          onClick={() => selection && onDeleteRow(selection.row)}
          disabled={!selection}
          title={t('spreadsheetToolbar.deleteRowTitle')}
          aria-label={t('spreadsheetToolbar.deleteRowAria')}
        >
          <ArrowDownToLine size={14} />
          <Trash2 size={10} />
        </button>
        <button
          type="button"
          onClick={() => selection && onInsertColumnLeft(selection.col)}
          disabled={!selection}
          title={t('spreadsheetToolbar.insertColumnLeftTitle')}
          aria-label={t('spreadsheetToolbar.insertColumnLeftAria')}
        >
          <ArrowLeftToLine size={14} />
        </button>
        <button
          type="button"
          onClick={() => selection && onDeleteColumn(selection.col)}
          disabled={!selection}
          title={t('spreadsheetToolbar.deleteColumnTitle')}
          aria-label={t('spreadsheetToolbar.deleteColumnAria')}
        >
          <ArrowRightToLine size={14} />
          <Trash2 size={10} />
        </button>
        <button type="button" onClick={handlePasteClick} title={t('spreadsheetToolbar.pasteTitle')} aria-label={t('spreadsheetToolbar.pasteAria')}>
          <ClipboardPaste size={14} />
        </button>
      </div>

      <div className="spreadsheet-edit-toolbar__group spreadsheet-edit-toolbar__save">
        <button type="button" onClick={onSave} title={t('spreadsheetToolbar.saveTitle')} aria-label={t('spreadsheetToolbar.saveAria')}>
          <Save size={14} />
          <span>{t('spreadsheetToolbar.saveAria')}</span>
        </button>
        {saveFormats.length > 1 && (
          <>
            <select
              aria-label={t('spreadsheetToolbar.saveAsFormatAria')}
              value={saveFormat}
              onChange={(e) => setSaveFormat(e.target.value)}
            >
              {saveFormats.map((f) => (
                <option key={f.id} value={f.id}>
                  {f.label}
                </option>
              ))}
            </select>
            <button
              type="button"
              onClick={() => onSaveAs(saveFormat)}
              title={t('spreadsheetToolbar.saveAsLabel')}
              aria-label={t('spreadsheetToolbar.saveAsAria')}
            >
              {t('spreadsheetToolbar.saveAsLabel')}
            </button>
          </>
        )}
      </div>
    </div>
  )
}

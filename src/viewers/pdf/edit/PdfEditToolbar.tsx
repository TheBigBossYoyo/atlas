/**
 * The PDF editing toolbar (PDF-EDIT-1), shown under the main PDF toolbar
 * while edit mode is on: tools, the active tool's colour and size, undo/redo,
 * page operations on the current page, and Save / Save As.
 */

import {
  ArrowDown,
  ArrowUp,
  Highlighter,
  MousePointer2,
  PenLine,
  Redo2,
  RotateCcw,
  RotateCw,
  Save,
  SaveAll,
  Trash2,
  Type,
  TextCursorInput,
  Undo2,
  X,
} from 'lucide-react'
import type { LucideIcon } from 'lucide-react'

import { useTranslate } from '../../../i18n'
import type { MessageKey } from '../../../i18n/messages.en'
import type { PdfEditTool, RgbColor } from './types'
import { EDIT_COLORS, type PdfEditor } from './usePdfEditor'

const TOOLS: ReadonlyArray<{ readonly tool: PdfEditTool; readonly Icon: LucideIcon; readonly labelKey: MessageKey }> = [
  { tool: 'select', Icon: MousePointer2, labelKey: 'pdf.edit.toolSelect' },
  { tool: 'text', Icon: Type, labelKey: 'pdf.edit.toolText' },
  { tool: 'highlight', Icon: Highlighter, labelKey: 'pdf.edit.toolHighlight' },
  { tool: 'ink', Icon: PenLine, labelKey: 'pdf.edit.toolInk' },
  { tool: 'edit-text', Icon: TextCursorInput, labelKey: 'pdf.edit.toolEditText' },
]

const FONT_SIZES = [8, 10, 12, 14, 18, 24, 32, 48]
const INK_WIDTHS = [1, 2, 3, 5, 8]

function sameColor(a: RgbColor, b: RgbColor): boolean {
  return a.r === b.r && a.g === b.g && a.b === b.b
}

export type PdfEditToolbarProps = {
  readonly editor: PdfEditor
  /** 0-based index of the page the page operations act on. */
  readonly currentPageIndex: number
  readonly pageCount: number
}

export function PdfEditToolbar({ editor, currentPageIndex, pageCount }: PdfEditToolbarProps) {
  const t = useTranslate()
  const { tool, setTool, activeColor, protection, fontTools } = editor
  const editsBlocked = protection.encrypted
  // In edit-text mode, colour, size and typeface apply to text added on empty space.
  const showsColor = true
  const showsFontSize = tool === 'text' || tool === 'select' || tool === 'edit-text'
  const showsFamily = tool === 'edit-text'
  const showsInkWidth = tool === 'ink'

  return (
    <div className="pdf-edit-toolbar" role="toolbar" aria-label={t('pdf.edit.toolbarAria')}>
      <div className="pdf-edit-toolbar__group" role="radiogroup" aria-label={t('pdf.edit.toolsAria')}>
        {TOOLS.map(({ tool: value, Icon, labelKey }) => (
          <button
            key={value}
            type="button"
            role="radio"
            aria-checked={tool === value}
            className={`pdf-viewer__toolbar-button${tool === value ? ' pdf-viewer__toolbar-button--active' : ''}`}
            title={t(labelKey)}
            aria-label={t(labelKey)}
            // Text box, highlight, ink and text editing all need pdf-lib's
            // font metrics and, for saving, a rewrite pdf-lib can't do on an
            // encrypted file.
            disabled={value !== 'select' && (editsBlocked || fontTools === null)}
            onClick={() => setTool(value)}
          >
            <Icon size={18} />
          </button>
        ))}
      </div>

      {showsColor && (
        <div className="pdf-edit-toolbar__group" role="radiogroup" aria-label={t('pdf.edit.colorAria')}>
          {EDIT_COLORS.map(({ key, color }) => (
            <button
              key={key}
              type="button"
              role="radio"
              aria-checked={sameColor(color, activeColor)}
              className={`pdf-edit-toolbar__swatch${sameColor(color, activeColor) ? ' pdf-edit-toolbar__swatch--active' : ''}`}
              style={{ backgroundColor: `rgb(${color.r * 255}, ${color.g * 255}, ${color.b * 255})` }}
              title={t(`pdf.edit.color.${key}` as MessageKey)}
              aria-label={t(`pdf.edit.color.${key}` as MessageKey)}
              onClick={() => editor.setToolColor(color)}
            />
          ))}
        </div>
      )}

      {showsFontSize && (
        <label className="pdf-edit-toolbar__size">
          <span className="pdf-edit-toolbar__size-label">{t('pdf.edit.fontSize')}</span>
          <select value={editor.fontSize} onChange={(event) => editor.setFontSize(Number(event.target.value))}>
            {FONT_SIZES.map((size) => (
              <option key={size} value={size}>
                {size}
              </option>
            ))}
          </select>
        </label>
      )}

      {showsFamily && (
        <label className="pdf-edit-toolbar__size">
          <span className="pdf-edit-toolbar__size-label">{t('pdf.edit.textFamily')}</span>
          <select value={editor.textFamily} onChange={(event) => editor.setTextFamily(event.target.value as typeof editor.textFamily)}>
            <option value="helvetica">{t('pdf.edit.family.sans')}</option>
            <option value="times">{t('pdf.edit.family.serif')}</option>
            <option value="courier">{t('pdf.edit.family.mono')}</option>
          </select>
        </label>
      )}

      {showsInkWidth && (
        <label className="pdf-edit-toolbar__size">
          <span className="pdf-edit-toolbar__size-label">{t('pdf.edit.inkWidth')}</span>
          <select value={editor.inkWidth} onChange={(event) => editor.setInkWidth(Number(event.target.value))}>
            {INK_WIDTHS.map((width) => (
              <option key={width} value={width}>
                {width}
              </option>
            ))}
          </select>
        </label>
      )}

      <div className="pdf-edit-toolbar__group">
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.undo')}
          aria-label={t('pdf.edit.undo')}
          disabled={!editor.canUndo}
          onClick={editor.undo}
        >
          <Undo2 size={18} />
        </button>
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.redo')}
          aria-label={t('pdf.edit.redo')}
          disabled={!editor.canRedo}
          onClick={editor.redo}
        >
          <Redo2 size={18} />
        </button>
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.deleteSelection')}
          aria-label={t('pdf.edit.deleteSelection')}
          disabled={editor.selectedId === null}
          onClick={editor.removeSelected}
        >
          <X size={18} />
        </button>
      </div>

      <div className="pdf-edit-toolbar__group" role="group" aria-label={t('pdf.edit.pageGroupAria', { n: currentPageIndex + 1 })}>
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.rotatePageLeft')}
          aria-label={t('pdf.edit.rotatePageLeft')}
          disabled={editsBlocked}
          onClick={() => editor.rotatePageAt(currentPageIndex, -90)}
        >
          <RotateCcw size={18} />
        </button>
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.rotatePageRight')}
          aria-label={t('pdf.edit.rotatePageRight')}
          disabled={editsBlocked}
          onClick={() => editor.rotatePageAt(currentPageIndex, 90)}
        >
          <RotateCw size={18} />
        </button>
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.movePageUp')}
          aria-label={t('pdf.edit.movePageUp')}
          disabled={editsBlocked || currentPageIndex <= 0}
          onClick={() => editor.movePageAt(currentPageIndex, -1)}
        >
          <ArrowUp size={18} />
        </button>
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.movePageDown')}
          aria-label={t('pdf.edit.movePageDown')}
          disabled={editsBlocked || currentPageIndex >= pageCount - 1}
          onClick={() => editor.movePageAt(currentPageIndex, 1)}
        >
          <ArrowDown size={18} />
        </button>
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.deletePage')}
          aria-label={t('pdf.edit.deletePage')}
          disabled={editsBlocked || pageCount <= 1}
          onClick={() => editor.deletePageAt(currentPageIndex)}
        >
          <Trash2 size={18} />
        </button>
      </div>

      <div className="pdf-edit-toolbar__group pdf-edit-toolbar__group--end">
        <button
          type="button"
          className="pdf-edit-toolbar__save"
          title={t('pdf.edit.saveTitle')}
          disabled={editor.isSaving || !editor.isDirty}
          onClick={() => void editor.save()}
        >
          <Save size={16} />
          <span>{editor.isSaving ? t('pdf.edit.saving') : t('pdf.edit.save')}</span>
        </button>
        <button
          type="button"
          className="pdf-viewer__toolbar-button"
          title={t('pdf.edit.saveAs')}
          aria-label={t('pdf.edit.saveAs')}
          disabled={editor.isSaving}
          onClick={() => void editor.saveAs()}
        >
          <SaveAll size={18} />
        </button>
      </div>

      {(editor.message || protection.encrypted || protection.signed || tool === 'edit-text') && (
        <div className="pdf-edit-toolbar__notice" role={editor.message ? 'alert' : 'status'}>
          {editor.message ??
            (protection.encrypted
              ? t('pdf.edit.encryptedNotice')
              : protection.signed
                ? t('pdf.edit.signedNotice')
                : t('pdf.edit.editTextNotice'))}
        </div>
      )}
    </div>
  )
}

/** USR-16 — slide editing commands shown in the deck toolbar when the deck is editable. */
import { memo, type ReactNode } from 'react'
import {
  ArrowDown,
  ArrowUp,
  Copy,
  Plus,
  Redo2,
  Save,
  SaveAll,
  Trash2,
  Type,
  Undo2,
} from 'lucide-react'

export type SlideDeckEditor = {
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly onUndo: () => void
  readonly onRedo: () => void
  readonly onAddSlide: () => void
  readonly onDuplicateSlide: () => void
  readonly onDeleteSlide: () => void
  readonly onMoveSlide: (delta: -1 | 1) => void
  readonly onInsertTextBox: () => Promise<string | null>
  readonly onSave: () => void
  readonly onSaveAs: () => void
  readonly canEditNotes: boolean
  readonly onNotesChange: (text: string) => void
  readonly onShapeText: (sourceId: string, text: string) => void
  readonly onShapeBox: (sourceId: string, box: { x: number; y: number; w: number; h: number }) => Promise<void>
  readonly onDeleteShape: (sourceId: string) => void
  readonly saveError: string | null
}

type SlideEditToolbarProps = {
  readonly editor: SlideDeckEditor
  readonly slideIndex: number
  readonly slideCount: number
  readonly onTextBoxInserted: (sourceId: string) => void
}

function ToolButton({
  label,
  onClick,
  disabled = false,
  children,
}: {
  readonly label: string
  readonly onClick: () => void
  readonly disabled?: boolean
  readonly children: ReactNode
}) {
  return (
    <button type="button" className="slide-deck__toolbar-button" aria-label={label} title={label} disabled={disabled} onClick={onClick}>
      {children}
    </button>
  )
}

function SlideEditToolbarBase({ editor, slideIndex, slideCount, onTextBoxInserted }: SlideEditToolbarProps) {
  return (
    <div className="slide-deck__toolbar-group" role="group" aria-label="Edit presentation">
      <ToolButton label="Undo" onClick={editor.onUndo} disabled={!editor.canUndo}>
        <Undo2 size={16} />
      </ToolButton>
      <ToolButton label="Redo" onClick={editor.onRedo} disabled={!editor.canRedo}>
        <Redo2 size={16} />
      </ToolButton>
      <span className="slide-deck__toolbar-divider" aria-hidden="true" />
      <ToolButton label="New slide" onClick={editor.onAddSlide}>
        <Plus size={16} />
      </ToolButton>
      <ToolButton label="Duplicate slide" onClick={editor.onDuplicateSlide}>
        <Copy size={16} />
      </ToolButton>
      <ToolButton label="Delete slide" onClick={editor.onDeleteSlide} disabled={slideCount <= 1}>
        <Trash2 size={16} />
      </ToolButton>
      <ToolButton label="Move slide up" onClick={() => editor.onMoveSlide(-1)} disabled={slideIndex <= 0}>
        <ArrowUp size={16} />
      </ToolButton>
      <ToolButton label="Move slide down" onClick={() => editor.onMoveSlide(1)} disabled={slideIndex >= slideCount - 1}>
        <ArrowDown size={16} />
      </ToolButton>
      <span className="slide-deck__toolbar-divider" aria-hidden="true" />
      <ToolButton
        label="Insert text box"
        onClick={() => {
          void editor.onInsertTextBox().then((sourceId) => {
            if (sourceId) onTextBoxInserted(sourceId)
          })
        }}
      >
        <Type size={16} />
      </ToolButton>
      <span className="slide-deck__toolbar-divider" aria-hidden="true" />
      <ToolButton label="Save As" onClick={editor.onSaveAs}>
        <SaveAll size={16} />
      </ToolButton>
      <button type="button" className="slide-deck__save-button" aria-label="Save" title="Save (Ctrl+S)" onClick={editor.onSave}>
        <Save size={14} />
        Save
      </button>
    </div>
  )
}

export const SlideEditToolbar = memo(SlideEditToolbarBase)

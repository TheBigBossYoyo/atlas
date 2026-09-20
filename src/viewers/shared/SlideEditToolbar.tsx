/** USR-16 — slide editing commands shown in the deck toolbar when the deck is editable. */
import { memo, useCallback, useRef, type KeyboardEvent, type ReactNode } from 'react'
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
import { useTranslate } from '../../i18n'

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
  /** Fires synchronously at click time, before the (async) insert even starts — see `handleInsertTextBox`. */
  readonly onInsertTextBoxStart: () => void
  /** Fires once the insert settles, with the new shape's id, or `null` if it failed. */
  readonly onTextBoxInserted: (sourceId: string | null) => void
}

function ToolButton({
  label,
  onClick,
  onKeyDown,
  disabled = false,
  children,
}: {
  readonly label: string
  readonly onClick: () => void
  readonly onKeyDown?: (event: KeyboardEvent<HTMLButtonElement>) => void
  readonly disabled?: boolean
  readonly children: ReactNode
}) {
  return (
    <button
      type="button"
      className="slide-deck__toolbar-button"
      aria-label={label}
      title={label}
      disabled={disabled}
      onClick={onClick}
      onKeyDown={onKeyDown}
    >
      {children}
    </button>
  )
}

function SlideEditToolbarBase({
  editor,
  slideIndex,
  slideCount,
  onInsertTextBoxStart,
  onTextBoxInserted,
}: SlideEditToolbarProps) {
  const t = useTranslate()

  // Insert Text Box is meant to be typed into right away, with no rescuing
  // click. Three things stand between a click and the keystrokes landing:
  //
  // 1. The button keeps browser focus after the click that triggered it.
  //    Typing right into the new box is the whole point of this button, and
  //    "Hello world" contains a space — which, on a still-*focused* button,
  //    the browser turns into a native click, firing `onInsertTextBox()` a
  //    second time and leaving a duplicate, empty text box on the slide.
  //    (Guarding re-entrancy in the handler below does NOT fix this: the
  //    first insert's queued edit + re-parse routinely finishes well within
  //    the time it takes to type a few characters, so the guard has already
  //    cleared by the time Space arrives — confirmed by instrumenting this
  //    handler and watching two genuinely separate calls land, each with
  //    its own resolved sourceId.) Blurring the button on click was tried
  //    and rejected: it leaves focus on nothing while the re-parse is in
  //    flight, and a Space typed into that gap is caught by the slide
  //    deck's own next-slide-on-Space shortcut instead — confirmed by
  //    watching the deck actually navigate to slide 2 mid-repro. So instead
  //    we keep focus on the button and swallow the Space/Enter keydown
  //    ourselves (below), which stops the browser from ever synthesizing
  //    that second click, before it happens rather than after.
  // 2. Even a single insert races the caret: the promise resolves with the
  //    new shape's id before the re-parsed slide containing that shape has
  //    committed to `SlideDeck`'s `slides` — see `pendingTextBoxSourceId`
  //    there, which resolves the id once the shape actually exists instead
  //    of trusting a same-tick lookup.
  // 3. Even with (1) and (2) fixed, the insert + re-parse round trip is
  //    still asynchronous, and under enough load (e.g. a shared dev machine
  //    running several other heavy processes) it can take longer than the
  //    gap between keystrokes from a fast typist. Rather than hope that
  //    race resolves in our favor, `onInsertTextBoxStart` (fired
  //    synchronously, before the async call below even starts) tells
  //    `SlideDeck` to start buffering keystrokes itself from the very first
  //    one, and hand the buffer to the new shape's editor as its seed text
  //    once it actually mounts — so no keystroke typed from the moment of
  //    the click onward is ever lost, regardless of how long the round trip
  //    takes.
  const insertPendingRef = useRef(false)
  const handleInsertTextBox = useCallback(() => {
    if (insertPendingRef.current) return
    insertPendingRef.current = true
    onInsertTextBoxStart()
    void editor
      .onInsertTextBox()
      .then((sourceId) => {
        onTextBoxInserted(sourceId)
      })
      .finally(() => {
        insertPendingRef.current = false
      })
  }, [editor, onInsertTextBoxStart, onTextBoxInserted])

  const suppressButtonActivationKeys = useCallback((event: KeyboardEvent<HTMLButtonElement>) => {
    if (event.key === ' ' || event.key === 'Enter') {
      event.preventDefault()
      event.stopPropagation()
    }
  }, [])

  return (
    <div className="slide-deck__toolbar-group" role="group" aria-label={t('slides.editToolbar.editPresentationAria')}>
      <ToolButton label={t('slides.editToolbar.undo')} onClick={editor.onUndo} disabled={!editor.canUndo}>
        <Undo2 size={16} />
      </ToolButton>
      <ToolButton label={t('slides.editToolbar.redo')} onClick={editor.onRedo} disabled={!editor.canRedo}>
        <Redo2 size={16} />
      </ToolButton>
      <span className="slide-deck__toolbar-divider" aria-hidden="true" />
      <ToolButton label={t('slides.editToolbar.newSlide')} onClick={editor.onAddSlide}>
        <Plus size={16} />
      </ToolButton>
      <ToolButton label={t('slides.editToolbar.duplicateSlide')} onClick={editor.onDuplicateSlide}>
        <Copy size={16} />
      </ToolButton>
      <ToolButton label={t('slides.editToolbar.deleteSlide')} onClick={editor.onDeleteSlide} disabled={slideCount <= 1}>
        <Trash2 size={16} />
      </ToolButton>
      <ToolButton label={t('slides.editToolbar.moveSlideUp')} onClick={() => editor.onMoveSlide(-1)} disabled={slideIndex <= 0}>
        <ArrowUp size={16} />
      </ToolButton>
      <ToolButton label={t('slides.editToolbar.moveSlideDown')} onClick={() => editor.onMoveSlide(1)} disabled={slideIndex >= slideCount - 1}>
        <ArrowDown size={16} />
      </ToolButton>
      <span className="slide-deck__toolbar-divider" aria-hidden="true" />
      <ToolButton
        label={t('slides.editToolbar.insertTextBox')}
        onClick={handleInsertTextBox}
        onKeyDown={suppressButtonActivationKeys}
      >
        <Type size={16} />
      </ToolButton>
      <span className="slide-deck__toolbar-divider" aria-hidden="true" />
      <ToolButton label={t('slides.editToolbar.saveAs')} onClick={editor.onSaveAs}>
        <SaveAll size={16} />
      </ToolButton>
      <button type="button" className="slide-deck__save-button" aria-label={t('slides.editToolbar.save')} title={t('slides.editToolbar.saveTitle')} onClick={editor.onSave}>
        <Save size={14} />
        {t('slides.editToolbar.save')}
      </button>
    </div>
  )
}

export const SlideEditToolbar = memo(SlideEditToolbarBase)

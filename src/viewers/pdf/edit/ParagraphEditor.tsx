/**
 * The in-place editor for a paragraph of the page's own text (PDF-EDIT-2).
 *
 * A contentEditable box laid over the paragraph, filled with its text as
 * styled spans — each span carries its style index, so a bold word stays bold
 * however the text around it changes. Ctrl+B / Ctrl+I toggle bold and italic
 * on a selection; the result is read back as spans, and a span whose weight or
 * slant no longer matches its style gets a derived style (written in the
 * matching bundled font, since the document may not have that variant).
 */

import { useEffect, useRef } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'

import { fillEditor, readEditor, type ParagraphEditorValue } from './paragraphEditorDom'

export type { ParagraphEditorValue } from './paragraphEditorDom'

export type ParagraphEditorProps = {
  readonly initial: ParagraphEditorValue
  /** Screen box (CSS px, relative to the page) the editor covers. */
  readonly box: { readonly left: number; readonly top: number; readonly width: number; readonly height: number }
  /** CSS px per PDF point at the current zoom. */
  readonly scale: number
  /** Turn on screen, degrees clockwise — the view's rotation relative to the text. */
  readonly screenTurn: number
  /** Whether lines wrap at the box's width (a paragraph) or run on (a heading, new text). */
  readonly wrap: boolean
  readonly onBlur: () => void
  readonly label: string
  readonly onCommit: (value: ParagraphEditorValue) => void
  readonly onCancel: () => void
  /** Hands the parent a way to read the current value (for committing from outside). */
  readonly registerRead: (read: (() => ParagraphEditorValue) | null) => void
}

export function ParagraphEditor({ initial, box, scale, screenTurn, wrap, label, onCommit, onCancel, onBlur, registerRead }: ParagraphEditorProps) {
  const rootRef = useRef<HTMLDivElement | null>(null)
  const initialRef = useRef(initial)
  const scaleRef = useRef(scale)
  // The content as of the last change, kept for a read after unmount: when
  // the page scrolls out of the render window, React detaches `rootRef`
  // before the parent's cleanup asks for what was typed.
  const latestRef = useRef(initial)

  // Filled once, imperatively: contentEditable content is the browser's while
  // the user types, and re-rendering it from React would lose the caret.
  useEffect(() => {
    const root = rootRef.current
    if (!root) return
    fillEditor(root, initialRef.current, scaleRef.current)
    root.focus()
    const selection = window.getSelection()
    if (selection) {
      selection.selectAllChildren(root)
      selection.collapseToEnd()
    }
  }, [])

  useEffect(() => {
    registerRead(() => {
      const root = rootRef.current
      return root ? readEditor(root, initialRef.current) : latestRef.current
    })
    return () => registerRead(null)
  }, [registerRead])

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const root = rootRef.current
    if (!root) return
    const mod = event.ctrlKey || event.metaKey
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onCancel()
    } else if (event.key === 'Enter' && mod) {
      event.preventDefault()
      event.stopPropagation()
      onCommit(readEditor(root, initialRef.current))
    } else if (event.key === 'Enter') {
      event.preventDefault()
      document.execCommand('insertLineBreak')
    } else if (event.key === 'Tab') {
      event.preventDefault()
      document.execCommand('insertText', false, '\t')
    } else if (mod && (event.key.toLowerCase() === 'b' || event.key.toLowerCase() === 'i')) {
      event.preventDefault()
      event.stopPropagation()
      document.execCommand(event.key.toLowerCase() === 'b' ? 'bold' : 'italic')
    } else if (mod && ['s', 'z', 'y'].includes(event.key.toLowerCase())) {
      // Let Ctrl+S reach the shell (it flushes this editor first); keep the
      // editor's own undo/redo for typing.
      if (event.key.toLowerCase() !== 's') event.stopPropagation()
    }
  }

  return (
    <div
      ref={rootRef}
      role="textbox"
      aria-multiline="true"
      aria-label={label}
      contentEditable
      suppressContentEditableWarning
      spellCheck
      className="pdf-edit-layer__paragraph-editor"
      style={{
        left: box.left,
        top: box.top,
        width: wrap ? box.width : undefined,
        minWidth: wrap ? undefined : Math.max(box.width, 40),
        minHeight: box.height,
        whiteSpace: wrap ? 'pre-wrap' : 'pre',
        transform: Math.abs(screenTurn) < 0.01 ? undefined : `rotate(${screenTurn}deg)`,
        tabSize: 4,
      }}
      onKeyDown={handleKeyDown}
      onInput={() => {
        if (rootRef.current) latestRef.current = readEditor(rootRef.current, initialRef.current)
      }}
      onBlur={onBlur}
      onPaste={(event) => {
        event.preventDefault()
        document.execCommand('insertText', false, event.clipboardData.getData('text/plain'))
      }}
    />
  )
}

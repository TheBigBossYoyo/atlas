import { act, fireEvent, render, screen, waitFor } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

import { PdfEditLayer } from '../PdfEditLayer'
import type { PdfPageOverlayInfo } from '../../PdfPage'
import type { PdfAnnotationDraft, PdfEditTool, TextEdit } from '../types'
import type { SpanStyle, TextParagraph } from '../content/paragraphs'
import { editParagraph } from '../content/textEdits'
import type { PdfEditor } from '../usePdfEditor'

/**
 * A 1:1 viewport with the y axis flipped, like pdf.js's at scale 1 on a
 * 1000pt-tall page: viewport (x, y) <-> PDF (x, 1000 - y). jsdom lays every
 * element out at (0, 0), so a pointer's clientX/clientY ARE viewport
 * coordinates here.
 */
const viewport = {
  width: 600,
  height: 1000,
  scale: 1,
  rotation: 0,
  transform: [1, 0, 0, -1, 0, 1000],
  convertToPdfPoint: (x: number, y: number) => [x, 1000 - y],
  convertToViewportPoint: (x: number, y: number) => [x, 1000 - y],
} as unknown as PdfPageOverlayInfo['viewport']

function info(overrides: Partial<PdfPageOverlayInfo> = {}): PdfPageOverlayInfo {
  return {
    sourcePageNumber: 1,
    viewport,
    intrinsicRotation: 0,
    textContent: {
      items: [{ str: 'Old heading', transform: [20, 0, 0, 20, 100, 900], width: 120, height: 20, fontName: 'f1', dir: 'ltr', hasEOL: false }],
      styles: { f1: { fontFamily: 'serif', ascent: 0.9, descent: -0.2, vertical: false } },
      lang: null,
    } as unknown as PdfPageOverlayInfo['textContent'],
    fontNames: new Map([['f1', 'Georgia-Bold']]),
    canvas: null,
    dpr: 1,
    ...overrides,
  }
}

function fakeEditor(tool: PdfEditTool, overrides: Partial<PdfEditor> = {}): PdfEditor {
  return {
    isEditing: true,
    tool,
    selectedId: null,
    setSelectedId: vi.fn(),
    toolColors: { text: { r: 0, g: 0, b: 0 }, highlight: { r: 1, g: 1, b: 0 }, ink: { r: 1, g: 0, b: 0 } },
    fontSize: 12,
    inkWidth: 2,
    fontTools: { widthOf: (_font: unknown, text: string, size: number) => text.length * size * 0.5, canEncode: (_f: unknown, text: string) => [...text].every((ch) => (ch.codePointAt(0) ?? 0) <= 0xff) },
    setMessage: vi.fn(),
    addDraft: vi.fn(),
    updateDraft: vi.fn(),
    removeById: vi.fn(),
    editText: vi.fn(),
    textFamily: 'helvetica',
    registerFlush: undefined,
    textAnalysis: {
      paragraphs: vi.fn(async () => [PARAGRAPH, BLOCKED]),
      undrawable: vi.fn(async () => null),
    },
    ...overrides,
  } as unknown as PdfEditor
}

const STYLE: SpanStyle = {
  fontKey: '5 0 R',
  fontResource: 'F1',
  font: { family: 'times', bold: false, italic: false },
  size: 20,
  fill: { space: null, color: '0 0 1 rg' },
  charSpacing: 0,
  wordSpacing: 0,
  horizontalScale: 1,
  renderMode: 0,
}

/** A one-line paragraph at y = 900 (screen y = 100), x 100..220. */
const PARAGRAPH: TextParagraph = {
  id: 'page:7:0',
  streamKey: 'page',
  pieces: [{ opIndex: 7, element: 0 }],
  insertAfterOp: 8,
  origin: { x: 100, y: 900 },
  direction: { x: 1, y: 0 },
  leading: 24,
  width: 120,
  lineCount: 1,
  align: 'left',
  firstLineIndent: 0,
  tabStops: [],
  styles: [STYLE, { ...STYLE, font: { ...STYLE.font, bold: true } }],
  spans: [{ text: 'Old ', style: 0 }, { text: 'heading', style: 1 }],
  bounds: [{ x: 100, y: 919 }, { x: 220, y: 919 }, { x: 220, y: 895 }, { x: 100, y: 895 }],
  blocked: null,
}

/** Scanner text at y = 500 (screen y = 500). */
const BLOCKED: TextParagraph = {
  ...PARAGRAPH,
  id: 'page:20:0',
  bounds: [{ x: 100, y: 520 }, { x: 300, y: 520 }, { x: 300, y: 495 }, { x: 100, y: 495 }],
  blocked: 'invisible',
}

function layer(): HTMLElement {
  return document.querySelector('.pdf-edit-layer') as HTMLElement
}

describe('PdfEditLayer', () => {
  it('draws pending annotations in PDF space', () => {
    const annotations: PdfAnnotationDraft[] = [
      { kind: 'highlight', id: 'h', page: 1, rects: [{ x: 10, y: 10, width: 50, height: 10 }], color: { r: 1, g: 1, b: 0 } },
      { kind: 'ink', id: 'i', page: 1, points: [{ x: 0, y: 0 }, { x: 5, y: 5 }], color: { r: 1, g: 0, b: 0 }, width: 2 },
      { kind: 'text', id: 't', page: 1, anchor: { x: 50, y: 500 }, rotation: 0, text: 'Note\nline two', fontSize: 12, color: { r: 0, g: 0, b: 1 } },
    ]
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={annotations} textEdits={[]} previewCurrent={false} editor={fakeEditor('select')} />)

    expect(document.querySelectorAll('.pdf-edit-layer__highlights rect')).toHaveLength(1)
    expect(document.querySelectorAll('.pdf-edit-layer__svg polyline')).toHaveLength(2) // stroke + hit area
    const texts = Array.from(document.querySelectorAll('.pdf-edit-layer__svg text')).map((node) => node.textContent)
    expect(texts).toEqual(['Note', 'line two'])
    expect(document.querySelector('.pdf-edit-layer__svg > g')?.getAttribute('transform')).toBe('matrix(1 0 0 -1 0 1000)')
  })

  it('sketches a text edit the preview does not have yet, and leaves it to the preview once it does', () => {
    const edit: TextEdit = editParagraph('e', 1, PARAGRAPH, [{ text: 'New heading', style: 0 }])
    const { rerender } = render(
      <PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[edit]} previewCurrent={false} editor={fakeEditor('select')} />,
    )
    const sketched = () => Array.from(document.querySelectorAll('.pdf-edit-layer__svg text')).map((node) => node.textContent)
    expect(sketched()).toEqual(['New heading'])
    rerender(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[edit]} previewCurrent editor={fakeEditor('select')} />)
    expect(sketched()).toEqual([])
    // Still selectable (to delete or re-open) through its area.
    expect(document.querySelector('[data-edit-id="e"]')).not.toBeNull()
  })

  it('renders nothing when the viewport has no transform (pdf.js not ready)', () => {
    const { container } = render(
      <PdfEditLayer
        info={info({ viewport: { ...viewport, transform: undefined } as unknown as PdfPageOverlayInfo['viewport'] })}
        documentRotation={0}
        annotations={[]}
        textEdits={[]}
        previewCurrent={false}
        editor={fakeEditor('text')}
      />,
    )
    expect(container).toBeEmptyDOMElement()
  })

  it('text tool: a click opens a box at that point, Ctrl+Enter adds it — once, even when blur follows', () => {
    const editor = fakeEditor('text')
    render(<PdfEditLayer info={info()} documentRotation={90} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)

    fireEvent.pointerDown(layer(), { button: 0, clientX: 40, clientY: 300 })
    const box = screen.getByRole('textbox', { name: 'Text box' })
    fireEvent.change(box, { target: { value: 'Hello' } })
    fireEvent.keyDown(box, { key: 'Enter', ctrlKey: true })
    fireEvent.blur(box)

    expect(editor.addDraft).toHaveBeenCalledTimes(1)
    expect(editor.addDraft).toHaveBeenCalledWith(
      expect.objectContaining({ kind: 'text', text: 'Hello', anchor: { x: 40, y: 700 }, rotation: 90, page: 1, fontSize: 12 }),
    )
    expect(screen.queryByRole('textbox', { name: 'Text box' })).toBeNull()
  })

  it('text tool: refuses characters the standard fonts cannot draw and keeps the box open', () => {
    const editor = fakeEditor('text')
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)
    fireEvent.pointerDown(layer(), { button: 0, clientX: 10, clientY: 10 })
    const box = screen.getByRole('textbox', { name: 'Text box' })
    fireEvent.change(box, { target: { value: 'Привет' } })
    fireEvent.keyDown(box, { key: 'Enter', ctrlKey: true })
    expect(editor.addDraft).not.toHaveBeenCalled()
    expect(editor.setMessage).toHaveBeenCalledWith(expect.stringContaining('fonts available'))
    expect(screen.getByRole('textbox', { name: 'Text box' })).toBeInTheDocument()

    fireEvent.keyDown(box, { key: 'Escape' })
    expect(screen.queryByRole('textbox', { name: 'Text box' })).toBeNull()
  })

  it('an empty new text box is discarded, and emptying an existing one deletes it', () => {
    const existing: PdfAnnotationDraft = {
      kind: 'text',
      id: 'old',
      page: 1,
      anchor: { x: 50, y: 500 },
      rotation: 0,
      text: 'Bye',
      fontSize: 12,
      color: { r: 0, g: 0, b: 0 },
    }
    const editor = fakeEditor('text')
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[existing]} textEdits={[]} previewCurrent={false} editor={editor} />)

    fireEvent.pointerDown(layer(), { button: 0, clientX: 10, clientY: 10 })
    fireEvent.blur(screen.getByRole('textbox', { name: 'Text box' }))
    expect(editor.addDraft).not.toHaveBeenCalled()

    fireEvent.doubleClick(document.querySelector('[data-edit-id="old"]') as Element)
    const box = screen.getByRole('textbox', { name: 'Text box' })
    expect(box).toHaveValue('Bye')
    fireEvent.change(box, { target: { value: '' } })
    fireEvent.blur(box)
    expect(editor.removeById).toHaveBeenCalledWith('old')
  })

  it('ink tool: a drag becomes one stroke, with points that add nothing dropped', () => {
    const editor = fakeEditor('ink')
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)
    fireEvent.pointerDown(layer(), { button: 0, clientX: 10, clientY: 10 })
    fireEvent.pointerMove(layer(), { clientX: 10.2, clientY: 10 })
    fireEvent.pointerMove(layer(), { clientX: 20, clientY: 20 })
    fireEvent.pointerUp(layer())
    expect(editor.addDraft).toHaveBeenCalledWith(
      expect.objectContaining({ kind: 'ink', width: 2, points: [{ x: 10, y: 990 }, { x: 20, y: 980 }] }),
    )
  })

  it('edit-text tool: clicking a paragraph opens it in place with its styled spans, and an untouched commit changes nothing', async () => {
    const editor = fakeEditor('edit-text')
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)
    await waitFor(() => expect(editor.textAnalysis?.paragraphs).toHaveBeenCalledWith(1))
    await act(async () => {})

    // Hovering outlines the paragraph under the pointer.
    fireEvent.pointerMove(layer(), { clientX: 150, clientY: 95 })
    expect(document.querySelector('.pdf-edit-layer__hover')).not.toBeNull()

    fireEvent.pointerDown(layer(), { button: 0, clientX: 150, clientY: 95 })
    const box = await screen.findByRole('textbox', { name: 'Paragraph text' })
    expect(box.textContent).toBe('Old heading')
    expect((box.querySelector('[data-style="1"]') as HTMLElement).style.fontWeight).toBe('700')

    await act(async () => {
      fireEvent.keyDown(box, { key: 'Enter', ctrlKey: true })
    })
    expect(editor.editText).toHaveBeenCalledWith(expect.objectContaining({ paragraphId: 'page:7:0' }), true)
  })

  it('edit-text tool: a changed paragraph is committed as a real edit of its operators', async () => {
    const editor = fakeEditor('edit-text')
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)
    await waitFor(() => expect(editor.textAnalysis?.paragraphs).toHaveBeenCalled())
    await act(async () => {})
    fireEvent.pointerDown(layer(), { button: 0, clientX: 150, clientY: 95 })
    const box = await screen.findByRole('textbox', { name: 'Paragraph text' })
    ;(box.querySelector('[data-style="0"]') as HTMLElement).textContent = 'New '

    await act(async () => {
      fireEvent.keyDown(box, { key: 'Enter', ctrlKey: true })
    })
    expect(editor.editText).toHaveBeenCalledWith(
      expect.objectContaining({
        paragraphId: 'page:7:0',
        pieces: [{ opIndex: 7, element: 0 }],
        insertAfterOp: 8,
        spans: [{ text: 'New ', style: 0 }, { text: 'heading', style: 1 }],
      }),
      false,
    )
  })

  it('edit-text tool: refuses characters no font can draw and keeps the editor open with what was typed', async () => {
    const editor = fakeEditor('edit-text', {
      textAnalysis: { paragraphs: vi.fn(async () => [PARAGRAPH]), undrawable: vi.fn(async () => '漢') },
    } as Partial<PdfEditor>)
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)
    await waitFor(() => expect(editor.textAnalysis?.paragraphs).toHaveBeenCalled())
    await act(async () => {})
    fireEvent.pointerDown(layer(), { button: 0, clientX: 150, clientY: 95 })
    const box = await screen.findByRole('textbox', { name: 'Paragraph text' })
    ;(box.querySelector('[data-style="0"]') as HTMLElement).textContent = '漢 '
    await act(async () => {
      fireEvent.keyDown(box, { key: 'Enter', ctrlKey: true })
    })
    expect(editor.editText).not.toHaveBeenCalled()
    expect(editor.setMessage).toHaveBeenCalledWith(expect.stringContaining('漢'))
    expect((await screen.findByRole('textbox', { name: 'Paragraph text' })).textContent).toBe('漢 heading')
  })

  it('edit-text tool: what was typed is committed when the page scrolls out of the render window mid-edit', async () => {
    const editor = fakeEditor('edit-text')
    const { unmount } = render(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)
    await waitFor(() => expect(editor.textAnalysis?.paragraphs).toHaveBeenCalled())
    await act(async () => {})
    fireEvent.pointerDown(layer(), { button: 0, clientX: 150, clientY: 95 })
    const box = await screen.findByRole('textbox', { name: 'Paragraph text' })
    ;(box.querySelector('[data-style="0"]') as HTMLElement).textContent = 'Typed '
    fireEvent.input(box)

    // The page leaves the window: the layer unmounts with the editor still open.
    await act(async () => {
      unmount()
    })
    await waitFor(() =>
      expect(editor.editText).toHaveBeenCalledWith(
        expect.objectContaining({ spans: [{ text: 'Typed ', style: 0 }, { text: 'heading', style: 1 }] }),
        false,
      ),
    )
  })

  it('edit-text tool: scanner text says why it cannot be edited', async () => {
    const editor = fakeEditor('edit-text')
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)
    await waitFor(() => expect(editor.textAnalysis?.paragraphs).toHaveBeenCalled())
    await act(async () => {})
    fireEvent.pointerDown(layer(), { button: 0, clientX: 150, clientY: 495 })
    expect(screen.queryByRole('textbox')).toBeNull()
    expect(editor.setMessage).toHaveBeenCalledWith(expect.stringContaining('scanned image'))
  })

  it('edit-text tool: a click on empty space adds new text there', async () => {
    const editor = fakeEditor('edit-text')
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[]} textEdits={[]} previewCurrent={false} editor={editor} />)
    await waitFor(() => expect(editor.textAnalysis?.paragraphs).toHaveBeenCalled())
    await act(async () => {})
    fireEvent.pointerDown(layer(), { button: 0, clientX: 50, clientY: 700 })
    const box = await screen.findByRole('textbox', { name: 'Paragraph text' })
    box.textContent = 'Fresh words'
    await act(async () => {
      fireEvent.keyDown(box, { key: 'Enter', ctrlKey: true })
    })
    expect(editor.editText).toHaveBeenCalledWith(
      expect.objectContaining({ paragraphId: null, insertAfterOp: -1, spans: [{ text: 'Fresh words', style: 0 }] }),
      false,
    )
  })

  it('select tool: pressing a shape selects it and dragging it moves it by the PDF-space distance', () => {
    const ink: PdfAnnotationDraft = { kind: 'ink', id: 'i', page: 1, points: [{ x: 10, y: 10 }], color: { r: 0, g: 0, b: 0 }, width: 2 }
    const editor = fakeEditor('select')
    render(<PdfEditLayer info={info()} documentRotation={0} annotations={[ink]} textEdits={[]} previewCurrent={false} editor={editor} />)
    const shape = document.querySelector('[data-edit-id="i"]') as Element

    act(() => {
      fireEvent.pointerDown(shape, { button: 0, clientX: 10, clientY: 990 })
    })
    expect(editor.setSelectedId).toHaveBeenCalledWith('i')
    fireEvent.pointerMove(shape, { clientX: 30, clientY: 980 })
    fireEvent.pointerUp(shape)
    expect(editor.updateDraft).toHaveBeenCalledWith(expect.objectContaining({ id: 'i', points: [{ x: 30, y: 20 }] }))
  })

  it('outlines the selected edit, and shapes take no pointer input outside edit mode', () => {
    const ink: PdfAnnotationDraft = { kind: 'ink', id: 'i', page: 1, points: [{ x: 10, y: 10 }], color: { r: 0, g: 0, b: 0 }, width: 2 }
    const { rerender } = render(
      <PdfEditLayer info={info()} documentRotation={0} annotations={[ink]} textEdits={[]} previewCurrent={false} editor={fakeEditor('select', { selectedId: 'i' })} />,
    )
    expect(document.querySelector('.pdf-edit-layer__selection')).not.toBeNull()

    rerender(
      <PdfEditLayer info={info()} documentRotation={0} annotations={[ink]} textEdits={[]} previewCurrent={false} editor={fakeEditor('select', { isEditing: false })} />,
    )
    expect((document.querySelector('[data-edit-id="i"]') as SVGElement).style.pointerEvents).toBe('none')
    expect(layer().style.pointerEvents).toBe('none')
  })
})

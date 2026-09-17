/**
 * USR-18 — CodeMirror 6 editor surface for source files (loaded lazily by
 * CodeViewer). Provides editing with undo/redo, line numbers, folding,
 * bracket matching/closing, auto-indent, multiple cursors (Alt+click,
 * Ctrl+D), search/replace (Ctrl+F / Ctrl+H), go to line (Ctrl+G), word wrap
 * and language highlighting loaded on demand from the file name.
 */
import { memo, useEffect, useRef } from 'react'
import { basicSetup } from 'codemirror'
import { indentWithTab } from '@codemirror/commands'
import { LanguageDescription } from '@codemirror/language'
import { languages } from '@codemirror/language-data'
import { gotoLine, openSearchPanel, selectNextOccurrence } from '@codemirror/search'
import { Compartment, EditorState, type Extension } from '@codemirror/state'
import { EditorView, keymap } from '@codemirror/view'
import { oneDark } from '@codemirror/theme-one-dark'

export type CodeEditorApi = {
  readonly getText: () => string
  readonly openSearch: () => void
  readonly goToLine: () => void
  /** 1-based line. */
  readonly revealLine: (line: number) => void
  readonly focus: () => void
}

type CodeEditorProps = {
  readonly initialText: string
  readonly fileName: string
  readonly dark: boolean
  readonly wrap: boolean
  readonly onDocChange: (lineCount: number, text: () => string) => void
  readonly onReady: (api: CodeEditorApi | null) => void
}

const baseTheme = EditorView.theme({
  '&': { height: '100%', fontSize: '13px' },
  '.cm-scroller': { fontFamily: "'JetBrains Mono', ui-monospace, 'Cascadia Code', 'Fira Code', Consolas, monospace", lineHeight: '1.6' },
})

const editorKeys = keymap.of([
  indentWithTab,
  { key: 'Mod-g', run: gotoLine, preventDefault: true },
  { key: 'Mod-d', run: selectNextOccurrence, preventDefault: true },
])

function CodeEditorBase({ initialText, fileName, dark, wrap, onDocChange, onReady }: CodeEditorProps) {
  const hostRef = useRef<HTMLDivElement | null>(null)
  const viewRef = useRef<EditorView | null>(null)
  const themeCompartment = useRef(new Compartment())
  const wrapCompartment = useRef(new Compartment())
  const languageCompartment = useRef(new Compartment())
  const callbacks = useRef({ onDocChange, onReady })

  useEffect(() => {
    callbacks.current = { onDocChange, onReady }
  }, [onDocChange, onReady])

  // Create the view once per file; later prop changes are applied through compartments.
  useEffect(() => {
    const host = hostRef.current
    if (!host) return undefined
    const extensions: Extension[] = [
      basicSetup,
      editorKeys,
      baseTheme,
      themeCompartment.current.of(dark ? oneDark : []),
      wrapCompartment.current.of(wrap ? EditorView.lineWrapping : []),
      languageCompartment.current.of([]),
      EditorView.updateListener.of((update) => {
        if (update.docChanged) {
          const { doc } = update.state
          callbacks.current.onDocChange(doc.lines, () => doc.toString())
        }
      }),
    ]
    const view = new EditorView({ state: EditorState.create({ doc: initialText, extensions }), parent: host })
    viewRef.current = view

    let cancelled = false
    const description = LanguageDescription.matchFilename(languages, fileName)
    void description?.load().then((support) => {
      if (!cancelled) view.dispatch({ effects: languageCompartment.current.reconfigure(support) })
    })

    callbacks.current.onReady({
      getText: () => view.state.doc.toString(),
      openSearch: () => {
        openSearchPanel(view)
      },
      goToLine: () => {
        view.focus()
        gotoLine(view)
      },
      revealLine: (line) => {
        const target = view.state.doc.line(Math.min(Math.max(line, 1), view.state.doc.lines))
        view.dispatch({ selection: { anchor: target.from }, effects: EditorView.scrollIntoView(target.from, { y: 'start' }) })
        view.focus()
      },
      focus: () => view.focus(),
    })

    return () => {
      cancelled = true
      callbacks.current.onReady(null)
      view.destroy()
      viewRef.current = null
    }
    // initialText/fileName identify the file; theme and wrap update in place below.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialText, fileName])

  useEffect(() => {
    viewRef.current?.dispatch({ effects: themeCompartment.current.reconfigure(dark ? oneDark : []) })
  }, [dark])

  useEffect(() => {
    viewRef.current?.dispatch({ effects: wrapCompartment.current.reconfigure(wrap ? EditorView.lineWrapping : []) })
  }, [wrap])

  return <div ref={hostRef} className="code-editor" data-testid="code-editor" />
}

export const CodeEditor = memo(CodeEditorBase)

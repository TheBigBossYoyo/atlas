import { lazy, memo, Suspense, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { ListOrdered, Play, Save, Search, Square, Trash2, WrapText, X } from 'lucide-react'

import type { NavItem, ViewerProps } from '../formats/types'
import { useTheme } from '../hooks/useTheme'
import { getLangForExt } from './extToLang'
import type { CodeEditorApi } from './code/CodeEditor'
import { useCodeRun } from './code/useCodeRun'
import {
  useRegisterViewerFind,
  useRegisterViewerSave,
  useSetNavItems,
  useSetViewerDirty,
  useSetViewerStats,
} from './shared/useViewerContext'
import { useTranslate } from '../i18n'
import { translateWriteError } from '../i18n/translateWriteError'
import { getTextFileMeta, carryTextFileMeta } from '../utils/textDecoding'
import './__styles__/viewer-code.css'

/** USR-18 — CodeMirror is a sizeable chunk; only code files pay for it. */
const LazyCodeEditor = lazy(async () => ({ default: (await import('./code/CodeEditor')).CodeEditor }))

const DARK_THEMES = new Set(['dark', 'nord', 'dracula'])
const WRAP_STORAGE_KEY = 'atlas.codeEditor.wrap'

const SCRIPT_SYMBOLS: ReadonlyArray<RegExp> = [
  /(?:export\s+)?(?:async\s+)?function\s+(\w+)/,
  /(?:export\s+)?(?:const|let)\s+(\w+)\s*=\s*(?:async\s+)?\(/,
  /class\s+(\w+)/,
]

const SYMBOL_PATTERNS: Readonly<Record<string, ReadonlyArray<RegExp>>> = {
  typescript: SCRIPT_SYMBOLS,
  tsx: SCRIPT_SYMBOLS,
  javascript: SCRIPT_SYMBOLS,
  jsx: SCRIPT_SYMBOLS,
  python: [/^(?:def|class)\s+(\w+)/],
  rust: [/^(?:pub\s+)?(?:fn|struct|enum)\s+(\w+)/],
}

function readWrapPreference(): boolean {
  try {
    return window.localStorage.getItem(WRAP_STORAGE_KEY) === '1'
  } catch {
    return false
  }
}

function extensionOf(path: string): string {
  const match = /\.([^./\\]+)$/.exec(path)
  return match ? match[1] : ''
}

function baseName(path: string): string {
  return path.replace(/\\/g, '/').split('/').pop() || path
}

function CodeViewerBase({ file }: ViewerProps) {
  const original = file.kind === 'text' ? file.content : ''
  const lang = getLangForExt(`.${extensionOf(file.path)}`) || 'text'
  const { theme } = useTheme()
  const t = useTranslate()

  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const setDirty = useSetViewerDirty()
  const registerSave = useRegisterViewerSave()
  const registerFind = useRegisterViewerFind()

  const apiRef = useRef<CodeEditorApi | null>(null)
  const [lineCount, setLineCount] = useState(() => original.split('\n').length)
  const [wrap, setWrap] = useState(readWrapPreference)
  const [savePath, setSavePath] = useState(file.path)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [symbolSource, setSymbolSource] = useState(original)
  const runner = useCodeRun(file.path)

  const handleReady = useCallback((api: CodeEditorApi | null) => {
    apiRef.current = api
  }, [])

  const handleDocChange = useCallback(
    (lines: number) => {
      setLineCount(lines)
      setDirty(true)
    },
    [setDirty],
  )

  const save = useCallback(async (): Promise<boolean> => {
    const api = apiRef.current
    if (!api) return false
    setSaveError(null)
    const text = api.getText()
    // NIGHT/text-roundtrip — reapply this file's source encoding/BOM/newline
    // convention (SHELL-1/SHELL-2's same defect applied here too: CodeMirror
    // normalizes everything to `\n` internally, and no encoding/BOM was ever
    // remembered past the initial decode). `text` itself stays `\n`-only —
    // `meta` travels alongside it and main does the re-encoding, the same
    // contract as the markdown save path in App.tsx.
    const meta = getTextFileMeta(savePath)
    const extension = extensionOf(savePath)
    const result = await window.electronAPI?.saveFile?.({
      content: text,
      suggestedName: baseName(savePath),
      filters: extension ? [{ name: 'Source file', extensions: [extension] }] : [],
      existingPath: savePath,
      meta,
    })
    if (!result?.saved) {
      setSaveError(result?.error ? translateWriteError(t, result) : t('codeViewer.saveCancelled'))
      return false
    }
    // Save As to a new file keeps the source's conventions (this fix's
    // spec) — carry `meta` to the new path before switching `savePath` over
    // to it, so the next save still finds it.
    if (result.path) {
      carryTextFileMeta(savePath, result.path)
      setSavePath(result.path)
    }
    setDirty(false)
    setSymbolSource(text)
    return true
  }, [savePath, setDirty, t])

  useEffect(() => {
    registerSave(save)
    return () => registerSave(null)
  }, [registerSave, save])

  useEffect(() => {
    registerFind(() => apiRef.current?.openSearch())
    return () => registerFind(null)
  }, [registerFind])

  useEffect(() => {
    setStats({ kind: 'code', language: lang, lines: lineCount })
  }, [setStats, lang, lineCount])

  // Outline of top-level symbols (refreshed on load and save, not per keystroke).
  const navItems = useMemo<NavItem[]>(() => {
    const patterns = SYMBOL_PATTERNS[lang] ?? []
    if (patterns.length === 0) return []
    return symbolSource.split('\n').flatMap((line, index) => {
      const name = patterns.map((pattern) => pattern.exec(line)?.[1]).find(Boolean)
      return name ? [{ id: `line-${index}`, label: name, onSelect: () => apiRef.current?.revealLine(index + 1) }] : []
    })
  }, [lang, symbolSource])

  useEffect(() => {
    setNavItems(navItems)
  }, [setNavItems, navItems])

  const toggleWrap = (): void => {
    setWrap((current) => {
      try {
        window.localStorage.setItem(WRAP_STORAGE_KEY, current ? '0' : '1')
      } catch {
        // Preference only.
      }
      return !current
    })
  }

  return (
    <div className="code-viewer">
      <div className="code-viewer__toolbar" role="toolbar" aria-label={t('codeViewer.toolbarAria')}>
        <span className="code-viewer__language">{lang}</span>
        <div className="code-viewer__actions">
          {runner.isAvailable && (
            <button
              type="button"
              className="code-viewer__button"
              aria-label={runner.state.status === 'running' ? t('codeViewer.stop') : t('codeViewer.run')}
              title={runner.state.status === 'running' ? t('codeViewer.stopTitle') : t('codeViewer.runTitle')}
              onClick={() => {
                if (runner.state.status === 'running') {
                  runner.stop()
                  return
                }
                // The file on disk is what runs, so an edited file is saved first.
                void save().then((saved) => (saved ? runner.run() : undefined))
              }}
            >
              {runner.state.status === 'running' ? <Square size={15} /> : <Play size={15} />}
              {runner.state.status === 'running' ? t('codeViewer.stop') : t('codeViewer.run')}
            </button>
          )}
          <button type="button" className="code-viewer__button" aria-label={t('codeViewer.findReplaceAria')} title={t('codeViewer.findReplaceTitle')} onClick={() => apiRef.current?.openSearch()}>
            <Search size={15} />
          </button>
          <button type="button" className="code-viewer__button" aria-label={t('codeViewer.goToLineAria')} title={t('codeViewer.goToLineTitle')} onClick={() => apiRef.current?.goToLine()}>
            <ListOrdered size={15} />
          </button>
          <button type="button" className="code-viewer__button" aria-label={t('codeViewer.wordWrapAria')} aria-pressed={wrap} title={t('codeViewer.wordWrapAria')} onClick={toggleWrap}>
            <WrapText size={15} />
          </button>
          <button type="button" className="code-viewer__button code-viewer__button--primary" aria-label={t('codeViewer.saveAria')} title={t('codeViewer.saveTitle')} onClick={() => void save()}>
            <Save size={15} />
            {t('codeViewer.saveAria')}
          </button>
        </div>
      </div>
      {saveError && (
        <div className="code-viewer__error" role="alert">
          {saveError}
        </div>
      )}
      <div className="code-viewer__body">
        <Suspense fallback={<pre className="code-viewer__pre"><code>{original}</code></pre>}>
          <LazyCodeEditor
            initialText={original}
            fileName={baseName(file.path)}
            dark={DARK_THEMES.has(theme)}
            wrap={wrap}
            onDocChange={handleDocChange}
            onReady={handleReady}
          />
        </Suspense>
      </div>
      {runner.isOpen && (
        <div className="code-viewer__output" role="region" aria-label={t('codeViewer.outputRegionAria')}>
          <div className="code-viewer__output-header">
            <span className="code-viewer__output-status" role="status">
              {runner.state.status === 'running'
                ? t('codeViewer.running')
                : runner.state.status === 'finished'
                  ? runner.state.summary
                  : t('codeViewer.output')}
            </span>
            <div className="code-viewer__actions">
              <button type="button" className="code-viewer__button" aria-label={t('codeViewer.clearOutputAria')} title={t('codeViewer.clearOutputAria')} onClick={runner.clear}>
                <Trash2 size={14} />
              </button>
              <button type="button" className="code-viewer__button" aria-label={t('codeViewer.closeOutputAria')} title={t('codeViewer.closeOutputAria')} onClick={runner.close}>
                <X size={14} />
              </button>
            </div>
          </div>
          <pre className="code-viewer__output-body">
            {runner.chunks.map((chunk, index) => (
              <span key={index} className={`code-viewer__output-chunk code-viewer__output-chunk--${chunk.stream}`}>
                {chunk.text}
              </span>
            ))}
          </pre>
        </div>
      )}
    </div>
  )
}

export const CodeViewer = memo(CodeViewerBase)

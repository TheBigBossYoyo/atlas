import { memo, useMemo, useState } from 'react'
import { FileQuestion, FileWarning, FolderOpen, Type } from 'lucide-react'

import type { LoadedFile } from '../formats/types'
import { isLegacyOfficeMagic, guessLegacyOfficeKind, legacyOfficeMessage } from '../formats/legacyOffice'
import { decodeTextBuffer } from '../utils/textDecoding'
import { VirtualizedPlainText } from './shared/VirtualizedPlainText'
import './__styles__/viewer-unknown.css'

type UnknownViewerProps = {
  readonly file: LoadedFile
}

/**
 * wave-3 shell-polish follow-up — "Open as text" is a best-effort fallback
 * for a file Atlas couldn't otherwise identify; it may not be text at all
 * (a heuristic guess gone wrong) and could be arbitrarily large. Capping how
 * much of it is ever decoded/rendered bounds both the string-decode cost and
 * the DOM this view can produce, independent of `VirtualizedPlainText`'s own
 * row virtualization (which only bounds *mounted* rows, not the decoded
 * string itself).
 */
const TEXT_PREVIEW_MAX_BYTES = 5 * 1024 * 1024 // 5 MiB

function fileNameOf(filePath: string): string {
  return filePath.split(/[\\/]/).pop() ?? filePath
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  const units = ['KB', 'MB', 'GB']
  let value = bytes / 1024
  let unitIndex = 0
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024
    unitIndex += 1
  }
  return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[unitIndex]}`
}

/**
 * P2.11/LOAD-18 — was `<span>{format}</span>` + the raw path, with zero
 * actionable UX. Now a real empty state: icon, file name/size, an
 * explanation of why Atlas can't render it, and two actions ("open as
 * text" / "reveal in folder"). Also gives legacy binary Office documents
 * (`.doc`/`.xls`/`.ppt`, LOAD-11) their own honest message instead of the
 * generic "unknown format" copy.
 */
function UnknownViewerBase({ file }: UnknownViewerProps) {
  const [showAsText, setShowAsText] = useState(false)
  const [revealError, setRevealError] = useState<string | null>(null)

  const buffer = file.kind === 'binary' ? file.content : null
  const fileName = fileNameOf(file.path)
  const sizeLabel = buffer ? formatFileSize(buffer.byteLength) : null

  const legacyKind = useMemo(() => {
    if (!buffer || !isLegacyOfficeMagic(buffer)) {
      return null
    }
    return guessLegacyOfficeKind(file.path)
  }, [buffer, file.path])

  const isTruncated = Boolean(buffer && buffer.byteLength > TEXT_PREVIEW_MAX_BYTES)

  const decodedLines = useMemo(() => {
    if (!showAsText || !buffer) {
      return []
    }
    const preview = buffer.byteLength > TEXT_PREVIEW_MAX_BYTES ? buffer.slice(0, TEXT_PREVIEW_MAX_BYTES) : buffer
    return decodeTextBuffer(preview).split('\n')
  }, [showAsText, buffer])

  const handleRevealInFolder = async () => {
    setRevealError(null)
    if (!window.electronAPI) {
      setRevealError('Reveal in folder requires the Atlas desktop app.')
      return
    }
    const result = await window.electronAPI.revealInFolder(file.path)
    if (!result.ok) {
      setRevealError('Could not reveal this file — it may have moved or been deleted.')
    }
  }

  if (showAsText) {
    return (
      <div className="unknown-viewer unknown-viewer--as-text">
        <div className="unknown-viewer__text-toolbar">
          <button type="button" onClick={() => setShowAsText(false)}>
            Back
          </button>
          <span className="unknown-viewer__text-filename">{fileName}</span>
          {isTruncated && (
            <span className="unknown-viewer__text-truncated">
              Showing the first {formatFileSize(TEXT_PREVIEW_MAX_BYTES)} only
            </span>
          )}
        </div>
        <VirtualizedPlainText
          lines={decodedLines}
          className="unknown-viewer__text-content"
          lineClassName="unknown-viewer__text-line"
        />
      </div>
    )
  }

  return (
    // UX-17 — `.unknown-viewer` is kept (not replaced) — it owns this
    // viewer's own required scroll-container behavior (DAT-02/UX-04,
    // viewerScrollContainers test) and its full empty-state content
    // (P2.11/LOAD-18: icon, filename/size, explanation, actions); the shared
    // `.viewer-fallback` class layers the app-wide crash/loading/unknown
    // theming (padding/background/color tokens) on top of it.
    <div className="unknown-viewer viewer-fallback">
      {legacyKind !== null ? (
        <FileWarning size={48} strokeWidth={1.5} className="viewer-fallback__icon" aria-hidden="true" />
      ) : (
        <FileQuestion size={48} strokeWidth={1.5} className="viewer-fallback__icon" aria-hidden="true" />
      )}

      <h2 className="unknown-viewer__filename viewer-fallback__title">{fileName}</h2>
      {sizeLabel && <p className="unknown-viewer__size">{sizeLabel}</p>}

      <p className="unknown-viewer__explanation">
        {legacyKind !== null
          ? legacyOfficeMessage(legacyKind)
          : "Atlas doesn't recognize this file's format, so there's nothing to preview here."}
      </p>

      <div className="unknown-viewer__actions">
        <button type="button" onClick={() => setShowAsText(true)}>
          <Type size={16} aria-hidden="true" />
          Open as text
        </button>
        <button type="button" onClick={() => void handleRevealInFolder()}>
          <FolderOpen size={16} aria-hidden="true" />
          Reveal in folder
        </button>
      </div>

      {revealError && (
        <p className="unknown-viewer__error" role="alert">
          {revealError}
        </p>
      )}

      <p className="unknown-viewer__path" title={file.path}>
        {file.path}
      </p>
    </div>
  )
}

export const UnknownViewer = memo(UnknownViewerBase)

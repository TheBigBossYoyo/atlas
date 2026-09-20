import { memo, useMemo, useState } from 'react'
import { FileQuestion, FileWarning, FolderOpen, Type } from 'lucide-react'

import type { LoadedFile } from '../formats/types'
import { isLegacyOfficeMagic, guessLegacyOfficeKind, legacyOfficeMessage } from '../formats/legacyOffice'
import { decodeTextBuffer } from '../utils/textDecoding'
import { useTranslate } from '../i18n'
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

/**
 * Review fix (wave-3 shell-polish): a byte-boundary cap has to trim off a
 * trailing UTF-8 sequence that got cut in half, or `decodeTextBuffer`'s own
 * `isValidUtf8` check — which rejects a buffer outright the moment any
 * multi-byte sequence runs off the end, not just its own last character —
 * fails for the *entire* preview and silently falls back to decoding it all
 * as Windows-1252 instead, turning every non-ASCII character in an
 * otherwise-legitimate preview into mojibake well before the truncation
 * point, not just the one character actually cut off. Scans back at most 4
 * bytes (UTF-8's longest sequence) for the start of a sequence left without
 * enough continuation bytes and drops it; a no-op for content that happens
 * to end cleanly (pure ASCII, a complete multi-byte character, or genuinely
 * invalid UTF-8 that `isValidUtf8` was always going to reject regardless).
 */
function trimIncompleteUtf8Tail(bytes: Uint8Array): Uint8Array {
  const len = bytes.length
  const maxBack = Math.min(4, len)
  for (let back = 1; back <= maxBack; back += 1) {
    const byte = bytes[len - back]
    if ((byte & 0xc0) === 0x80) continue // continuation byte — keep looking back for its lead

    let expectedLength: number
    if (byte <= 0x7f) expectedLength = 1
    else if ((byte & 0xe0) === 0xc0) expectedLength = 2
    else if ((byte & 0xf0) === 0xe0) expectedLength = 3
    else if ((byte & 0xf8) === 0xf0) expectedLength = 4
    else expectedLength = 1 // not a valid lead byte either way — leave it for isValidUtf8 to reject

    return expectedLength > back ? bytes.subarray(0, len - back) : bytes
  }
  return bytes
}

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
  const t = useTranslate()
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
    if (buffer.byteLength <= TEXT_PREVIEW_MAX_BYTES) {
      return decodeTextBuffer(buffer).split('\n')
    }
    const capped = buffer.slice(0, TEXT_PREVIEW_MAX_BYTES)
    const trimmed = trimIncompleteUtf8Tail(new Uint8Array(capped))
    // `.slice()` (not `.subarray()`) so the ArrayBuffer handed to
    // `decodeTextBuffer` is exactly `trimmed`'s bytes, not `capped`'s full
    // (untrimmed) backing buffer at a shorter view length.
    return decodeTextBuffer(trimmed.slice().buffer).split('\n')
  }, [showAsText, buffer])

  const handleRevealInFolder = async () => {
    setRevealError(null)
    if (!window.electronAPI) {
      setRevealError(t('unknown.viewer.revealRequiresDesktop'))
      return
    }
    const result = await window.electronAPI.revealInFolder(file.path)
    if (!result.ok) {
      setRevealError(t('unknown.viewer.revealFailed'))
    }
  }

  if (showAsText) {
    return (
      <div className="unknown-viewer unknown-viewer--as-text">
        <div className="unknown-viewer__text-toolbar">
          <button type="button" onClick={() => setShowAsText(false)}>
            {t('unknown.viewer.back')}
          </button>
          <span className="unknown-viewer__text-filename">{fileName}</span>
          {isTruncated && (
            <span className="unknown-viewer__text-truncated">
              {t('unknown.viewer.truncatedNotice', { size: formatFileSize(TEXT_PREVIEW_MAX_BYTES) })}
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
          : t('unknown.viewer.genericExplanation')}
      </p>

      <div className="unknown-viewer__actions">
        <button type="button" onClick={() => setShowAsText(true)}>
          <Type size={16} aria-hidden="true" />
          {t('unknown.viewer.openAsText')}
        </button>
        <button type="button" onClick={() => void handleRevealInFolder()}>
          <FolderOpen size={16} aria-hidden="true" />
          {t('unknown.viewer.revealInFolder')}
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

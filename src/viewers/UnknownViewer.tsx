import { memo } from 'react'
import { FileQuestion } from 'lucide-react'

import type { LoadedFile } from '../formats/types'
import './__styles__/viewer-unknown.css'

type UnknownViewerProps = {
  readonly file: LoadedFile
}

function UnknownViewerBase({ file }: UnknownViewerProps) {
  return (
    // UX-17 — themed, centered fallback (was two bare, unstyled <span>s).
    // `.unknown-viewer` is kept (not replaced) — it owns this viewer's own
    // required scroll-container behavior (DAT-02/UX-04, viewerScrollContainers
    // test) — `.viewer-fallback` layers the shared crash/loading/unknown look
    // on top of it.
    <div className="unknown-viewer viewer-fallback">
      <FileQuestion size={32} className="viewer-fallback__icon" aria-hidden="true" />
      <h2 className="viewer-fallback__title">Can&rsquo;t preview this file</h2>
      <p className="viewer-fallback__detail">
        {file.format} — {file.path}
      </p>
    </div>
  )
}

export const UnknownViewer = memo(UnknownViewerBase)

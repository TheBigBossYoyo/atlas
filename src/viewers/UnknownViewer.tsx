import { memo } from 'react'

import type { LoadedFile } from '../formats/types'
import './__styles__/viewer-unknown.css'

type UnknownViewerProps = {
  readonly file: LoadedFile
}

function UnknownViewerBase({ file }: UnknownViewerProps) {
  return (
    <div className="unknown-viewer">
      <span>{file.format}</span>
      <span>{file.path}</span>
    </div>
  )
}

export const UnknownViewer = memo(UnknownViewerBase)

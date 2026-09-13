import React from 'react'
import type { FormatId } from '../formats/types'

export type ViewerLoadingProps = {
  readonly format?: FormatId
}

export function ViewerLoading({ format }: ViewerLoadingProps): React.ReactElement {
  return (
    <div className="viewer-loading" role="status" aria-live="polite">
      Loading {format ?? 'document'}…
    </div>
  )
}

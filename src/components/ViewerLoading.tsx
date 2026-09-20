import React from 'react'
import { Loader2 } from 'lucide-react'
import type { FormatId } from '../formats/types'
import { useTranslate } from '../i18n'

export type ViewerLoadingProps = {
  readonly format?: FormatId
}

export function ViewerLoading({ format }: ViewerLoadingProps): React.ReactElement {
  const t = useTranslate()
  return (
    // UX-17 — themed, centered fallback (was plain unstyled text).
    <div className="viewer-fallback viewer-loading" role="status" aria-live="polite">
      <Loader2 size={28} className="viewer-fallback__icon viewer-loading__spinner" aria-hidden="true" />
      <p className="viewer-fallback__title">
        {t('viewerLoading.loading', { format: format ?? t('viewerLoading.documentFallback') })}
      </p>
    </div>
  )
}

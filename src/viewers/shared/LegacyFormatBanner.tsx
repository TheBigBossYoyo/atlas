import { History } from 'lucide-react'

import { useTranslate } from '../../i18n'
import '../__styles__/viewer-legacy-banner.css'

export type LegacyFormatBannerProps = {
  /** e.g. "Word 97-2003 document (.doc)". */
  readonly formatLabel: string
  /** e.g. ".docx" — the modern format to re-save as for full fidelity. */
  readonly modernExtension: string
}

/**
 * Shown atop `LegacyDocViewer`/`LegacyPptViewer` (wave-4 legacy-office) —
 * both are read-only, text-only best-effort extractions from a binary
 * format Atlas doesn't fully parse (no styling, images, or exact layout
 * survives the conversion), so this says so up front rather than letting a
 * reader assume what's shown is the whole document.
 */
export function LegacyFormatBanner({ formatLabel, modernExtension }: LegacyFormatBannerProps) {
  const t = useTranslate()
  return (
    <div className="legacy-format-banner" role="note">
      <History size={16} aria-hidden="true" className="legacy-format-banner__icon" />
      <span>
        {t('legacyFormat.notice', { formatLabel, modernExtension })}
      </span>
    </div>
  )
}

import type { TranslateFn } from '../i18n'

/**
 * A timestamp as "just now" / "5 minutes ago" / "3 days ago", falling back to a
 * plain locale date once it is a week old.
 *
 * Extracted from `WelcomeScreen`, which had it privately, when the version
 * history panel needed the same thing — two copies of a time formatter drift, and
 * the recent-files list and the version list sitting side by side with different
 * wording for the same age would look like a bug.
 */
export function formatRelativeTime(timestamp: number, t: TranslateFn): string {
  const diff = Date.now() - timestamp
  const minutes = Math.floor(diff / 60000)
  if (minutes < 1) return t('welcome.relative.justNow')
  if (minutes < 60) return t('welcome.relative.minutesAgo', { count: minutes })
  const hours = Math.floor(minutes / 60)
  if (hours < 24) return t('welcome.relative.hoursAgo', { count: hours })
  const days = Math.floor(hours / 24)
  if (days < 7) return t('welcome.relative.daysAgo', { count: days })
  return new Date(timestamp).toLocaleDateString()
}

/** A byte count as the version list shows it: "12 KB", "1.4 MB". */
export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

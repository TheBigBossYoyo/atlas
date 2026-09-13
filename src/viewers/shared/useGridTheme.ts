import { useEffect, useState } from 'react'

/**
 * glide-data-grid renders cells to a `<canvas>` and reads colors via JavaScript,
 * so CSS `var(--…)` strings are NOT resolved by the canvas APIs and end up
 * painting as black/transparent. This hook reads computed CSS variable values
 * from `document.documentElement` and returns a literal-color theme object that
 * glide-data-grid can consume. It re-resolves whenever the active app theme
 * changes (the `data-theme` attribute on `<html>`).
 */

export type CustomGridTheme = {
  bgCell: string
  bgCellMedium: string
  textDark: string
  textLight: string
  textMedium: string
  borderColor: string
  accentColor: string
  accentLight: string
  bgHeader: string
  textHeader: string
  bgHeaderHasFocus: string
  bgHeaderHovered: string
  fontFamily: string
  baseFontStyle: string
  headerFontStyle: string
  // Custom extensions
  bgRowOdd: string
  bgRowHover: string
  cellHorizontalPadding: number
  cellVerticalPadding: number
}

function readVar(styles: CSSStyleDeclaration, name: string, fallback: string): string {
  const value = styles.getPropertyValue(name).trim()
  return value.length > 0 ? value : fallback
}

function resolveGridTheme(): CustomGridTheme {
  const styles = getComputedStyle(document.documentElement)
  const bgPrimary = readVar(styles, '--bg-primary', '#ffffff')
  const bgSecondary = readVar(styles, '--bg-secondary', '#f6f8fa')
  const bgTertiary = readVar(styles, '--bg-tertiary', '#eaeef2')
  const textPrimary = readVar(styles, '--text-primary', '#1f2328')
  const textSecondary = readVar(styles, '--text-secondary', '#656d76')
  const borderPrimary = readVar(styles, '--border-primary', '#d0d7de')
  const accent = readVar(styles, '--accent', '#0969da')
  const accentSubtle = readVar(styles, '--accent-subtle', 'rgba(9, 105, 218, 0.08)')
  
  // Use body font-family to match app
  const bodyFont = getComputedStyle(document.body).fontFamily || '-apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif'

  return {
    bgCell: bgPrimary,
    bgCellMedium: bgSecondary,
    textDark: textPrimary,
    textLight: textSecondary,
    textMedium: textSecondary,
    borderColor: borderPrimary,
    accentColor: accent,
    accentLight: 'transparent',
    bgHeader: bgTertiary,
    textHeader: textPrimary,
    bgHeaderHasFocus: borderPrimary,
    bgHeaderHovered: borderPrimary,
    fontFamily: bodyFont,
    baseFontStyle: '13px',
    headerFontStyle: 'bold 13px',
    bgRowOdd: bgSecondary,
    bgRowHover: accentSubtle,
    cellHorizontalPadding: 12,
    cellVerticalPadding: 8,
  }
}

export function useGridTheme(): CustomGridTheme {
  const [theme, setTheme] = useState<CustomGridTheme>(() => resolveGridTheme())

  useEffect(() => {
    const update = () => {
      setTheme((prev) => {
        const next = resolveGridTheme()
        return           prev.bgCell === next.bgCell &&
          prev.textDark === next.textDark &&
          prev.borderColor === next.borderColor &&
          prev.accentColor === next.accentColor &&
          prev.bgHeader === next.bgHeader &&
          prev.fontFamily === next.fontFamily &&
          prev.bgRowHover === next.bgRowHover
          ? prev
          : next
      })
    }

    const observer = new MutationObserver(update)

    observer.observe(document.documentElement, {
      attributes: true,
      attributeFilter: ['data-theme', 'class'],
    })

    return () => {
      observer.disconnect()
    }
  }, [])

  return theme
}

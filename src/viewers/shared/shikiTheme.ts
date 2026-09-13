import type { Theme as ThemeId } from '../../types'
import { assertNever } from '../../formats/types'

// TODO Replace with `import type { BundledTheme } from 'shiki'` once W2 viewer installs the dep.
export type BundledTheme = 'github-light' | 'github-dark' | 'solarized-light' | 'nord' | 'dracula'

export function getShikiThemeForAppTheme(themeId: ThemeId): BundledTheme {
  switch (themeId) {
    case 'light':
      return 'github-light'
    case 'dark':
      return 'github-dark'
    case 'sepia':
      return 'solarized-light'
    case 'nord':
      return 'nord'
    case 'dracula':
      return 'dracula'
    default:
      return assertNever(themeId)
  }
}

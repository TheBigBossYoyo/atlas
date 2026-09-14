/** CSS-mapping helpers shared by `SlideCanvas` for rendering the parsed `SlideShape[]` model. */

import type { CSSProperties } from 'react'
import type { SlideBullet, SlideFill, SlideGeometry, SlideTransform } from './SlideDeck.types'

export function fillToCss(fill: SlideFill | undefined): string | undefined {
  if (!fill) {
    return undefined
  }

  switch (fill.kind) {
    case 'none':
      return 'transparent'
    case 'solid':
      return fill.color
    case 'gradient':
      return `linear-gradient(${fill.angleDeg ?? 0}deg, ${fill.colors.join(', ')})`
    default:
      return undefined
  }
}

const GEOMETRY_CLIP_PATH: Partial<Record<SlideGeometry, string>> = {
  triangle: 'polygon(50% 0%, 0% 100%, 100% 100%)',
  rightArrow: 'polygon(0% 25%, 60% 25%, 60% 0%, 100% 50%, 60% 100%, 60% 75%, 0% 75%)',
  leftArrow: 'polygon(100% 25%, 40% 25%, 40% 0%, 0% 50%, 40% 100%, 40% 75%, 100% 75%)',
  upArrow: 'polygon(25% 100%, 25% 40%, 0% 40%, 50% 0%, 100% 40%, 75% 40%, 75% 100%)',
  downArrow: 'polygon(25% 0%, 25% 60%, 0% 60%, 50% 100%, 100% 60%, 75% 60%, 75% 0%)',
}

/** Maps `SlideGeometry` -> CSS `border-radius`/`clip-path` for a shape's rendered outline. */
export function geometryToCss(geometry: SlideGeometry | undefined): CSSProperties {
  if (geometry === 'roundRect') {
    return { borderRadius: '10%' }
  }

  if (geometry === 'ellipse') {
    return { borderRadius: '50%' }
  }

  const clipPath = geometry ? GEOMETRY_CLIP_PATH[geometry] : undefined
  return clipPath ? { clipPath } : {}
}

/** `rotate()`/`scale()` for a shape's own rotation and horizontal/vertical flip (S10). */
export function transformToCss(transform: SlideTransform): string | undefined {
  const parts: string[] = []
  if (transform.rotationDeg) {
    parts.push(`rotate(${transform.rotationDeg}deg)`)
  }

  if (transform.flipH || transform.flipV) {
    parts.push(`scale(${transform.flipH ? -1 : 1}, ${transform.flipV ? -1 : 1})`)
  }

  return parts.length > 0 ? parts.join(' ') : undefined
}

/** Renders a bullet's marker text, advancing/resetting the per-level numbering counters in place (S4). */
export function bulletLabel(bullet: SlideBullet | undefined, counters: number[]): string | null {
  if (!bullet) {
    counters.length = 0
    return null
  }

  if (bullet.numbered) {
    counters.length = bullet.level + 1
    counters[bullet.level] = (counters[bullet.level] ?? 0) + 1
    return `${counters[bullet.level]}.`
  }

  counters.length = 0
  return bullet.char ?? '•'
}

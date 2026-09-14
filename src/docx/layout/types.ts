import type { FontMetrics } from '../fonts'
import type { Drawing, Paragraph, ParaProps, Run, RunProps } from '../model'
import type { Theme } from '../parser/theme'

export type EffectiveRunProps = Readonly<RunProps> & { readonly _revision?: 'ins' | 'del' }
export type EffectiveParaProps = Readonly<ParaProps>

export type LineItem =
  | {
      kind: 'word'
      text: string
      width: number
      runIndex: number
      charStart: number
      charEnd: number
      runProps: EffectiveRunProps
    }
  | {
      kind: 'space'
      width: number
      stretchable: boolean
      runIndex: number
      charOffset: number
    }
  | {
      kind: 'tab'
      width: number
      runIndex: number
      charOffset: number
    }
  | {
      kind: 'break'
      breakKind: 'line' | 'page' | 'column'
      runIndex: number
    }
  | {
      kind: 'hyphen-opportunity'
      text: ''
      penaltyWidth: number
      runIndex: number
      charOffset: number
    }
  | {
      kind: 'glyph-cluster'
      text: string
      width: number
      runIndex: number
      charStart: number
      charEnd: number
      runProps: EffectiveRunProps
    }
  | {
      // An embedded picture. Occupies one character offset in its run (the
      // editor's position model) and sits on the text baseline.
      kind: 'drawing'
      drawing: Drawing
      width: number
      height: number
      runIndex: number
      charStart: number
      charEnd: number
      runProps: EffectiveRunProps
    }

export type LineBox = {
  items: ReadonlyArray<LineItem>
  width: number
  ascent: number
  descent: number
  lineHeight: number
  isJustified: boolean
  justificationStretch: number
  /**
   * Extra space reserved above the text strut so drawings taller than the
   * text ascent fit inside the line. Already included in `lineHeight`; the
   * text strut occupies the remaining `lineHeight - drawingClearancePt`.
   * Omitted when no drawing on the line needs it.
   */
  drawingClearancePt?: number
  /**
   * True when this line's last item is a manual page break (Ctrl+Enter) —
   * i.e. a `break` item with `breakKind: 'page'`. `paginate.ts` (D10) uses
   * this to force a new page immediately after this line, regardless of
   * how much room remains, instead of treating it as an ordinary line
   * break. Mutually exclusive with `endsWithColumnBreak`. Omitted (not
   * `false`) when not set, matching `drawingClearancePt`'s convention.
   */
  endsWithPageBreak?: boolean
  /**
   * Same as `endsWithPageBreak`, but for a manual column break
   * (Ctrl+Shift+Enter): advances to the next column on the current page
   * (or opens a new page when already on the last column) instead of an
   * ordinary line break.
   */
  endsWithColumnBreak?: boolean
}

export type FontResolver = (
  family: string,
  variant: 'regular' | 'bold' | 'italic' | 'boldItalic',
) => Promise<FontMetrics>

export type TabStop = {
  positionPt: number
  alignment: 'left' | 'center' | 'right' | 'decimal'
  leader: 'none' | 'dot' | 'hyphen' | 'underscore'
}

export type LineBreakInput = {
  paragraph: Paragraph
  paraProps: EffectiveParaProps
  runs: ReadonlyArray<{
    run: Run
    runProps: EffectiveRunProps
  }>
  availableWidth: number
  fontResolver: FontResolver
  tabStops: ReadonlyArray<TabStop>
  theme?: Theme
  /**
   * Already-itemized items prepended before `runs`' own content — used by
   * D3 to give a list marker (and its trailing tab/space) real measured
   * width so it participates in line-breaking and hanging-indent alignment
   * instead of being painted separately outside layout. Being first in
   * document order, they naturally land on line 0. Their `runIndex` is a
   * caller-chosen sentinel (never a real index into `runs`) so they never
   * collide with, or shift, real content's run indices.
   */
  leadingItems?: ReadonlyArray<LineItem>
}

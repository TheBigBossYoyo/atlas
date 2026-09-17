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
      /**
       * Present only for a synthesized footnote/endnote reference mark
       * (D11/DXL-09) — lets `paginate.ts`'s placement pass recognize which
       * placed words are note markers (to track "which footnotes are
       * referenced on this page") without re-deriving that from raw text.
       * Absent for every ordinary word, including a literal digit typed by
       * the document author.
       */
      noteRef?: { readonly kind: 'footnote' | 'endnote'; readonly id: string }
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
      /**
       * The matched tab stop's leader glyph (D24/DXL-16), resolved once the
       * tab's width is known (see `breakLines.ts`'s `resolveTabItem`).
       * `'none'` (the default — no visible fill) when the item hasn't been
       * through tab-stop resolution yet, or the paragraph has no tab stops
       * at all (falls back to Word's default un-leadered tab stops).
       */
      leader: 'none' | 'dot' | 'hyphen' | 'underscore'
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
  /**
   * DEFER-2 — the paragraph reads right-to-left (`w:bidi`). Line breaking
   * stays in logical order; the renderer sets `dir="rtl"` on the line so the
   * browser applies the Unicode bidirectional algorithm within it.
   */
  rtl?: boolean
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
  /**
   * Precomputed horizontal offset (indent + center/right alignment) for
   * content that renders its lines directly rather than going through
   * `paginate.ts`'s body-content placement pass (`placeLineSlice` /
   * `computeLineLeftOffsetPt`, whose result instead lives on the placed
   * `PageLineRef.leftPt`) — currently header/footer content
   * (`buildBlockGroupLines`) and footnote body content
   * (`buildFootnoteContentLines`); endnotes flow through the ordinary body
   * pipeline and so never set this. Omitted (treat as 0, flush with the
   * container's own left edge) for any line that doesn't need it.
   */
  leftOffsetPt?: number
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
  /**
   * Resolved display marks for footnote/endnote references appearing in
   * this paragraph (D11 milestone 2), keyed by note id. `itemizeRuns` looks
   * up a `footnote-reference`/`endnote-reference` run child's id here to
   * render its superscripted marker; an id with no entry (nested note
   * references inside a footnote/endnote body are not resolved — see
   * `noteNumbering.ts`'s module doc) renders as an empty marker.
   */
  noteMarks?: NoteMarks
}

export type NoteMarks = {
  readonly footnote: ReadonlyMap<string, string>
  readonly endnote: ReadonlyMap<string, string>
}

import type {
  RunProps,
  ParaProps,
  Border,
  Color,
} from '../model/styles';
import type { TableProps, TableCellProps } from '../model/document';
import { resolveFontFamily } from '../fonts/families';
import { effectiveFontFamily } from '../layout/fontResolution';
import type { Theme } from '../parser/theme';

const twipToPt = (twip: number) => twip / 20;
const halfPointToPt = (hp: number) => hp / 2;
const eighthPointToPt = (ep: number) => ep / 8;

// MUST stay in sync with src/docx/layout/itemize.ts and breakLines.ts.
// When a run has no <w:sz>, the paginator measures it at 11pt — so the DOM
// MUST paint it at 11pt too, otherwise Chromium falls back to the 16px UI
// default and overshoots every reserved word/line box (3.0.5 defect).
const DEFAULT_FONT_SIZE_PT = 11;
const revisionAuthorPalette = ['#1d4ed8', '#9333ea', '#0f766e', '#b45309', '#be123c', '#475569'];

export type RevisionRenderKind = 'ins' | 'del';

function quoteFamily(name: string): string {
  const trimmed = name.trim().replace(/['"]/g, '');
  return /\s/.test(trimmed) ? `"${trimmed}"` : trimmed;
}

/**
 * Build a CSS font-family stack from a Word run font name.
 *
 * Why: layout measures glyph widths from the bundled metric-compatible
 * substitute (Carlito/Caladea/Tinos/Arimo/CourierPrime). The DOM MUST render
 * with that same font for advance widths to match the layout. We:
 *   1. Emit the requested Word name (registered via @font-face → substitute).
 *   2. Append the substitute name explicitly as a backup.
 *   3. Append a generic family so unknown fonts never collapse to nothing.
 */
function buildFontFamilyStack(wordName: string): string {
  const stack: string[] = [quoteFamily(wordName)];
  const resolved = resolveFontFamily(wordName);
  if (resolved && resolved.substituteName.toLowerCase() !== wordName.trim().toLowerCase()) {
    stack.push(quoteFamily(resolved.substituteName));
  }
  // Generic fallback: monospace for Courier-like, serif for Times/Cambria-like, else sans-serif.
  const lower = wordName.toLowerCase();
  const generic = lower.includes('courier') || lower.includes('mono')
    ? 'monospace'
    : lower.includes('times') || lower.includes('cambria') || lower.includes('serif')
      ? 'serif'
      : 'sans-serif';
  stack.push(generic);
  return stack.join(', ');
}

function resolveColor(color: Color | undefined): string | undefined {
  if (!color || color === 'auto') return undefined;
  if (color.startsWith('#')) return color;
  return `#${color}`;
}

function hashAuthor(author: string): number {
  let hash = 0;
  for (let index = 0; index < author.length; index += 1) {
    hash = (hash * 31 + author.charCodeAt(index)) >>> 0;
  }
  return hash;
}

export function revisionStyleToCss(
  kind: RevisionRenderKind,
  author?: string,
): React.CSSProperties {
  const normalizedAuthor = author?.trim();
  const color =
    normalizedAuthor && normalizedAuthor.length > 0
      ? revisionAuthorPalette[hashAuthor(normalizedAuthor) % revisionAuthorPalette.length]
      : kind === 'ins'
        ? '#2e7d32'
        : '#c62828';

  return {
    color,
    textDecoration: kind === 'ins' ? 'underline' : 'line-through',
    textDecorationColor: color,
  };
}

export function borderToCss(border: Border | undefined): string | undefined {
  if (!border || border.style === 'none' || border.style === 'nil') return undefined;
  // Fallbacks: default to 1pt if missing size, black if missing color
  const size = border.size !== undefined ? eighthPointToPt(border.size) : 1;
  const color = resolveColor(border.color) ?? 'var(--border-primary, #d0d7de)';
  const style = border.style === 'dashed' ? 'dashed'
    : border.style === 'dotted' ? 'dotted'
    : border.style === 'double' ? 'double'
    : 'solid'; // Treat all other exotic OOXML border styles as solid for now

  return `${size}pt ${style} ${color}`;
}

/**
 * Length unit for text-level CSS lengths (font-size, letter-spacing).
 *
 * - `'pt'`: normal CSS pt (default). Used by the legacy flow renderer
 *   (`run.tsx` / `paragraph.tsx`) where the DOM lives in real CSS coordinates.
 * - `'layoutPx'`: emit the numeric pt value as `px`. Used ONLY by `PageView`,
 *   whose inner sheet treats numeric layout-pt values as internal "px" units
 *   inside a `transform: scale(4/3)` wrapper. Emitting real `pt` there would
 *   double-scale the text (browser pt→px ×4/3, then transform ×4/3 again) and
 *   make rendered glyphs ~33% wider than the paginator predicted — that is the
 *   3.0.3 right-edge clipping + vertical line overlap.
 *
 * NEVER use `'layoutPx'` outside the fixed-layout PageView surface.
 */
export type LengthUnit = 'pt' | 'layoutPx';

export type RunStyleOptions = {
  readonly lengthUnit?: LengthUnit;
};

function emitLength(valuePt: number, unit: LengthUnit): string {
  return unit === 'layoutPx' ? `${valuePt}px` : `${valuePt}pt`;
}

export function runStyleToCss(
  props: RunProps,
  theme?: Theme,
  options?: RunStyleOptions,
): React.CSSProperties {
  const lengthUnit: LengthUnit = options?.lengthUnit ?? 'pt';
  const css: React.CSSProperties = {};
  if (props.bold) css.fontWeight = 'bold';
  if (props.italic) css.fontStyle = 'italic';
  
  if (props.strike || props.dstrike) {
    css.textDecoration = 'line-through';
  } else if (props.underline && props.underline.style !== 'none') {
    css.textDecoration = 'underline';
    if (props.underline.style === 'double') {
      css.textDecorationStyle = 'double';
    } else if (props.underline.style === 'dotted' || props.underline.style === 'dottedHeavy') {
      css.textDecorationStyle = 'dotted';
    } else if (props.underline.style === 'dash') {
      css.textDecorationStyle = 'dashed';
    }
  }

  if (props.color) {
    const colorStr = resolveColor(props.color);
    if (colorStr) css.color = colorStr;
  }
  if (props.highlight) {
    // Basic mapping for OOXML highlights
    css.backgroundColor = props.highlight === 'yellow' ? 'yellow'
      : props.highlight === 'green' ? 'green'
      : props.highlight === 'cyan' ? 'cyan'
      : props.highlight === 'magenta' ? 'magenta'
      : props.highlight === 'blue' ? 'blue'
      : props.highlight === 'red' ? 'red'
      : props.highlight === 'darkBlue' ? 'darkblue'
      : props.highlight === 'darkCyan' ? 'darkcyan'
      : props.highlight === 'darkGreen' ? 'darkgreen'
      : props.highlight === 'darkMagenta' ? 'darkmagenta'
      : props.highlight === 'darkRed' ? 'darkred'
      : props.highlight === 'darkYellow' ? 'goldenrod'
      : props.highlight === 'darkGray' ? 'darkgray'
      : props.highlight === 'lightGray' ? 'lightgray'
      : props.highlight === 'black' ? 'black'
      : undefined;
  } else if (props.shd?.fill && props.shd.fill !== 'auto') {
    const fill = resolveColor(props.shd.fill);
    if (fill) css.backgroundColor = fill;
  }

  if (props.sz !== undefined) {
    css.fontSize = emitLength(halfPointToPt(props.sz), lengthUnit);
  } else {
    // CRITICAL: emit the paginator's default size explicitly. Otherwise the
    // browser inherits its UA default (16px) and paints glyphs ~45% larger
    // than the 11pt boxes the paginator reserved → vertical line crash AND
    // right-edge clipping for every run that omits <w:sz>.
    css.fontSize = emitLength(DEFAULT_FONT_SIZE_PT, lengthUnit);
  }
  
  // Resolve effective font family using same logic as layout measurement
  // (themed attrs → literal attrs → Calibri). This MUST stay in sync with
  // src/docx/layout/fontResolution.ts so DOM glyph widths match the
  // widths the paginator computed.
  //
  // ALWAYS emit font-family — even when it resolves to the Calibri default —
  // so the DOM never inherits the browser/OS UI font (Segoe UI, San Francisco,
  // etc.) whose advance widths differ from the Carlito substitute that the
  // paginator measured against. (3.0.5 defect root cause.)
  const resolvedFamily = effectiveFontFamily(props, theme, 'latin');
  css.fontFamily = buildFontFamilyStack(resolvedFamily);
  
  if (props.spacing !== undefined) {
    css.letterSpacing = emitLength(twipToPt(props.spacing), lengthUnit);
  }
  
  if (props.vertAlign === 'superscript') {
    css.verticalAlign = 'super';
    css.fontSize = 'smaller';
  } else if (props.vertAlign === 'subscript') {
    css.verticalAlign = 'sub';
    css.fontSize = 'smaller';
  }

  if (props.caps) css.textTransform = 'uppercase';
  if (props.smallCaps) css.fontVariant = 'small-caps';

  // DOCX-12 — these were parsed nowhere (see `model/styles.ts`'s `RunProps`
  // doc comments) so a run carrying any of them not only reverted on save,
  // it also never looked right on screen while open. Best-effort CSS
  // approximations, not pixel-exact reproductions of Word's own rendering.
  if (props.outline) {
    // Hollow/outlined characters: stroke the glyph, hide its fill.
    (css as Record<string, string>).WebkitTextStroke = '1px currentColor';
    (css as Record<string, string>).WebkitTextFillColor = 'transparent';
  }
  if (props.emboss) {
    css.textShadow = '1px 1px 0 rgba(255,255,255,0.75), -1px -1px 0 rgba(0,0,0,0.55)';
    if (!props.color) css.color = 'transparent';
  } else if (props.imprint) {
    css.textShadow = '-1px -1px 0 rgba(255,255,255,0.75), 1px 1px 0 rgba(0,0,0,0.55)';
    if (!props.color) css.color = 'transparent';
  }
  if (props.em && props.em !== 'none') {
    const mark = props.em === 'comma' ? 'filled comma'
      : props.em === 'circle' ? 'filled circle'
      : 'filled dot'; // 'dot' and 'underDot' both use a filled dot glyph
    (css as Record<string, string>).textEmphasis = mark;
    if (props.em === 'underDot') {
      (css as Record<string, string>).textEmphasisPosition = 'under';
    }
  }
  if (props.bdr) {
    const border = borderToCss(props.bdr);
    if (border) {
      css.border = border;
      css.padding = '0 1px';
    }
  }
  if (props.charScale !== undefined && props.charScale !== 100) {
    css.display = 'inline-block';
    css.transform = `scaleX(${props.charScale / 100})`;
  }

  return css;
}

export function paraStyleToCss(props: ParaProps): React.CSSProperties {
  const css: React.CSSProperties = {};
  
  if (props.jc) {
    css.textAlign = props.jc === 'both' ? 'justify' : props.jc === 'start' ? 'left' : props.jc === 'end' ? 'right' : props.jc === 'center' ? 'center' : 'left';
  }
  
  if (props.spacing) {
    if (props.spacing.before !== undefined) css.marginTop = `${twipToPt(props.spacing.before)}pt`;
    if (props.spacing.after !== undefined) css.marginBottom = `${twipToPt(props.spacing.after)}pt`;
    if (props.spacing.line !== undefined) {
      if (props.spacing.lineRule === 'exact') {
        css.lineHeight = `${twipToPt(props.spacing.line)}pt`;
      } else {
        // approx 240 twips = 1.0 lines
        css.lineHeight = props.spacing.line / 240;
      }
    }
  }

  if (props.ind) {
    if (props.ind.left !== undefined) css.marginLeft = `${twipToPt(props.ind.left)}pt`;
    if (props.ind.right !== undefined) css.marginRight = `${twipToPt(props.ind.right)}pt`;
    
    if (props.ind.firstLine !== undefined) {
      css.textIndent = `${twipToPt(props.ind.firstLine)}pt`;
    } else if (props.ind.hanging !== undefined) {
      css.textIndent = `-${twipToPt(props.ind.hanging)}pt`;
      css.paddingLeft = `${twipToPt(props.ind.hanging)}pt`;
    }
  }

  if (props.shd?.fill && props.shd.fill !== 'auto') {
    const fill = resolveColor(props.shd.fill);
    if (fill) css.backgroundColor = fill;
  }
  
  if (props.pBdr) {
    if (props.pBdr.top) css.borderTop = borderToCss(props.pBdr.top);
    if (props.pBdr.bottom) css.borderBottom = borderToCss(props.pBdr.bottom);
    if (props.pBdr.left) css.borderLeft = borderToCss(props.pBdr.left);
    if (props.pBdr.right) css.borderRight = borderToCss(props.pBdr.right);
  }

  if (props.pageBreakBefore) {
    css.breakBefore = 'page';
  }

  return css;
}

export function tableStyleToCss(props: TableProps): React.CSSProperties {
  const css: React.CSSProperties = {
    borderCollapse: 'collapse',
  };

  if (props.jc) {
    if (props.jc === 'center') {
      css.marginLeft = 'auto';
      css.marginRight = 'auto';
    } else if (props.jc === 'end') {
      css.marginLeft = 'auto';
    }
  }

  if (props.tblW?.type === 'dxa' && typeof props.tblW.value === 'number') {
    css.width = `${twipToPt(props.tblW.value)}pt`;
  } else if (props.tblW?.type === 'pct' && typeof props.tblW.value === 'number') {
    css.width = `${props.tblW.value / 50}%`; // 5000 = 100%
  }

  if (props.shd?.fill && props.shd.fill !== 'auto') {
    const fill = resolveColor(props.shd.fill);
    if (fill) css.backgroundColor = fill;
  }

  return css;
}

export function cellStyleToCss(props: TableCellProps): React.CSSProperties {
  const css: React.CSSProperties = {};
  
  if (props.tcW?.type === 'dxa' && typeof props.tcW.value === 'number') {
    css.width = `${twipToPt(props.tcW.value)}pt`;
  } else if (props.tcW?.type === 'pct' && typeof props.tcW.value === 'number') {
    css.width = `${props.tcW.value / 50}%`;
  }

  if (props.vAlign) {
    css.verticalAlign = props.vAlign === 'center' ? 'middle' : props.vAlign === 'bottom' ? 'bottom' : 'top';
  }

  if (props.shd?.fill && props.shd.fill !== 'auto') {
    const fill = resolveColor(props.shd.fill);
    if (fill) css.backgroundColor = fill;
  }

  if (props.tcBorders) {
    if (props.tcBorders.top) css.borderTop = borderToCss(props.tcBorders.top);
    if (props.tcBorders.bottom) css.borderBottom = borderToCss(props.tcBorders.bottom);
    if (props.tcBorders.left) css.borderLeft = borderToCss(props.tcBorders.left);
    if (props.tcBorders.right) css.borderRight = borderToCss(props.tcBorders.right);
  }

  return css;
}

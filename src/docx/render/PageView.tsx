import React, { useMemo, useState } from 'react';
import type {
  Document,
  Hyperlink,
  HyperlinkChild,
  ParagraphChild,
} from '../model/document';
import type { Page, PageTableRef, PageTableRowRef } from '../layout/pageTypes';
import type { LineBox } from '../layout/types';
import { computePageFloats } from '../layout/floats';
import type { Relationship } from '../parser/relationships';
import type { Theme } from '../parser/theme';
import { resolveStretchableSpaceIndices } from '../layout/breakLines';
import { MARKER_RUN_INDEX } from '../layout/listMarkers';

import { borderToCss, revisionStyleToCss, type RevisionRenderKind, runStyleToCss } from './style';
import { AnchoredDrawing } from './AnchoredDrawing';
import { DrawingAnchorMarker, InlineDrawing } from './InlineDrawing';
import './__styles__/page-view.css';

export type PageViewProps = {
  page: Page;
  zoom: number;
  document: Document;
  theme?: Theme;
  /** Document-level relationships, used to resolve an external hyperlink's URL (D6). */
  relationships?: ReadonlyArray<Relationship>;
  /**
   * DXE-14 — called with a table's own path (`PageTableRef.blockPath`), the
   * grid column index and the new width in twips once the user finishes
   * dragging that column's resize handle (never mid-drag — see
   * `TableColumnResizeHandle`'s own doc comment for why only the handle
   * itself previews live). Omit to render every table with no resize
   * handles at all: a read-only preview (print, an eventual "view mode")
   * has no business offering a drag interaction that edits the document.
   */
  onResizeTableColumn?: (tablePath: ReadonlyArray<number>, columnIndex: number, widthTwips: number) => void;
};

type RevisionRunMeta = {
  readonly kind: RevisionRenderKind;
  readonly author?: string;
};

type HyperlinkRunMeta = {
  readonly href: string;
  readonly isExternal: boolean;
  readonly tooltip?: string;
};

type RunMeta = {
  readonly revision?: RevisionRunMeta;
  readonly hyperlink?: HyperlinkRunMeta;
};

// CSS px is 96dpi, pt is 72dpi. So 1pt = 1.333px.
const PT_TO_PX = 4 / 3;
// OOXML twips are 1/20 of a point.
const TWIPS_PER_PT = 20;
// DXE-14 — mirrors `commands.ts`'s own `MIN_COLUMN_WIDTH_TWIPS` (180 twips):
// clamped here too so a drag never submits a resize `applyEditorCommand`
// would reject outright (which would silently no-op — DocxViewer's
// `applyEditorCommand` swallows a throwing command rather than surfacing
// it, so the drag would otherwise just appear to do nothing).
const MIN_COLUMN_WIDTH_PT = 9;

/**
 * Renders a tab's leader (D24/DXL-16) as a bottom border spanning its
 * measured width, rather than repeating literal `.`/`-`/`_` glyphs — this
 * viewer doesn't know the leader glyph's per-character advance width at
 * render time (only `breakLines.ts`'s layout-time font metrics do), so a
 * border reliably fills the exact tab width instead of over/under-shooting
 * with a whole number of repeated characters. `text-bottom` vertical
 * alignment on the tab span (see the caller) keeps the rule close to the
 * text baseline rather than the bottom of the (taller) line box.
 */
function tabLeaderToCss(leader: 'none' | 'dot' | 'hyphen' | 'underscore'): React.CSSProperties {
  if (leader === 'dot') {
    return { borderBottom: '1px dotted currentColor' };
  }
  if (leader === 'hyphen') {
    return { borderBottom: '1px dashed currentColor' };
  }
  if (leader === 'underscore') {
    return { borderBottom: '1px solid currentColor' };
  }
  return {};
}

const EMPTY_STRETCH_INDICES: ReadonlySet<number> = new Set();
const EMPTY_RELATIONSHIPS: ReadonlyArray<Relationship> = [];
const EMPTY_BOOKMARK_NAMES: ReadonlyArray<string> = [];

/** Prefix for a bookmark's rendered anchor `id` — an internal hyperlink's
 * `href` points at `#${BOOKMARK_ANCHOR_ID_PREFIX}${anchor}`, letting the
 * browser's own same-page fragment navigation do the scrolling (D6). */
const BOOKMARK_ANCHOR_ID_PREFIX = 'docx-bookmark-';

type RenderLineFn = (
  line: LineBox,
  topPt: number,
  leftPt: number,
  key: string,
  paragraphPath?: ReadonlyArray<number>,
  paragraphLineIndex?: number,
) => React.ReactNode;

/**
 * Renders a single `PageTableRef` slice as a real `<table>`. Borders are
 * collapsed; column widths come from the source `LaidOutTable` so cells
 * align with the paginator's geometry. Cell content is replayed through
 * the same `renderLine` used by paragraph text so glyph measurements
 * remain consistent with the rest of the page.
 *
 * Cells whose `shouldRender` is `false` (vMerge=continue continuations)
 * are skipped — their visual region is already owned by the originating
 * `vMerge=restart` cell whose `rowSpan` extends down through them.
 *
 * Coordinates: `tableRef.topPt` / `tableRef.leftPt` are page-absolute and
 * are interpreted as internal "px" units inside the PageView scale wrapper
 * (see style.ts `layoutPx` notes).
 */
function renderPageTable(
  tableRef: PageTableRef,
  key: string,
  renderLine: RenderLineFn,
  scale: number,
  onResizeTableColumn?: (tablePath: ReadonlyArray<number>, columnIndex: number, widthTwips: number) => void,
): React.ReactNode {
  const wrapperStyle: React.CSSProperties = {
    position: 'absolute',
    top: `${tableRef.topPt}px`,
    left: `${tableRef.leftPt}px`,
    width: `${tableRef.widthPt}px`,
  };

  const tableStyle: React.CSSProperties = {
    borderCollapse: 'collapse',
    tableLayout: 'fixed',
    width: `${tableRef.widthPt}px`,
  };

  if (tableRef.shadingFill !== undefined) {
    tableStyle.backgroundColor = tableRef.shadingFill;
  }

  // DXE-14 — a continuation fragment (this table's rows spilling onto a
  // later page) shares the same underlying table, but resizing from a
  // continuation's handles would need this fragment's OWN column offset
  // reconciled with the source table's, which isn't worth the complexity
  // for what is, in practice, always reachable from the table's first page
  // too. Handles only render on the fragment that starts the table.
  const canResize = onResizeTableColumn !== undefined && !tableRef.isContinuation;

  return (
    <div
      key={key}
      className="docx-page__table-wrapper"
      data-block-path={tableRef.blockPath.join(',')}
      data-continuation={tableRef.isContinuation ? 'true' : 'false'}
      style={wrapperStyle}
    >
      <table className="docx-page__table" style={tableStyle}>
        <colgroup>
          {tableRef.columnWidthsPt.map((colWidth, colIdx) => (
            <col key={colIdx} style={{ width: `${colWidth}px` }} />
          ))}
        </colgroup>
        <tbody>
          {tableRef.rows.map((rowRef, rowIdx) =>
            renderPageTableRow(rowRef, rowIdx, tableRef, renderLine),
          )}
        </tbody>
      </table>
      {canResize &&
        tableRef.columnWidthsPt.map((widthPt, colIdx) => (
          <TableColumnResizeHandle
            key={`resize-${colIdx}`}
            tablePath={tableRef.blockPath}
            columnIndex={colIdx}
            leftPt={tableRef.columnWidthsPt.slice(0, colIdx + 1).reduce((sum, w) => sum + w, 0)}
            heightPt={tableRef.heightPt}
            currentWidthPt={widthPt}
            scale={scale}
            onResize={onResizeTableColumn}
          />
        ))}
    </div>
  );
}

/**
 * DXE-14 — a draggable vertical guide over a table's column border. Only the
 * guide line itself tracks the pointer while dragging (via local `dragPt`
 * state) — the table's actual columns don't live-reflow, since that would
 * mean re-pagination on every `pointermove`. The real `resize-table-column`
 * command (an undoable, single History step) only fires once, on release,
 * from `initialWidthPt + totalDelta` — never as an incremental sequence of
 * commands per pixel moved.
 *
 * Listens on `window` rather than relying on pointer capture on the handle
 * itself: simpler, and unaffected by the cursor moving off the (2px-wide)
 * handle mid-drag.
 */
function TableColumnResizeHandle({
  tablePath,
  columnIndex,
  leftPt,
  heightPt,
  currentWidthPt,
  scale,
  onResize,
}: {
  tablePath: ReadonlyArray<number>;
  columnIndex: number;
  /** The x-position, in the table's own pt coordinate space, of this
   * column's right edge — where the handle sits at rest. */
  leftPt: number;
  heightPt: number;
  currentWidthPt: number;
  scale: number;
  onResize: (tablePath: ReadonlyArray<number>, columnIndex: number, widthTwips: number) => void;
}) {
  const [dragDeltaPt, setDragDeltaPt] = useState<number | null>(null);

  const handlePointerDown = (event: React.PointerEvent<HTMLDivElement>) => {
    // Only the primary mouse button / a genuine touch-or-pen contact starts
    // a drag — a right-click here should still fall through to whatever
    // context menu the table itself offers, not silently swallow the event.
    if (event.button !== 0) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    const startClientX = event.clientX;

    const handleMove = (moveEvent: PointerEvent) => {
      setDragDeltaPt((moveEvent.clientX - startClientX) / scale);
    };
    const handleUp = (upEvent: PointerEvent) => {
      const deltaPt = (upEvent.clientX - startClientX) / scale;
      const nextWidthPt = Math.max(MIN_COLUMN_WIDTH_PT, currentWidthPt + deltaPt);
      onResize(tablePath, columnIndex, Math.round(nextWidthPt * TWIPS_PER_PT));
      setDragDeltaPt(null);
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleUp);
    };

    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleUp);
  };

  const currentLeftPt = leftPt + (dragDeltaPt ?? 0);

  return (
    <div
      className={`docx-page__table-col-resize${dragDeltaPt !== null ? ' docx-page__table-col-resize--active' : ''}`}
      style={{
        position: 'absolute',
        top: 0,
        left: `${currentLeftPt}px`,
        height: `${heightPt}px`,
      }}
      onPointerDown={handlePointerDown}
      role="separator"
      aria-orientation="vertical"
      aria-label={`Resize column ${columnIndex + 1}`}
    />
  );
}

function renderPageTableRow(
  rowRef: PageTableRowRef,
  rowIdx: number,
  tableRef: PageTableRef,
  renderLine: RenderLineFn,
): React.ReactNode {
  const rowStyle: React.CSSProperties = {
    height: `${rowRef.row.heightPt}px`,
  };

  return (
    <tr
      key={`row-${rowIdx}`}
      data-source-row={rowRef.sourceRowIndex}
      data-repeated-header={rowRef.isRepeatedHeader ? 'true' : 'false'}
      style={rowStyle}
    >
      {rowRef.row.cells.map((cell, cellIdx) => {
        if (!cell.shouldRender) {
          return null;
        }

        const cellStyle: React.CSSProperties = {
          width: `${cell.widthPt}px`,
          verticalAlign:
            cell.vAlign === 'center' ? 'middle' : cell.vAlign === 'bottom' ? 'bottom' : 'top',
          paddingTop: `${cell.paddingPt.top}px`,
          paddingRight: `${cell.paddingPt.right}px`,
          paddingBottom: `${cell.paddingPt.bottom}px`,
          paddingLeft: `${cell.paddingPt.left}px`,
        };

        applyCellBorders(cellStyle, cell.borders, tableRef.borders);

        if (cell.shadingFill !== undefined) {
          cellStyle.backgroundColor = cell.shadingFill;
        }

        const colSpanAttr = cell.gridSpan > 1 ? cell.gridSpan : undefined;
        const rowSpanAttr = cell.rowSpan > 1 ? cell.rowSpan : undefined;

        return (
          <td
            key={`cell-${cellIdx}`}
            colSpan={colSpanAttr}
            rowSpan={rowSpanAttr}
            data-col={cell.columnStart}
            style={cellStyle}
          >
            <div
              className="docx-page__table-cell-content"
              style={{
                position: 'relative',
                height: `${Math.max(0, cell.heightPt - cell.paddingPt.top - cell.paddingPt.bottom)}px`,
              }}
            >
              {(() => {
                let cumulativeTop = 0;
                return cell.contentLines.map((line, lineIdx) => {
                  const node = renderLine(
                    line,
                    cumulativeTop,
                    0,
                    `cell-line-${rowIdx}-${cellIdx}-${lineIdx}`,
                  );
                  cumulativeTop += line.lineHeight;
                  return node;
                });
              })()}
            </div>
          </td>
        );
      })}
    </tr>
  );
}

function applyCellBorders(
  cellStyle: React.CSSProperties,
  cellBorders: PageTableRef['rows'][number]['row']['cells'][number]['borders'],
  tableBorders: PageTableRef['borders'],
): void {
  const resolveEdge = (
    edge: 'top' | 'right' | 'bottom' | 'left',
  ): string | undefined => {
    const cellEdge = cellBorders?.[edge];
    if (cellEdge !== undefined) {
      return borderToCss(cellEdge);
    }

    if (tableBorders === undefined) {
      return undefined;
    }

    if (edge === 'top' || edge === 'bottom') {
      const insideH = tableBorders.insideH;
      if (insideH !== undefined) {
        return borderToCss(insideH);
      }
    } else {
      const insideV = tableBorders.insideV;
      if (insideV !== undefined) {
        return borderToCss(insideV);
      }
    }

    return borderToCss(tableBorders[edge]);
  };

  const top = resolveEdge('top');
  const right = resolveEdge('right');
  const bottom = resolveEdge('bottom');
  const left = resolveEdge('left');

  if (top !== undefined) cellStyle.borderTop = top;
  if (right !== undefined) cellStyle.borderRight = right;
  if (bottom !== undefined) cellStyle.borderBottom = bottom;
  if (left !== undefined) cellStyle.borderLeft = left;
}

/**
 * Renders a single Page as a fixed-size DOM sheet.
 * Choice: To avoid multiplying every single coordinate by the zoom factor,
 * we set the outer container to the scaled size, and use a CSS transform
 * to scale the inner contents. This keeps the internal layout exactly 1:1
 * with the pt values calculated by the paginator.
 */
export const PageView: React.FC<PageViewProps> = ({
  page,
  zoom,
  document,
  theme,
  relationships,
  onResizeTableColumn,
}) => {
  const scale = zoom * PT_TO_PX;
  const resolvedRelationships = relationships ?? EMPTY_RELATIONSHIPS;
  const runMetaByParagraph = useMemo(
    () => collectRunMetaByParagraph(document, resolvedRelationships),
    [document, resolvedRelationships],
  );
  const bookmarkNamesByParagraph = useMemo(() => collectBookmarkNamesByParagraph(document), [document]);
  // D4/DXL-03: computed fresh per page from the already-laid-out `page`
  // (see `floats.ts`'s doc comment for why this doesn't need to touch
  // paginate.ts/breakLines.ts). Mirrors the existing per-page memoization
  // pattern above (`runMetaByParagraph`/`bookmarkNamesByParagraph`), which
  // also walks the whole document once per page render.
  const pageFloats = useMemo(() => computePageFloats(page, document), [page, document]);
  const behindDocFloats = useMemo(() => pageFloats.filter((f) => f.behindDoc), [pageFloats]);
  const frontFloats = useMemo(() => pageFloats.filter((f) => !f.behindDoc), [pageFloats]);

  const outerStyle = useMemo(() => ({
    width: `${page.sizePt.width * scale}px`,
    height: `${page.sizePt.height * scale}px`,
    // Wave F.2 — pair with `content-visibility: auto` in page-view.css so
    // off-screen pages reserve the correct vertical space and don't cause
    // scroll-jump when content-visibility skips their layout.
    containIntrinsicSize: `${page.sizePt.width * scale}px ${page.sizePt.height * scale}px`,
  }), [page.sizePt.width, page.sizePt.height, scale]);

  const innerStyle = useMemo(() => ({
    transform: `scale(${scale})`,
    transformOrigin: 'top left',
    width: `${page.sizePt.width}px`,
    height: `${page.sizePt.height}px`,
  }), [scale, page.sizePt.width, page.sizePt.height]);

  const renderLine = (
    line: LineBox,
    topPt: number,
    leftPt: number,
    key: string,
    paragraphPath?: ReadonlyArray<number>,
    paragraphLineIndex?: number,
  ) => {
    const paragraphKey = paragraphPath?.join(',');
    const runMetas = paragraphKey ? runMetaByParagraph.get(paragraphKey) : undefined;
    // Bookmarks anchor at their paragraph's first rendered line — precise
    // enough for "jump to section" navigation without tracking exact
    // character offsets through pagination.
    const bookmarkNames =
      paragraphLineIndex === 0 && paragraphKey
        ? bookmarkNamesByParagraph.get(paragraphKey) ?? EMPTY_BOOKMARK_NAMES
        : EMPTY_BOOKMARK_NAMES;
    // Tall drawings reserve clearance above the text strut; the strut keeps
    // its own line height so the baseline lands where the paginator put it.
    const clearance = line.drawingClearancePt ?? 0;
    // D5: `line.isJustified`/`justificationStretch` were already computed by
    // the paginator but never read here — a "Justify" paragraph rendered
    // with its natural (ragged) width instead of a flush right edge. Widen
    // each stretchable space by the paginator's per-space stretch amount,
    // and widen the line container to match so the flush edge is real
    // (not clipped/overflowing).
    const stretchableIndices = line.isJustified
      ? resolveStretchableSpaceIndices(line.items)
      : EMPTY_STRETCH_INDICES;
    const lineWidthPt =
      stretchableIndices.size > 0
        ? line.width + stretchableIndices.size * line.justificationStretch
        : line.width;
    // D6: wrap a hyperlink-hosted item (including the spaces between its
    // words, so the whole span is clickable with no gaps) in a real <a>.
    // Each item gets its own <a> rather than one <a> spanning the whole
    // hyperlink — adjacent inline <a> tags render and click seamlessly,
    // and this avoids restructuring the per-item render loop into a
    // run-grouping pass.
    const wrapHyperlink = (node: React.ReactNode, runIndex: number, key: number): React.ReactNode => {
      const hyperlink = runMetas?.[runIndex]?.hyperlink;
      if (hyperlink === undefined) {
        return node;
      }
      return (
        <a
          key={key}
          href={hyperlink.href}
          title={hyperlink.tooltip}
          className="docx-hyperlink"
          style={{ cursor: 'pointer' }}
          {...(hyperlink.isExternal ? { target: '_blank', rel: 'noopener noreferrer' } : {})}
        >
          {node}
        </a>
      );
    };

    return (
      <div
        key={key}
        className="docx-page__line"
        data-paragraph-path={paragraphPath?.join(',')}
        {...(line.rtl === true ? { dir: 'rtl' as const } : {})}
        style={{
          top: `${topPt}px`,
          left: `${leftPt}px`,
          width: `${lineWidthPt}px`,
          height: `${line.lineHeight}px`,
          lineHeight: `${line.lineHeight - clearance}px`,
          ...(clearance > 0 ? { paddingTop: `${clearance}px`, boxSizing: 'border-box' as const } : {}),
        }}
      >
        {bookmarkNames.map((name) => (
          <a
            key={`bookmark-${name}`}
            id={`${BOOKMARK_ANCHOR_ID_PREFIX}${name}`}
            aria-hidden="true"
            tabIndex={-1}
            style={{ position: 'absolute', width: 0, height: 0, pointerEvents: 'none' }}
          />
        ))}
        {line.items.map((item, idx) => {
          if (item.kind === 'drawing') {
            if (item.drawing.layout === 'anchor') {
              // D4/DXL-03: the picture itself is painted separately as a
              // page-level float (see `renderFloats` below) — this item
              // only keeps the anchor's one-character offset addressable
              // for the editor.
              return (
                <DrawingAnchorMarker
                  key={`drawing-${item.runIndex}-${item.charStart}`}
                  runIndex={item.runIndex}
                  charStart={item.charStart}
                  charEnd={item.charEnd}
                />
              );
            }
            const revisionKind = runMetas?.[item.runIndex]?.revision?.kind ?? item.runProps._revision;
            return (
              <InlineDrawing
                key={`drawing-${item.runIndex}-${item.charStart}`}
                drawing={item.drawing}
                widthPt={item.width}
                heightPt={item.height}
                runIndex={item.runIndex}
                charStart={item.charStart}
                charEnd={item.charEnd}
                className={revisionKind === undefined ? undefined : `docx-revision--${revisionKind}`}
              />
            );
          }
          if (item.kind === 'word' || item.kind === 'glyph-cluster') {
            const isMarker = item.runIndex === MARKER_RUN_INDEX;
            const runMeta = isMarker ? undefined : runMetas?.[item.runIndex];
            const revision = runMeta?.revision;
            const revisionKind = revision?.kind ?? item.runProps._revision;
            const revisionClass =
              revisionKind === 'ins'
                ? ' docx-revision--ins'
                : revisionKind === 'del'
                  ? ' docx-revision--del'
                  : '';
            const markerClass = isMarker ? ' docx-list-marker' : '';
            const style =
              revisionKind === undefined
                ? runStyleToCss(item.runProps, theme, { lengthUnit: 'layoutPx' })
                : {
                    ...runStyleToCss(item.runProps, theme, { lengthUnit: 'layoutPx' }),
                    ...revisionStyleToCss(revisionKind, revision?.author),
                  };
            const wordSpan = (
              <span
                key={idx}
                className={`docx-run${revisionClass}${markerClass}`}
                data-run-index={item.runIndex}
                data-char-start={item.charStart}
                data-char-end={item.charEnd}
                style={{
                  // Do NOT lock the rendered width: the paginator already
                  // measured this word using canvas `measureText` under the
                  // same font the browser will paint with, so the natural
                  // browser advance MUST equal `item.width`. Locking the width
                  // (as 3.0.0–3.0.2 did) clipped any sub-pixel browser drift
                  // and squeezed accented French glyphs whose layout-time
                  // width was zero. Letting the browser flow the text natively
                  // also fixes "INTROD…" right-edge clipping.
                  ...style,
                  display: 'inline-block',
                  whiteSpace: 'pre',
                }}
              >
                {item.text}
              </span>
            );

            return wrapHyperlink(wordSpan, item.runIndex, idx);
          }
          if (item.kind === 'space') {
            // Spaces keep their measured width, EXCEPT on a justified line's
            // stretchable spaces, which absorb the paginator's per-space
            // stretch amount so the line's flush-right edge is real. Emit a
            // real space character so accessibility tools and copy/paste see
            // the document's actual whitespace, not an invisible spacer.
            const width = stretchableIndices.has(idx)
              ? item.width + line.justificationStretch
              : item.width;
            // USR-05 — spaces carry the same run/char addressing as words so
            // a selection or caret that starts/ends on a space maps back to a
            // model position (unaddressed spaces made such selections resolve
            // to null, silently dropping formatting commands).
            const spaceSpan = (
              <span
                key={idx}
                className="docx-space"
                data-run-index={item.runIndex}
                data-char-start={item.charOffset}
                data-char-end={item.charOffset + 1}
                style={{ display: 'inline-block', width: `${width}px`, whiteSpace: 'pre' }}
              >
                {' '}
              </span>
            );
            return wrapHyperlink(spaceSpan, item.runIndex, idx);
          }
          if (item.kind === 'tab') {
            return (
              <span
                key={idx}
                className="docx-tab"
                data-run-index={item.runIndex}
                data-char-start={item.charOffset}
                data-char-end={item.charOffset + 1}
                style={{
                  display: 'inline-block',
                  width: `${item.width}px`,
                  verticalAlign: 'text-bottom',
                  ...tabLeaderToCss(item.leader),
                }}
              />
            );
          }
          if (item.kind === 'hyphen-opportunity') {
            return null;
          }
          return null;
        })}
      </div>
    );
  };

  const headerHeight = page.headerLines.reduce((acc, l) => acc + l.lineHeight, 0);
  const footerHeight = page.footerLines.reduce((acc, l) => acc + l.lineHeight, 0);  return (
    <div className="docx-page" style={outerStyle}>
      <div style={innerStyle}>
        {/* Header */}
        {page.headerLines.length > 0 && (
          <div
            className="docx-page__header"
            style={{
              top: `${page.marginsPt.header}px`,
              left: `${page.marginsPt.left}px`,
              width: `${page.sizePt.width - page.marginsPt.left - page.marginsPt.right}px`,
              height: `${headerHeight}px`,
            }}
          >
            {page.headerLines.map((line, idx) => {
              const top = page.headerLines.slice(0, idx).reduce((sum, l) => sum + l.lineHeight, 0);
              return renderLine(line, top, line.leftOffsetPt ?? 0, `header-${idx}`);
            })}
          </div>
        )}

        {/* Floating drawings behind the text (D4/DXL-03: wp:anchor's
            behindDoc="1") — placed before the columns in DOM order so
            static/auto-z-index text paints over them. */}
        {behindDocFloats.map((pageFloat, floatIdx) => (
          <AnchoredDrawing key={`float-behind-${pageFloat.blockIndex}-${floatIdx}`} pageFloat={pageFloat} />
        ))}

        {/* Columns */}
        {page.columns.map((col, colIdx) => (
          <div
            key={`col-${colIdx}`}
            className="docx-page__column"
            style={{
              top: `${page.marginsPt.top}px`,
              left: `${col.leftPt}px`,
              width: `${col.widthPt}px`,
              height: `${page.sizePt.height - page.marginsPt.top - page.marginsPt.bottom}px`,
            }}
          >
            {col.lines.map((lineRef, lineIdx) =>
              renderLine(
                lineRef.line,
                // `lineRef.topPt` is page-absolute (paginate.ts's `placeLine`
                // starts it at `marginsPt.top` + any reserved header height),
                // and this column `<div>` is itself already positioned at
                // `top: marginsPt.top` above — so, exactly like `leftPt -
                // col.leftPt` on the next line, this needs to become
                // column-relative or the top margin gets applied twice and
                // every line renders `marginsPt.top` further down than it
                // should (D4/DXL-03 found this while adding floats.ts, which
                // renders page-absolute floats directly under the page layer
                // with no such wrapper — see floats.ts's `paragraphRect`
                // comment for the coupled reasoning).
                lineRef.topPt - page.marginsPt.top,
                lineRef.leftPt - col.leftPt,
                `line-${colIdx}-${lineIdx}`,
                lineRef.paragraphPath,
                lineRef.lineIndex,
              )
            )}
          </div>
        ))}

        {/* Tables — rendered as page-absolute siblings (their topPt/leftPt
            are page coordinates, not column-relative) so each fragment
            paints exactly where the paginator placed it. */}
        {page.columns.flatMap((col, colIdx) =>
          col.tables.map((tableRef, tableIdx) =>
            renderPageTable(tableRef, `table-${colIdx}-${tableIdx}`, renderLine, scale, onResizeTableColumn),
          ),
        )}

        {/* Floating drawings in front of the text (the common case:
            wp:anchor's default behindDoc="0") — placed after the columns
            and tables so they paint on top. */}
        {frontFloats.map((pageFloat, floatIdx) => (
          <AnchoredDrawing key={`float-front-${pageFloat.blockIndex}-${floatIdx}`} pageFloat={pageFloat} />
        ))}

        {/* Footnote area (D11 milestone 3/DXL-09) — sits just above the
            footer, sized to exactly the space `paginate.ts`'s
            `reserveFootnotesForLine` reserved for it (the last line's
            `topPt + lineHeight`, which already includes the leading
            separator gap — see `buildPageFootnoteLines`'s doc comment). */}
        {page.hasFootnoteSeparator && page.footnoteLines.length > 0 && (() => {
          const lastFootnoteLine = page.footnoteLines[page.footnoteLines.length - 1];
          const footnoteAreaHeight = lastFootnoteLine.topPt + lastFootnoteLine.line.lineHeight;
          return (
            <div
              className="docx-page__footnotes"
              style={{
                position: 'absolute',
                top: `${page.sizePt.height - page.marginsPt.bottom - footnoteAreaHeight}px`,
                left: `${page.marginsPt.left}px`,
                width: `${page.sizePt.width - page.marginsPt.left - page.marginsPt.right}px`,
                height: `${footnoteAreaHeight}px`,
              }}
            >
              <div
                className="docx-page__footnote-separator"
                style={{ position: 'absolute', top: 0, left: 0, width: '144px', borderTop: '1px solid currentColor' }}
              />
              {page.footnoteLines.map((footnoteLine, idx) =>
                renderLine(footnoteLine.line, footnoteLine.topPt, footnoteLine.leftPt, `footnote-${idx}`),
              )}
            </div>
          );
        })()}

        {/* Footer */}
        {page.footerLines.length > 0 && (
          <div
            className="docx-page__footer"
            style={{
              top: `${page.sizePt.height - page.marginsPt.footer - footerHeight}px`,
              left: `${page.marginsPt.left}px`,
              width: `${page.sizePt.width - page.marginsPt.left - page.marginsPt.right}px`,
              height: `${footerHeight}px`,
            }}
          >
            {page.footerLines.map((line, idx) => {
              const top = page.footerLines.slice(0, idx).reduce((sum, l) => sum + l.lineHeight, 0);
              return renderLine(line, top, line.leftOffsetPt ?? 0, `footer-${idx}`);
            })}
          </div>
        )}
      </div>
    </div>
  );
};

/**
 * Walks every paragraph once, producing per-run-index metadata (tracked-
 * change revision AND/OR enclosing hyperlink target) in the exact same
 * flattened run order paginate.ts's `collectParagraphRuns`/
 * `collectHyperlinkRuns` produce `LineItem.runIndex` values in — run, then
 * a hyperlink's own runs, then a revision's own runs, all in document
 * order. The two concerns share one walk (rather than two independent
 * ones) so they can never drift apart on what "the i-th run" means.
 */
function collectRunMetaByParagraph(
  document: Document,
  relationships: ReadonlyArray<Relationship>,
): ReadonlyMap<string, ReadonlyArray<RunMeta | undefined>> {
  const paragraphs = new Map<string, ReadonlyArray<RunMeta | undefined>>();

  for (const section of document.sections) {
    section.blocks.forEach((block, blockIndex) => {
      if (block.kind !== 'paragraph') {
        return;
      }

      const runs: Array<RunMeta | undefined> = [];
      appendParagraphChildRuns(block.children, runs, relationships, {});
      paragraphs.set(String(blockIndex), runs);
    });
  }

  return paragraphs;
}

function appendParagraphChildRuns(
  children: ReadonlyArray<ParagraphChild>,
  runs: Array<RunMeta | undefined>,
  relationships: ReadonlyArray<Relationship>,
  context: RunMeta,
): void {
  for (const child of children) {
    if (child.kind === 'run') {
      runs.push(hasRunMeta(context) ? context : undefined);
      continue;
    }

    if (child.kind === 'hyperlink') {
      const hyperlink = resolveHyperlinkMeta(child, relationships);
      appendHyperlinkChildRuns(child.children, runs, {
        ...context,
        ...(hyperlink !== undefined ? { hyperlink } : {}),
      });
      continue;
    }

    if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
      appendParagraphChildRuns(getRevisionChildren(child), runs, relationships, {
        ...context,
        revision: {
          kind: child.kind === 'ins-revision' ? 'ins' : 'del',
          ...(child.author !== undefined ? { author: child.author } : {}),
        },
      });
    }
  }
}

function appendHyperlinkChildRuns(
  children: ReadonlyArray<HyperlinkChild>,
  runs: Array<RunMeta | undefined>,
  context: RunMeta,
): void {
  for (const child of children) {
    if (child.kind === 'run') {
      runs.push(hasRunMeta(context) ? context : undefined);
    }
  }
}

function hasRunMeta(context: RunMeta): boolean {
  return context.revision !== undefined || context.hyperlink !== undefined;
}

function getRevisionChildren(child: Extract<ParagraphChild, { kind: 'ins-revision' | 'del-revision' }>): ReadonlyArray<ParagraphChild> {
  return child.children as ReadonlyArray<ParagraphChild>;
}

/**
 * External hyperlink targets come straight from the (untrusted) document's
 * own relationship parts — a crafted `.docx` could set one to
 * `javascript:`/`vbscript:`/`data:` etc. Chromium already refuses to
 * execute a `javascript:` URI opened via `target="_blank"`, and Electron's
 * own `setWindowOpenHandler` additionally allow-lists http/https before
 * calling `shell.openExternal` (see `electron/main.cjs`'s
 * `isAllowedExternalScheme`) — but this renders the raw string as a
 * literal `href` regardless, so it's still worth validating at the source
 * rather than depending solely on those other layers. Mirrors this
 * codebase's existing convention (`OdtViewer.tsx`'s DOMPurify sanitization,
 * `main.cjs`'s own scheme allow-list) of never trusting a URL scheme from
 * file content.
 */
const SAFE_HYPERLINK_SCHEMES: ReadonlySet<string> = new Set(['http:', 'https:', 'mailto:']);

function isSafeHyperlinkHref(href: string): boolean {
  try {
    // A base is required for a protocol-relative or bare host string; a
    // genuinely relative (schemeless) target is not a scheme we allow, so
    // this intentionally has no fallback base that would let one through.
    return SAFE_HYPERLINK_SCHEMES.has(new URL(href).protocol);
  } catch {
    return false;
  }
}

/**
 * Resolves a `Hyperlink` node to a renderable target: an external URL via
 * the document's relationships (only when the relationship is genuinely
 * marked `TargetMode="External"`, matching how DOCX generators — including
 * the one that built this wave's corpus fixture — author external
 * hyperlinks), or an internal same-document anchor via `w:anchor`
 * (resolved against a bookmark's rendered `id`, see
 * `BOOKMARK_ANCHOR_ID_PREFIX`). Returns `undefined` for a hyperlink this
 * viewer can't safely resolve (e.g. a relationship id with no matching
 * External relationship, or an unsafe URL scheme — see
 * `isSafeHyperlinkHref`).
 */
function resolveHyperlinkMeta(
  hyperlink: Hyperlink,
  relationships: ReadonlyArray<Relationship>,
): HyperlinkRunMeta | undefined {
  if (hyperlink.relationshipId !== undefined) {
    const relationship = relationships.find((candidate) => candidate.id === hyperlink.relationshipId);
    if (
      relationship === undefined ||
      relationship.targetMode !== 'External' ||
      !isSafeHyperlinkHref(relationship.target)
    ) {
      return undefined;
    }
    return {
      href: relationship.target,
      isExternal: true,
      ...(hyperlink.tooltip !== undefined ? { tooltip: hyperlink.tooltip } : {}),
    };
  }

  if (hyperlink.anchor !== undefined) {
    return {
      href: `#${BOOKMARK_ANCHOR_ID_PREFIX}${hyperlink.anchor}`,
      isExternal: false,
      ...(hyperlink.tooltip !== undefined ? { tooltip: hyperlink.tooltip } : {}),
    };
  }

  return undefined;
}

function collectBookmarkNamesByParagraph(document: Document): ReadonlyMap<string, ReadonlyArray<string>> {
  const paragraphs = new Map<string, ReadonlyArray<string>>();

  for (const section of document.sections) {
    section.blocks.forEach((block, blockIndex) => {
      if (block.kind !== 'paragraph') {
        return;
      }

      const names = collectParagraphBookmarkNames(block.children);
      if (names.length > 0) {
        paragraphs.set(String(blockIndex), names);
      }
    });
  }

  return paragraphs;
}

function collectParagraphBookmarkNames(children: ReadonlyArray<ParagraphChild>): ReadonlyArray<string> {
  const names: string[] = [];

  for (const child of children) {
    if (child.kind === 'bookmark' && child.boundary === 'start' && child.name !== undefined) {
      names.push(child.name);
      continue;
    }

    if (child.kind === 'hyperlink') {
      for (const hyperlinkChild of child.children) {
        if (
          hyperlinkChild.kind === 'bookmark' &&
          hyperlinkChild.boundary === 'start' &&
          hyperlinkChild.name !== undefined
        ) {
          names.push(hyperlinkChild.name);
        }
      }
    }
  }

  return names;
}

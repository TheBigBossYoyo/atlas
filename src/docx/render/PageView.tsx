import React, { useMemo } from 'react';
import type {
  Document,
  HyperlinkChild,
  Paragraph,
  ParagraphChild,
} from '../model/document';
import type { Page, PageTableRef, PageTableRowRef } from '../layout/pageTypes';
import type { LineBox } from '../layout/types';
import type { Theme } from '../parser/theme';
import { resolveStretchableSpaceIndices } from '../layout/breakLines';

import { borderToCss, revisionStyleToCss, type RevisionRenderKind, runStyleToCss } from './style';
import { InlineDrawing } from './InlineDrawing';
import './__styles__/page-view.css';

export type PageViewProps = {
  page: Page;
  zoom: number;
  document: Document;
  theme?: Theme;
};

type RevisionRunMeta = {
  readonly kind: RevisionRenderKind;
  readonly author?: string;
};

// CSS px is 96dpi, pt is 72dpi. So 1pt = 1.333px.
const PT_TO_PX = 4 / 3;

const EMPTY_STRETCH_INDICES: ReadonlySet<number> = new Set();

type RenderLineFn = (
  line: LineBox,
  topPt: number,
  leftPt: number,
  key: string,
  paragraphPath?: ReadonlyArray<number>,
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
    </div>
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
export const PageView: React.FC<PageViewProps> = ({ page, zoom, document, theme }) => {
  const scale = zoom * PT_TO_PX;
  const revisionRunsByParagraph = useMemo(() => collectRevisionRunsByParagraph(document), [document]);

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
  ) => {
    const paragraphKey = paragraphPath?.join(',');
    const revisionRuns = paragraphKey ? revisionRunsByParagraph.get(paragraphKey) : undefined;
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

    return (
      <div
        key={key}
        className="docx-page__line"
        data-paragraph-path={paragraphPath?.join(',')}
        style={{
          top: `${topPt}px`,
          left: `${leftPt}px`,
          width: `${lineWidthPt}px`,
          height: `${line.lineHeight}px`,
          lineHeight: `${line.lineHeight - clearance}px`,
          ...(clearance > 0 ? { paddingTop: `${clearance}px`, boxSizing: 'border-box' as const } : {}),
        }}
      >
        {line.items.map((item, idx) => {
          if (item.kind === 'drawing') {
            const revisionKind = revisionRuns?.[item.runIndex]?.kind ?? item.runProps._revision;
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
            const revision = revisionRuns?.[item.runIndex];
            const revisionKind = revision?.kind ?? item.runProps._revision;
            const revisionClass =
              revisionKind === 'ins'
                ? ' docx-revision--ins'
                : revisionKind === 'del'
                  ? ' docx-revision--del'
                  : '';
            const style =
              revisionKind === undefined
                ? runStyleToCss(item.runProps, theme, { lengthUnit: 'layoutPx' })
                : {
                    ...runStyleToCss(item.runProps, theme, { lengthUnit: 'layoutPx' }),
                    ...revisionStyleToCss(revisionKind, revision?.author),
                  };
            return (
              <span
                key={idx}
                className={`docx-run${revisionClass}`}
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
            return (
              <span
                key={idx}
                style={{ display: 'inline-block', width: `${width}px`, whiteSpace: 'pre' }}
              >
                {' '}
              </span>
            );
          }
          if (item.kind === 'tab') {
            return <span key={idx} style={{ display: 'inline-block', width: `${item.width}px` }} />;
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
              return renderLine(line, top, 0, `header-${idx}`);
            })}
          </div>
        )}

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
                lineRef.topPt,
                lineRef.leftPt - col.leftPt,
                `line-${colIdx}-${lineIdx}`,
                lineRef.paragraphPath,
              )
            )}
          </div>
        ))}

        {/* Tables — rendered as page-absolute siblings (their topPt/leftPt
            are page coordinates, not column-relative) so each fragment
            paints exactly where the paginator placed it. */}
        {page.columns.flatMap((col, colIdx) =>
          col.tables.map((tableRef, tableIdx) =>
            renderPageTable(tableRef, `table-${colIdx}-${tableIdx}`, renderLine),
          ),
        )}

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
              return renderLine(line, top, 0, `footer-${idx}`);
            })}
          </div>
        )}
      </div>
    </div>
  );
};

function collectRevisionRunsByParagraph(
  document: Document,
): ReadonlyMap<string, ReadonlyArray<RevisionRunMeta | undefined>> {
  const paragraphs = new Map<string, ReadonlyArray<RevisionRunMeta | undefined>>();

  for (const section of document.sections) {
    section.blocks.forEach((block, blockIndex) => {
      if (block.kind !== 'paragraph') {
        return;
      }

      paragraphs.set(String(blockIndex), collectParagraphRevisionRuns(block));
    });
  }

  return paragraphs;
}

function collectParagraphRevisionRuns(
  paragraph: Paragraph,
): ReadonlyArray<RevisionRunMeta | undefined> {
  const runs: Array<RevisionRunMeta | undefined> = [];
  appendParagraphChildRuns(paragraph.children, runs);
  return runs;
}

function appendParagraphChildRuns(
  children: ReadonlyArray<ParagraphChild>,
  runs: Array<RevisionRunMeta | undefined>,
  revision?: RevisionRunMeta,
): void {
  for (const child of children) {
    if (child.kind === 'run') {
      runs.push(revision);
      continue;
    }

    if (child.kind === 'hyperlink') {
      appendHyperlinkChildRuns(child.children, runs, revision);
      continue;
    }

    if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
      appendParagraphChildRuns(
        getRevisionChildren(child),
        runs,
        {
          kind: child.kind === 'ins-revision' ? 'ins' : 'del',
          ...(child.author !== undefined ? { author: child.author } : {}),
        },
      );
    }
  }
}

function appendHyperlinkChildRuns(
  children: ReadonlyArray<HyperlinkChild>,
  runs: Array<RevisionRunMeta | undefined>,
  revision?: RevisionRunMeta,
): void {
  for (const child of children) {
    if (child.kind === 'run') {
      runs.push(revision);
    }
  }
}

function getRevisionChildren(child: Extract<ParagraphChild, { kind: 'ins-revision' | 'del-revision' }>): ReadonlyArray<ParagraphChild> {
  return child.children as ReadonlyArray<ParagraphChild>;
}

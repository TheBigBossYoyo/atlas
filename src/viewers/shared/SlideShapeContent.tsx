/** S3/S4/S6/S21 — renders the text content of a `SlideTextBox`/`SlideTable`: styled runs, bullets, and table grids. */

import type { CSSProperties, ReactNode } from 'react'
import type { SlideParagraph, SlideTableCell, SlideTextRun } from './SlideDeck.types'
import { bulletLabel } from './slideStyleHelpers'

const BULLET_INDENT_PX = 20

function RunSpan({ run, index }: { readonly run: SlideTextRun; readonly index: number }) {
  if (run.text === '\n') {
    return <br />
  }

  const decorations: string[] = []
  if (run.underline) {
    decorations.push('underline')
  }
  if (run.strikethrough) {
    decorations.push('line-through')
  }

  const style: CSSProperties = {
    fontWeight: run.bold ? 700 : undefined,
    fontStyle: run.italic ? 'italic' : undefined,
    textDecoration: decorations.length > 0 ? decorations.join(' ') : undefined,
    color: run.color,
    fontSize: run.fontSizePx,
    fontFamily: run.fontFamily,
  }

  return (
    <span key={index} style={style}>
      {run.text}
    </span>
  )
}

function ParagraphRow({
  paragraph,
  counters,
  fontScale,
}: {
  readonly paragraph: SlideParagraph
  readonly counters: number[]
  readonly fontScale: number
}) {
  const label = bulletLabel(paragraph.bullet, counters)

  const style: CSSProperties = {
    paddingLeft: paragraph.level * BULLET_INDENT_PX,
    marginTop: paragraph.spaceBeforePx,
    marginBottom: paragraph.spaceAfterPx,
    textAlign: paragraph.align,
    fontSize: fontScale !== 1 ? `${fontScale * 100}%` : undefined,
  }

  return (
    <div className="slide-shape__paragraph" style={style}>
      {label !== null && <span className="slide-shape__bullet">{label}</span>}
      {paragraph.runs.map((run, index) => (
        <RunSpan key={index} run={run} index={index} />
      ))}
    </div>
  )
}

export function TextParagraphs({
  paragraphs,
  fontScale = 1,
}: {
  readonly paragraphs: ReadonlyArray<SlideParagraph>
  readonly fontScale?: number
}): ReactNode {
  const counters: number[] = []

  return (
    <>
      {paragraphs.map((paragraph, index) => (
        <ParagraphRow key={index} paragraph={paragraph} counters={counters} fontScale={fontScale} />
      ))}
    </>
  )
}

function TableCellContent({ cell }: { readonly cell: SlideTableCell }) {
  if (cell.runs.length === 0) {
    return <>{cell.text}</>
  }

  return (
    <>
      {cell.runs.map((run, index) => (
        <RunSpan key={index} run={run} index={index} />
      ))}
    </>
  )
}

export function SlideTableGrid({
  rows,
}: {
  readonly rows: ReadonlyArray<ReadonlyArray<SlideTableCell>>
}) {
  return (
    <table className="slide-shape__table">
      <tbody>
        {rows.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {row.map((cell, cellIndex) => (
              <td key={cellIndex} className="slide-shape__table-cell">
                <TableCellContent cell={cell} />
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

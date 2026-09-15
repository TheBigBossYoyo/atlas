/**
 * Frozen-rows fallback rendering (T4/DAT-10 remainder).
 *
 * `@glideapps/glide-data-grid` v6 has `freezeColumns` but no frozen-TOP-rows
 * equivalent at all (only `freezeTrailingRows`, which pins rows to the
 * *bottom*, the opposite of what a spreadsheet's freeze-panes feature
 * means) — confirmed against the installed version's own type definitions.
 * Running a SECOND `DataEditor` instance just for the frozen rows and trying
 * to keep its horizontal scroll pixel-synced with the main one would need
 * either two independently-scrollable canvases fighting each other or a
 * one-directional "read from the main grid, drive the second one" wiring
 * that glide-data-grid's public API doesn't cleanly expose (its own
 * `scrollTo` is cell-granular, not sub-cell-pixel-granular, so it can't
 * reproduce the main grid's exact partial-column scroll offset).
 *
 * So, per the plan's own explicit fallback ("render frozen rows as
 * additional header rows"): this renders the frozen row(s) as a small,
 * genuinely-editable strip of plain `<input>` elements — not a second
 * canvas grid — sitting above the main `DataEditor` and kept in horizontal
 * sync with it via a CSS `translateX`, driven by the main grid's own
 * `onVisibleRegionChanged` (which reports its horizontal scroll offset in
 * pixels — see `SpreadsheetViewer`'s `onVisibleRegionChanged` handler).
 *
 * Deliberately uncontrolled inputs (`defaultValue`, not `value`): the
 * document model is the source of truth once an edit commits (on blur or
 * Enter), but while the user is mid-keystroke this input's own DOM value is
 * authoritative, exactly like glide-data-grid's own cell-edit overlay. The
 * `key` includes the current committed text so the input remounts (picking
 * up the new `defaultValue`) whenever that cell changes from any OTHER
 * source (undo/redo, a paste, a different edit) rather than only on the
 * next full unmount/remount of this component.
 */
import { Fragment } from 'react'

const ROW_HEIGHT_PX = 32

export type FrozenRowsStripProps = {
  /** Just the frozen row(s)' display text — rows[0] is the first frozen row. */
  readonly rows: ReadonlyArray<ReadonlyArray<string>>
  /**
   * Per-cell formula text (no leading `=`), same shape as `rows` — see
   * `useSpreadsheetGrid.ts`'s header for why this matters: without it, an
   * `<input>` for an existing formula cell would show its computed value,
   * and committing it unchanged (blur/Enter) would silently replace the
   * formula with that frozen plain value.
   */
  readonly formulas: ReadonlyArray<ReadonlyArray<string | undefined>>
  readonly columnWidthsPx: ReadonlyArray<number>
  readonly rowMarkerWidthPx: number
  /** Horizontal scroll offset (px) to mirror from the main grid — see module header. */
  readonly translateXPx: number
  readonly onCommit: (row: number, col: number, text: string) => void
}

export function FrozenRowsStrip({
  rows,
  formulas,
  columnWidthsPx,
  rowMarkerWidthPx,
  translateXPx,
  onCommit,
}: FrozenRowsStripProps) {
  if (rows.length === 0) return null

  return (
    <div className="spreadsheet-viewer__frozen-rows" style={{ height: rows.length * ROW_HEIGHT_PX }}>
      <div className="spreadsheet-viewer__frozen-rows-spacer" style={{ width: rowMarkerWidthPx }} />
      <div
        className="spreadsheet-viewer__frozen-rows-track"
        style={{ transform: `translateX(${-translateXPx}px)` }}
      >
        {rows.map((row, r) => (
          <div className="spreadsheet-viewer__frozen-row" key={r}>
            {row.map((text, c) => {
              const formula = formulas[r]?.[c]
              const editableValue = formula !== undefined ? `=${formula}` : text
              return (
                <Fragment key={c}>
                  <input
                    key={`${r}-${c}-${editableValue}`}
                    defaultValue={editableValue}
                    style={{ width: columnWidthsPx[c] ?? 120 }}
                    aria-label={`Frozen row ${r + 1}, column ${c + 1}`}
                    onBlur={(e) => onCommit(r, c, e.currentTarget.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.currentTarget.blur()
                      }
                    }}
                  />
                </Fragment>
              )
            })}
          </div>
        ))}
      </div>
    </div>
  )
}

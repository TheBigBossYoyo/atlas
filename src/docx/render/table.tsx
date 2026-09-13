import React, { useMemo } from 'react';
import type { Table, TableRow, TableCell, Document as DocxDocument } from '../model/document';
import { tableStyleToCss, cellStyleToCss } from './style';
import { ParagraphNode } from './paragraph';

export function TableNode({ table, doc }: { table: Table; doc: DocxDocument }) {
  const effectiveProps = useMemo(() => table.props ?? {}, [table.props]);
  const css = useMemo(() => tableStyleToCss(effectiveProps), [effectiveProps]);

  return (
    <table className="docx-table" style={css}>
      <tbody>
        {table.rows.map((row, i) => {
          if (row.kind === 'table-row') {
            return <TableRowNode key={i} row={row} doc={doc} />;
          }
          return null; // UnknownNode
        })}
      </tbody>
    </table>
  );
}

function TableRowNode({ row, doc }: { row: TableRow; doc: DocxDocument }) {
  return (
    <tr className="docx-table-row" style={{ height: row.props?.trHeight?.val ? `${row.props.trHeight.val / 20}pt` : undefined }}>
      {row.cells.map((cell, i) => {
        if (cell.kind === 'table-cell') {
          return <TableCellNode key={i} cell={cell} doc={doc} />;
        }
        return null;
      })}
    </tr>
  );
}

const TableCellNode = React.memo(function TableCellNode({ cell, doc }: { cell: TableCell; doc: DocxDocument }) {
  const css = useMemo(() => cellStyleToCss(cell.props ?? {}), [cell.props]);
  
  if (cell.props?.vMerge === 'continue') {
    // Usually skipped or rendered empty, depending on gridSpan/vMerge handling.
    // For A.6 we render them but could apply styles.
  }

  return (
    <td 
      className="docx-table-cell" 
      style={css}
      colSpan={cell.props?.gridSpan}
    >
      {cell.blocks.map((block, i) => {
        if (block.kind === 'paragraph') {
          return <ParagraphNode key={i} paragraph={block} doc={doc} />;
        } else if (block.kind === 'table') {
          return <TableNode key={i} table={block} doc={doc} />;
        }
        return null;
      })}
    </td>
  );
});

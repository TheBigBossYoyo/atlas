import { Fragment } from 'react';

import type { ToolbarCommand } from './toolbarTypes';
import { TABLE_EDIT_ACTIONS } from './tableEditActions';

/**
 * DXE-14 — the shared list of table-editing menu buttons, rendered inside
 * both the toolbar's "Table" dropdown (`Toolbar.tsx`) and the DOCX editor's
 * right-click context menu (`DocxViewer.tsx`) so the two entry points can
 * never drift on labels, ordering, or grouping.
 */
export const TableEditMenuItems = ({
  onCommand,
  onAfterCommand,
}: {
  onCommand: (cmd: ToolbarCommand) => void;
  onAfterCommand?: () => void;
}) => {
  return (
    <>
      {TABLE_EDIT_ACTIONS.map((action) => (
        <Fragment key={action.kind}>
          {action.dividerBefore === true && <div className="docx-toolbar__menu-divider" role="separator" />}
          <button
            type="button"
            className="docx-toolbar__menu-item"
            onClick={() => {
              onCommand({ kind: action.kind });
              onAfterCommand?.();
            }}
          >
            {action.label}
          </button>
        </Fragment>
      ))}
    </>
  );
};

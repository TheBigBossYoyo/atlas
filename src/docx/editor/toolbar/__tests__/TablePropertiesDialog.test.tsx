import { render, fireEvent } from '@testing-library/react';
import { describe, it, expect, vi } from 'vitest';

import { TablePropertiesDialog } from '../TablePropertiesDialog';

describe('TablePropertiesDialog — F2', () => {
  it('submits Apply even when the seeded width is not on a 0.1in step boundary', () => {
    // 9000 twips -> 6.25in (twipsToInches rounds to the nearest 0.01in).
    // Word's own default single-column table width lands here, and it isn't
    // a multiple of 0.1 from min 0.1 (nearest valid values on that grid are
    // 6.2/6.3) — the width input's old `step={0.1}` made the whole form
    // fail native HTML5 validation, silently blocking every submit.
    const onApply = vi.fn();
    const { getByRole } = render(
      <TablePropertiesDialog
        seed={{ widthTwips: 9000, alignment: null, bordersOn: true }}
        onApply={onApply}
        onCancel={vi.fn()}
      />,
    );

    fireEvent.click(getByRole('button', { name: 'Apply' }));

    expect(onApply).toHaveBeenCalledTimes(1);
    expect(onApply).toHaveBeenCalledWith(
      expect.objectContaining({ kind: 'set-table-properties', bordersOn: true }),
    );
  });
});

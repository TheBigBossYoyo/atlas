import { useState } from 'react';

import type { ToolbarCommand } from './toolbarTypes';
import { useTranslate } from '../../../i18n';

const TWIPS_PER_INCH = 1440;
const DEFAULT_WIDTH_INCHES = 6;

function twipsToInches(twips: number): number {
  return Math.round((twips / TWIPS_PER_INCH) * 100) / 100;
}

function inchesToTwips(inches: number): number {
  return Math.round(inches * TWIPS_PER_INCH);
}

export type TablePropertiesSeed = {
  readonly widthTwips: number | null;
  readonly alignment: 'left' | 'center' | 'right' | null;
  readonly bordersOn: boolean;
};

/**
 * DXE-14 — table properties dialog basics (width/alignment/borders),
 * shared between the toolbar's "Table" dropdown and the DOCX editor's
 * right-click context menu. Width is edited in inches (friendlier than
 * raw twips for a dialog) and converted at the boundary; unchecking
 * "Width" submits `widthTwips: null` (Word's own "Automatic" table width)
 * rather than a stale numeric value.
 */
export const TablePropertiesDialog = ({
  seed,
  onApply,
  onCancel,
}: {
  seed: TablePropertiesSeed;
  onApply: (cmd: Extract<ToolbarCommand, { kind: 'set-table-properties' }>) => void;
  onCancel: () => void;
}) => {
  const t = useTranslate();
  const [widthEnabled, setWidthEnabled] = useState(seed.widthTwips !== null);
  const [widthInches, setWidthInches] = useState(
    seed.widthTwips !== null ? twipsToInches(seed.widthTwips) : DEFAULT_WIDTH_INCHES,
  );
  const [alignment, setAlignment] = useState<'left' | 'center' | 'right'>(seed.alignment ?? 'left');
  const [bordersOn, setBordersOn] = useState(seed.bordersOn);

  const ALIGNMENT_LABELS = {
    left: t('docx.tableProperties.alignLeft'),
    center: t('docx.tableProperties.alignCenter'),
    right: t('docx.tableProperties.alignRight'),
  } as const;

  return (
    <form
      className="docx-toolbar__table-props-form"
      aria-label={t('docx.tableProperties.ariaLabel')}
      onSubmit={(event) => {
        event.preventDefault();
        onApply({
          kind: 'set-table-properties',
          widthTwips: widthEnabled ? inchesToTwips(widthInches) : null,
          alignment,
          bordersOn,
        });
      }}
    >
      <div className="docx-toolbar__form-row">
        <label>
          <input
            type="checkbox"
            checked={widthEnabled}
            onChange={(event) => setWidthEnabled(event.target.checked)}
          />
          {t('docx.tableProperties.widthLabel')}
        </label>
        <input
          type="number"
          min={0.1}
          step={0.1}
          disabled={!widthEnabled}
          value={widthInches}
          aria-label={t('docx.tableProperties.widthAria')}
          onChange={(event) => setWidthInches(Number(event.target.value))}
        />
      </div>
      <fieldset className="docx-toolbar__form-row">
        <legend className="docx-toolbar__visually-hidden">{t('docx.tableProperties.alignmentLegend')}</legend>
        {(['left', 'center', 'right'] as const).map((option) => (
          <label key={option}>
            <input
              type="radio"
              name="docx-table-align"
              checked={alignment === option}
              onChange={() => setAlignment(option)}
            />
            {ALIGNMENT_LABELS[option]}
          </label>
        ))}
      </fieldset>
      <div className="docx-toolbar__form-row">
        <label>
          <input type="checkbox" checked={bordersOn} onChange={(event) => setBordersOn(event.target.checked)} />
          {t('docx.tableProperties.showBorders')}
        </label>
      </div>
      <div className="docx-toolbar__form-row docx-toolbar__form-actions">
        <button type="button" onClick={onCancel}>
          {t('docx.tableProperties.cancel')}
        </button>
        <button type="submit">{t('docx.tableProperties.apply')}</button>
      </div>
    </form>
  );
};

import { describe, it, expect } from 'vitest';
import { runStyleToCss, paraStyleToCss, tableStyleToCss, cellStyleToCss, borderToCss } from '../style';
import { twip, halfPoint, eighthPoint, hexColor, pct } from '../../model/styles';

describe('runStyleToCss', () => {
  it('handles bold', () => {
    expect(runStyleToCss({ bold: true }).fontWeight).toBe('bold');
  });

  it('handles italic', () => {
    expect(runStyleToCss({ italic: true }).fontStyle).toBe('italic');
  });

  it('handles strike', () => {
    expect(runStyleToCss({ strike: true }).textDecoration).toBe('line-through');
  });

  it('handles underline single', () => {
    const css = runStyleToCss({ underline: { style: 'single' } });
    expect(css.textDecoration).toBe('underline');
  });

  it('handles underline double', () => {
    const css = runStyleToCss({ underline: { style: 'double' } });
    expect(css.textDecoration).toBe('underline');
    expect(css.textDecorationStyle).toBe('double');
  });

  it('handles color', () => {
    expect(runStyleToCss({ color: hexColor('FF0000') }).color).toBe('#FF0000');
  });

  it('handles highlight', () => {
    expect(runStyleToCss({ highlight: 'yellow' }).backgroundColor).toBe('yellow');
  });

  it('handles font size half points', () => {
    expect(runStyleToCss({ sz: halfPoint(24) }).fontSize).toBe('12pt');
  });

  it('emits font-family stack with substitute and generic fallback (Arial → Arimo)', () => {
    expect(runStyleToCss({ rFonts: { ascii: 'Arial' } }).fontFamily).toBe('Arial, Arimo, sans-serif');
  });

  it('quotes multi-word family names and appends substitute (Times New Roman → Tinos)', () => {
    expect(runStyleToCss({ rFonts: { ascii: 'Times New Roman' } }).fontFamily).toBe('"Times New Roman", Tinos, serif');
  });

  it('uses Calibri → Carlito with sans-serif fallback', () => {
    expect(runStyleToCss({ rFonts: { ascii: 'Calibri' } }).fontFamily).toBe('Calibri, Carlito, sans-serif');
  });

  it('falls back to hAnsi when ascii is missing', () => {
    expect(runStyleToCss({ rFonts: { hAnsi: 'Cambria' } }).fontFamily).toBe('Cambria, Caladea, serif');
  });

  it('uses monospace generic for Courier-family fonts', () => {
    expect(runStyleToCss({ rFonts: { ascii: 'Courier New' } }).fontFamily).toBe('"Courier New", "Courier Prime", monospace');
  });

  it('passes through unknown families with sans-serif fallback', () => {
    expect(runStyleToCss({ rFonts: { ascii: 'Helvetica' } }).fontFamily).toBe('Helvetica, sans-serif');
  });

  it('handles spacing (twips)', () => {
    expect(runStyleToCss({ spacing: twip(20) }).letterSpacing).toBe('1pt');
  });
});

describe('paraStyleToCss', () => {
  it('handles jc both (justify)', () => {
    expect(paraStyleToCss({ jc: 'both' }).textAlign).toBe('justify');
  });

  it('handles line spacing exact', () => {
    const css = paraStyleToCss({ spacing: { line: twip(240), lineRule: 'exact' } });
    expect(css.lineHeight).toBe('12pt');
  });

  it('handles line spacing auto', () => {
    const css = paraStyleToCss({ spacing: { line: twip(360), lineRule: 'auto' } });
    expect(css.lineHeight).toBe(1.5);
  });

  it('handles indent firstLine', () => {
    const css = paraStyleToCss({ ind: { firstLine: twip(720) } });
    expect(css.textIndent).toBe('36pt');
  });

  it('handles indent hanging', () => {
    const css = paraStyleToCss({ ind: { hanging: twip(720) } });
    expect(css.textIndent).toBe('-36pt');
    expect(css.paddingLeft).toBe('36pt');
  });
});

describe('tableStyleToCss', () => {
  it('handles center justify', () => {
    const css = tableStyleToCss({ jc: 'center' });
    expect(css.marginLeft).toBe('auto');
    expect(css.marginRight).toBe('auto');
  });

  it('handles percent width', () => {
    const css = tableStyleToCss({ tblW: { type: 'pct', value: pct(5000) } });
    expect(css.width).toBe('100%');
  });
});

describe('cellStyleToCss', () => {
  it('handles dxa width', () => {
    const css = cellStyleToCss({ tcW: { type: 'dxa', value: twip(1440) } });
    expect(css.width).toBe('72pt');
  });

  it('handles vAlign', () => {
    expect(cellStyleToCss({ vAlign: 'center' }).verticalAlign).toBe('middle');
  });
});

describe('borderToCss', () => {
  it('formats dashed border', () => {
    expect(borderToCss({ style: 'dashed', color: hexColor('FF0000'), size: eighthPoint(4) })).toBe('0.5pt dashed #FF0000');
  });
});

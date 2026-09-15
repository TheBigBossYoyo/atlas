import { describe, expect, it } from 'vitest';
import { buildTabularBodyHtml } from '../tabularHtml';

describe('buildTabularBodyHtml', () => {
  it('renders the first row as <thead><th> (repeats per printed page — standard browser behavior)', () => {
    const html = buildTabularBodyHtml([{ name: 'Sheet1', rows: [['Name', 'Age'], ['Ada', '36']] }]);
    expect(html).toContain('<thead><tr><th>Name</th><th>Age</th></tr></thead>');
    expect(html).toContain('<tbody><tr><td>Ada</td><td>36</td></tr></tbody>');
  });

  it('applies rowspan/colspan for a merge and omits the cells it covers', () => {
    const html = buildTabularBodyHtml([
      {
        name: 'Sheet1',
        rows: [
          ['Title', '', 'X'],
          ['A', 'B', 'C'],
        ],
        merges: [{ r0: 0, c0: 0, r1: 0, c1: 1 }],
      },
    ]);
    // Merged top-left cell carries the span; the covered cell (row 0, col 1) is omitted entirely.
    expect(html).toContain('<th rowspan="1" colspan="2">Title</th><th>X</th>');
  });

  it('applies a rowspan merge across multiple rows and omits every covered cell', () => {
    const html = buildTabularBodyHtml([
      {
        name: 'Sheet1',
        rows: [
          ['Header'],
          ['Merged'],
          ['Covered'],
        ],
        merges: [{ r0: 1, c0: 0, r1: 2, c1: 0 }],
      },
    ]);
    expect(html).toContain('<td rowspan="2" colspan="1">Merged</td>');
    expect(html).not.toContain('Covered</td>');
  });

  it('renders a placeholder for an empty sheet instead of an empty <table>', () => {
    const html = buildTabularBodyHtml([{ name: 'Empty', rows: [] }]);
    expect(html).toContain('(empty sheet)');
    expect(html).not.toContain('<table>');
  });

  it('escapes HTML-significant characters in both sheet names and cell content', () => {
    const html = buildTabularBodyHtml([{ name: '<script>', rows: [['<img onerror=x>']] }]);
    expect(html).not.toContain('<script>');
    expect(html).not.toContain('<img onerror=x>');
    expect(html).toContain('&lt;script&gt;');
  });

  it('renders one <section class="export-sheet"> per sheet, in order', () => {
    const html = buildTabularBodyHtml([
      { name: 'First', rows: [['a']] },
      { name: 'Second', rows: [['b']] },
    ]);
    const firstIndex = html.indexOf('First');
    const secondIndex = html.indexOf('Second');
    expect(firstIndex).toBeGreaterThanOrEqual(0);
    expect(secondIndex).toBeGreaterThan(firstIndex);
    expect((html.match(/class="export-sheet"/g) ?? []).length).toBe(2);
  });
});

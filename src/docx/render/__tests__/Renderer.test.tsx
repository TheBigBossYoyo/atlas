import { describe, it, expect } from 'vitest';
import { render } from '@testing-library/react';
import { DocxRenderer } from '../Renderer';
import type { Document, Section, Paragraph, Table, Run } from '../../model/document';

describe('DocxRenderer', () => {
  const emptyDoc: Document = {
    kind: 'document',
    sections: [],
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  };

  it('renders an empty document', () => {
    const { container } = render(<DocxRenderer document={emptyDoc} />);
    expect(container.querySelector('.docx-document')).toBeInTheDocument();
  });

  it('renders a single paragraph with bold run', () => {
    const run: Run = {
      kind: 'run',
      props: { bold: true },
      children: [{ kind: 'text', value: 'Hello bold' }],
    };
    const para: Paragraph = {
      kind: 'paragraph',
      children: [run],
    };
    const doc = {
      ...emptyDoc,
      sections: [{ kind: 'section', props: {}, blocks: [para] } as Section],
    };

    const { container } = render(<DocxRenderer document={doc} />);
    const runEl = container.querySelector('.docx-run');
    expect(runEl).toHaveStyle({ fontWeight: 'bold' });
    expect(runEl).toHaveTextContent('Hello bold');
  });

  it('renders a nested table', () => {
    const table: Table = {
      kind: 'table',
      rows: [
        {
          kind: 'table-row',
          cells: [
            { kind: 'table-cell', blocks: [{ kind: 'paragraph', children: [{ kind: 'run', children: [{ kind: 'text', value: 'Cell 1' }] }] }] }
          ]
        }
      ]
    };
    const doc = {
      ...emptyDoc,
      sections: [{ kind: 'section', props: {}, blocks: [table] } as Section],
    };

    const { container } = render(<DocxRenderer document={doc} />);
    const td = container.querySelector('td.docx-table-cell');
    expect(td).toBeInTheDocument();
    expect(td).toHaveTextContent('Cell 1');
  });

  it('renders hyperlink', () => {
    const para: Paragraph = {
      kind: 'paragraph',
      children: [{
        kind: 'hyperlink',
        children: [{ kind: 'run', children: [{ kind: 'text', value: 'Link text' }] }]
      }],
    };
    const doc = {
      ...emptyDoc,
      sections: [{ kind: 'section', props: {}, blocks: [para] } as Section],
    };

    const { container } = render(<DocxRenderer document={doc} />);
    const a = container.querySelector('a.docx-hyperlink');
    expect(a).toBeInTheDocument();
    expect(a).toHaveTextContent('Link text');
    expect(a).toHaveAttribute('target', '_blank');
  });

  it('renders invisible bookmark', () => {
    const para: Paragraph = {
      kind: 'paragraph',
      children: [{ kind: 'bookmark', id: '1', boundary: 'start', name: 'bm1' }],
    };
    const doc = {
      ...emptyDoc,
      sections: [{ kind: 'section', props: {}, blocks: [para] } as Section],
    };

    const { container } = render(<DocxRenderer document={doc} />);
    const a = container.querySelector('a.docx-bookmark');
    expect(a).toBeInTheDocument();
    expect(a).toHaveAttribute('id', 'bm1');
    expect(a).toHaveStyle({ display: 'none' });
  });

  it('renders page break', () => {
    const para: Paragraph = {
      kind: 'paragraph',
      children: [{ kind: 'run', children: [{ kind: 'break', breakType: 'page' }] }],
    };
    const doc = {
      ...emptyDoc,
      sections: [{ kind: 'section', props: {}, blocks: [para] } as Section],
    };

    const { container } = render(<DocxRenderer document={doc} />);
    const br = container.querySelector('br.docx-break--page');
    expect(br).toBeInTheDocument();
    expect(br).toHaveStyle({ breakBefore: 'page' });
  });

  it('renders drawing placeholder', () => {
    const para: Paragraph = {
      kind: 'paragraph',
      children: [{ kind: 'run', children: [{ kind: 'drawing', layout: 'inline' }] }],
    };
    const doc = {
      ...emptyDoc,
      sections: [{ kind: 'section', props: {}, blocks: [para] } as Section],
    };

    const { container } = render(<DocxRenderer document={doc} />);
    const figure = container.querySelector('figure.docx-drawing--placeholder');
    expect(figure).toBeInTheDocument();
    expect(figure).toHaveTextContent('[Drawing]');
  });
  
  it('renders paragraph marker for lists', () => {
    const numbering = new Map();
    numbering.set('num1', {
      numId: 'num1',
      levels: new Map([[0, { level: 0, text: { value: '%1.', placeholders: [1] }, start: 1 }]])
    });

    const doc = {
      ...emptyDoc,
      numbering,
      sections: [{
        kind: 'section',
        props: {},
        blocks: [{
          kind: 'paragraph',
          props: { numPr: { numId: 'num1', ilvl: 0 } },
          children: []
        }]
      } as Section],
    };

    const { container } = render(<DocxRenderer document={doc} />);
    const marker = container.querySelector('.docx-list-marker');
    expect(marker).toBeInTheDocument();
    expect(marker).toHaveTextContent('1.');
  });
});

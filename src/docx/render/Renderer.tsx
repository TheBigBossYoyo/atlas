import type { Document as DocxDocument, Section, Block } from '../model/document';
import { ParagraphNode } from './paragraph';
import { TableNode } from './table';

export function DocxRenderer({ document }: { document: DocxDocument }) {
  return (
    <div className="docx-document">
      {document.sections.map((section, i) => (
        <SectionNode key={i} section={section} doc={document} />
      ))}
    </div>
  );
}

function SectionNode({ section, doc }: { section: Section; doc: DocxDocument }) {
  // A.6 just renders blocks in flow.
  return (
    <section className="docx-section">
      {section.blocks.map((block, i) => (
        <BlockNode key={i} block={block} doc={doc} />
      ))}
    </section>
  );
}

function BlockNode({ block, doc }: { block: Block; doc: DocxDocument }) {
  if (block.kind === 'paragraph') {
    return <ParagraphNode paragraph={block} doc={doc} />;
  }
  if (block.kind === 'table') {
    return <TableNode table={block} doc={doc} />;
  }
  return null;
}

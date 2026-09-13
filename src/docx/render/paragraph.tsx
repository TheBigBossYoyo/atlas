import React, { useMemo } from 'react';
import type { Paragraph, Document as DocxDocument, ParagraphChild } from '../model/document';
import type { Theme } from '../parser/theme';
import { resolveParaProps } from '../parser/cascade';
import { paraStyleToCss } from './style';
import { RunNode } from './run';

function getMarkerText(numPr: { numId?: string; ilvl?: number }, doc: DocxDocument): string {
  if (!numPr.numId || numPr.ilvl === undefined) return '• ';
  const def = doc.numbering.get(numPr.numId);
  if (!def) return '• ';
  const lvlDef = def.levelOverrides?.get(numPr.ilvl)?.levelDefinition ?? def.levels.get(numPr.ilvl);
  if (!lvlDef || !lvlDef.text) return '• ';
  return lvlDef.text.value.replace(/%\d/g, () => {
    return (lvlDef.start ?? 1).toString();
  }) + ' ';
}

export const ParagraphNode = React.memo(function ParagraphNode({ paragraph, doc, theme }: { paragraph: Paragraph; doc: DocxDocument; theme?: Theme }) {
  const effective = useMemo(() => {
    return resolveParaProps(paragraph.props, paragraph.props?.pStyle, doc.styles, { pPr: doc.defaults?.paragraph });
  }, [paragraph.props, doc.styles, doc.defaults]);

  const css = useMemo(() => paraStyleToCss(effective), [effective]);
  
  const marker = effective.numPr ? getMarkerText(effective.numPr, doc) : null;

  return (
    <p className="docx-paragraph" style={css}>
      {marker && <span className="docx-list-marker">{marker}</span>}
      {paragraph.children.map((child, i) => (
        <ParagraphChildNode key={i} child={child} pStyle={paragraph.props?.pStyle} doc={doc} theme={theme} />
      ))}
    </p>
  );
});

function ParagraphChildNode({ child, pStyle, doc, theme }: { child: ParagraphChild; pStyle?: string; doc: DocxDocument; theme?: Theme }) {
  if (child.kind === 'run') {
    return <RunNode run={child} pStyle={pStyle} doc={doc} theme={theme} />;
  }
  if (child.kind === 'hyperlink') {
    return (
      <a href="#" className="docx-hyperlink" target="_blank" rel="noopener noreferrer">
        {child.children.map((c, i) => {
          if (c.kind === 'run') return <RunNode key={i} run={c} pStyle={pStyle} doc={doc} theme={theme} />;
          return null;
        })}
      </a>
    );
  }
  if (child.kind === 'bookmark') {
    return <a id={child.name} className="docx-bookmark" style={{ display: 'none' }} aria-hidden="true"></a>;
  }
  return null;
}

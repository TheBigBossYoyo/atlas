import React, { useMemo } from 'react';
import type { Run, Document as DocxDocument, RunChild, Drawing } from '../model/document';
import type { Theme } from '../parser/theme';
import { runStyleToCss } from './style';
import { resolveRunProps } from '../parser/cascade';
import { useMediaResolver } from './mediaContext';

const EMU_PER_POINT = 12700;

export const RunNode = React.memo(function RunNode({ run, pStyle, doc, theme }: { run: Run; pStyle?: string; doc: DocxDocument; theme?: Theme }) {
  const effective = useMemo(() => {
    return resolveRunProps(run.props, pStyle, doc.styles, { rPr: doc.defaults?.run });
  }, [run.props, pStyle, doc.styles, doc.defaults]);

  const css = useMemo(() => runStyleToCss(effective, theme), [effective, theme]);

  return (
    <span className="docx-run" style={css}>
      {run.children.map((child, i) => (
        <RunChildNode key={i} child={child} />
      ))}
    </span>
  );
});

function RunChildNode({ child }: { child: RunChild }) {
  const media = useMediaResolver();

  if (child.kind === 'text') {
    return <span style={{ whiteSpace: child.preserveSpace ? 'pre-wrap' : 'normal' }}>{child.value}</span>;
  }
  if (child.kind === 'tab') {
    return <span className="docx-tab">{'\u00a0\u00a0\u00a0\u00a0'}</span>;
  }
  if (child.kind === 'break') {
    if (child.breakType === 'page') {
      return <br className="docx-break docx-break--page" style={{ breakBefore: 'page' }} />;
    }
    return <br className="docx-break" />;
  }
  if (child.kind === 'drawing') {
    return <DrawingChild drawing={child} resolver={media} />;
  }
  return null;
}

function DrawingChild({
  drawing,
  resolver,
}: {
  drawing: Drawing;
  resolver: ReturnType<typeof useMediaResolver>;
}) {
  const url = drawing.relationshipId !== undefined ? resolver.resolve(drawing.relationshipId) : null;
  const widthPt = drawing.extent !== undefined ? drawing.extent.cx / EMU_PER_POINT : undefined;
  const heightPt = drawing.extent !== undefined ? drawing.extent.cy / EMU_PER_POINT : undefined;

  if (url === null) {
    return (
      <figure className="docx-drawing docx-drawing--placeholder">
        [Drawing]
      </figure>
    );
  }

  return (
    <img
      className="docx-drawing__image"
      src={url}
      alt={drawing.description ?? drawing.name ?? ''}
      style={{
        width: widthPt !== undefined ? `${widthPt}pt` : undefined,
        height: heightPt !== undefined ? `${heightPt}pt` : undefined,
        display: 'inline-block',
        verticalAlign: 'baseline',
      }}
    />
  );
}

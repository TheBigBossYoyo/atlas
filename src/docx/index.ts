import {
  unzipDocx,
  parseRelationships,
  parseContentTypes,
  parseDocument,
  parseStyles,
  parseNumbering,
  parseTheme,
  parseHeader,
  parseFooter,
  parseFootnotes,
  parseEndnotes,
  parseComments,
} from './parser';

import {
  writeDocumentXml,
  writeStylesXml,
  writeNumberingXml,
  writeHeaderXml,
  writeFooterXml,
  writeFootnotesXml,
  writeEndnotesXml,
  writeCommentsXml,
  writeRelationshipsXml,
  writeContentTypesXml,
  packDocx,
} from './serializer';
import type { DocxPart } from './serializer';

import type { Document } from './model/document';
import type { Theme } from './parser';
import type { Relationship, ContentTypes, NumberingPart, StylesPart } from './parser';

export type DocxBundle = {
  document: Document;
  theme?: Theme;
  relationships?: ReadonlyArray<Relationship>;
  packageRelationships?: ReadonlyArray<Relationship>;
  contentTypes?: ContentTypes;
  numberingPart?: NumberingPart;
  stylesPart?: StylesPart;
  /** Raw archive bytes for passthrough of unhandled parts (theme, settings, fontTable, media, customXml, etc.). */
  rawArchive?: ReadonlyMap<string, Uint8Array>;
};

function getXml(files: ReadonlyMap<string, Uint8Array>, path: string): string | undefined {
  const buf = files.get(path);
  if (!buf) return undefined;
  return new TextDecoder().decode(buf);
}

export async function loadDocx(buffer: ArrayBuffer): Promise<DocxBundle> {
  const archive = await unzipDocx(buffer);

  const relsXml = getXml(archive.files, 'word/_rels/document.xml.rels');
  const rels = relsXml ? parseRelationships(relsXml) : [];

  const pkgRelsXml = getXml(archive.files, '_rels/.rels');
  const pkgRels = pkgRelsXml ? parseRelationships(pkgRelsXml) : [];

  const ctXml = getXml(archive.files, '[Content_Types].xml');
  const contentTypes = ctXml ? parseContentTypes(ctXml) : undefined;

  const docXml = getXml(archive.files, 'word/document.xml');
  if (!docXml) throw new Error('Missing word/document.xml');
  const baseDoc = parseDocument(docXml);

  const stylesXml = getXml(archive.files, 'word/styles.xml');
  const stylesPart: StylesPart = stylesXml
    ? parseStyles(stylesXml)
    : { styles: new Map(), docDefaults: {} };

  const numberingXml = getXml(archive.files, 'word/numbering.xml');
  const numberingPart = numberingXml ? parseNumbering(numberingXml) : undefined;
  const numberingMap = new Map();
  if (numberingPart) {
    for (const [numId, numInstance] of numberingPart.nums.entries()) {
      const abstractNum = numInstance.abstractNumId
        ? numberingPart.abstractNums.get(numInstance.abstractNumId)
        : undefined;
      numberingMap.set(numId, {
        numId,
        abstractNumId: numInstance.abstractNumId,
        styleLink: abstractNum?.styleLink,
        numberStyleLink: abstractNum?.numberStyleLink,
        levels: abstractNum?.levels || new Map(),
        levelOverrides: numInstance.levelOverrides,
      });
    }
  }

  const themeXml = getXml(archive.files, 'word/theme/theme1.xml');
  const theme = themeXml ? parseTheme(themeXml) : undefined;

  // Headers / footers / footnotes / endnotes / comments
  const headers = new Map<string, ReturnType<typeof parseHeader>>();
  const footers = new Map<string, ReturnType<typeof parseFooter>>();
  for (const rel of rels) {
    if (rel.type.endsWith('/header')) {
      const path = `word/${rel.target.replace(/^\//, '')}`;
      const xml = getXml(archive.files, path);
      if (xml) headers.set(rel.id, parseHeader(xml));
    } else if (rel.type.endsWith('/footer')) {
      const path = `word/${rel.target.replace(/^\//, '')}`;
      const xml = getXml(archive.files, path);
      if (xml) footers.set(rel.id, parseFooter(xml));
    }
  }

  const headersMap = new Map();
  for (const [id, h] of headers) {
    if (h) headersMap.set(id, { kind: 'header', id, blocks: h.blocks });
  }
  const footersMap = new Map();
  for (const [id, f] of footers) {
    if (f) footersMap.set(id, { kind: 'footer', id, blocks: f.blocks });
  }

  const footnotesXml = getXml(archive.files, 'word/footnotes.xml');
  const footnotesMap = footnotesXml ? new Map(parseFootnotes(footnotesXml)) : new Map();

  const endnotesXml = getXml(archive.files, 'word/endnotes.xml');
  const endnotesMap = endnotesXml ? new Map(parseEndnotes(endnotesXml)) : new Map();

  const commentsXml = getXml(archive.files, 'word/comments.xml');
  const commentsMap = commentsXml ? new Map(parseComments(commentsXml)) : new Map();

  const fullDoc: Document = {
    ...baseDoc,
    styles: stylesPart.styles,
    numbering: numberingMap,
    defaults: stylesPart.docDefaults
      ? {
          paragraph: stylesPart.docDefaults.pPr,
          run: stylesPart.docDefaults.rPr,
        }
      : undefined,
    comments: commentsMap,
    footnotes: footnotesMap,
    endnotes: endnotesMap,
    headers: headersMap,
    footers: footersMap,
  };

  return {
    document: fullDoc,
    theme,
    relationships: rels,
    packageRelationships: pkgRels,
    contentTypes,
    numberingPart,
    stylesPart,
    rawArchive: archive.files,
  };
}

/**
 * Serialises a DocxBundle back to a `.docx` Uint8Array.
 *
 * Strategy: start from `rawArchive` (passthrough preserves theme/settings/
 * fontTable/media/customXml exactly), then overwrite parts we own:
 *   - word/document.xml
 *   - word/styles.xml (if stylesPart present)
 *   - word/numbering.xml (if numberingPart present)
 *   - word/comments.xml (if any)
 *   - word/footnotes.xml / word/endnotes.xml (if any)
 *   - word/header*.xml / word/footer*.xml (resolved via relationships)
 *   - word/_rels/document.xml.rels (if relationships present)
 *   - [Content_Types].xml (if contentTypes present)
 */
export async function saveDocx(bundle: DocxBundle): Promise<Uint8Array> {
  const parts = new Map<string, string | Uint8Array>();

  // 1. Passthrough base from raw archive.
  if (bundle.rawArchive) {
    for (const [path, bytes] of bundle.rawArchive.entries()) {
      parts.set(path, bytes);
    }
  }

  // 2. Always write document.xml from current model.
  parts.set('word/document.xml', writeDocumentXml(bundle.document));

  // 3. Styles.
  if (bundle.stylesPart) {
    parts.set('word/styles.xml', writeStylesXml(bundle.stylesPart));
  }

  // 4. Numbering.
  if (bundle.numberingPart) {
    parts.set('word/numbering.xml', writeNumberingXml(bundle.numberingPart));
  }

  // 5. Comments.
  if (bundle.document.comments.size > 0) {
    parts.set(
      'word/comments.xml',
      writeCommentsXml(Array.from(bundle.document.comments.values())),
    );
  }

  // 6. Footnotes / endnotes.
  if (bundle.document.footnotes.size > 0) {
    parts.set(
      'word/footnotes.xml',
      writeFootnotesXml(Array.from(bundle.document.footnotes.values())),
    );
  }
  if (bundle.document.endnotes.size > 0) {
    parts.set(
      'word/endnotes.xml',
      writeEndnotesXml(Array.from(bundle.document.endnotes.values())),
    );
  }

  // 7. Headers / footers — resolve part path from relationships by id.
  if (bundle.relationships) {
    for (const rel of bundle.relationships) {
      const path = `word/${rel.target.replace(/^\//, '')}`;
      if (rel.type.endsWith('/header')) {
        const header = bundle.document.headers.get(rel.id);
        if (header) parts.set(path, writeHeaderXml(header));
      } else if (rel.type.endsWith('/footer')) {
        const footer = bundle.document.footers.get(rel.id);
        if (footer) parts.set(path, writeFooterXml(footer));
      }
    }
    parts.set(
      'word/_rels/document.xml.rels',
      writeRelationshipsXml(bundle.relationships),
    );
  }

  // 8. Package-level relationships.
  if (bundle.packageRelationships) {
    parts.set('_rels/.rels', writeRelationshipsXml(bundle.packageRelationships));
  }

  // 9. [Content_Types].xml.
  if (bundle.contentTypes) {
    parts.set('[Content_Types].xml', writeContentTypesXml(bundle.contentTypes));
  }

  // 10. Pack.
  const docxParts: DocxPart[] = [];
  for (const [path, content] of parts.entries()) {
    docxParts.push({ path, content });
  }
  return packDocx(docxParts);
}

export { DocxRenderer } from './render';

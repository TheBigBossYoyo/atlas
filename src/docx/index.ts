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
  parseCommentsExtended,
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
  writeCommentsExtendedXml,
  resolveCommentExtendedKey,
  writeRelationshipsXml,
  writeContentTypesXml,
  packDocx,
  addRelationship,
  addOverride,
  validateDocxPackage,
} from './serializer';
import type { DocxPart } from './serializer';

import type { Comment, Document } from './model/document';
import type { Theme } from './parser';
import type { Relationship, ContentTypes, NumberingPart, StylesPart } from './parser';

const COMMENTS_PART_PATH = 'word/comments.xml';
const COMMENTS_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

const COMMENTS_EXTENDED_PART_PATH = 'word/commentsExtended.xml';
const COMMENTS_EXTENDED_RELATIONSHIP_TYPE =
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const COMMENTS_EXTENDED_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';

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

  // D16 / DXS-11: a comment's resolved/done state lives in a separate part,
  // keyed by the paraId Word stamps on the comment's own first body
  // paragraph — fold it onto each Comment so the rest of Atlas never has to
  // know commentsExtended.xml exists.
  const commentsExtendedXml = getXml(archive.files, COMMENTS_EXTENDED_PART_PATH);
  const resolvedStateByKey = commentsExtendedXml ? parseCommentsExtended(commentsExtendedXml) : new Map();
  const commentsWithResolvedState = applyResolvedState(commentsMap, resolvedStateByKey);

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
    comments: commentsWithResolvedState,
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

  // Tracked locally (rather than read straight off `bundle`) so a newly
  // written comments part gets its relationship + content-type registered
  // even when the loaded document never had one (DXS-06).
  let relationships = bundle.relationships;
  let contentTypes = bundle.contentTypes;

  // 5. Comments (+ resolved-state extension part).
  //
  // Wave 1 follow-up: when the in-memory document has zero comments —
  // whether it never had any, or the user deleted the last one — remove
  // any comments.xml/commentsExtended.xml the raw-archive passthrough
  // (step 1) would otherwise carry through untouched, along with their
  // relationship + content-type entries. Previously the stale archive copy
  // passed straight through, so a document Atlas itself emptied of
  // comments still shipped a comments.xml describing comments that no
  // longer exist in the model at all.
  const commentsList = Array.from(bundle.document.comments.values());

  if (commentsList.length > 0) {
    parts.set(COMMENTS_PART_PATH, writeCommentsXml(commentsList));
    const registered = ensurePartRegistered(
      relationships,
      contentTypes,
      COMMENTS_RELATIONSHIP_TYPE,
      'comments.xml',
      COMMENTS_PART_PATH,
      COMMENTS_CONTENT_TYPE,
    );
    relationships = registered.relationships;
    contentTypes = registered.contentTypes;

    if (commentsList.some((comment) => comment.resolved !== undefined)) {
      parts.set(COMMENTS_EXTENDED_PART_PATH, writeCommentsExtendedXml(commentsList));
      const registeredExtended = ensurePartRegistered(
        relationships,
        contentTypes,
        COMMENTS_EXTENDED_RELATIONSHIP_TYPE,
        'commentsExtended.xml',
        COMMENTS_EXTENDED_PART_PATH,
        COMMENTS_EXTENDED_CONTENT_TYPE,
      );
      relationships = registeredExtended.relationships;
      contentTypes = registeredExtended.contentTypes;
    } else {
      // No comment currently carries resolved state, but a prior save's
      // commentsExtended.xml could still be sitting in the raw archive.
      const removedExtended = removePartRegistration(
        parts,
        relationships,
        contentTypes,
        COMMENTS_EXTENDED_PART_PATH,
        COMMENTS_EXTENDED_RELATIONSHIP_TYPE,
      );
      relationships = removedExtended.relationships;
      contentTypes = removedExtended.contentTypes;
    }
  } else {
    const removedComments = removePartRegistration(
      parts,
      relationships,
      contentTypes,
      COMMENTS_PART_PATH,
      COMMENTS_RELATIONSHIP_TYPE,
    );
    relationships = removedComments.relationships;
    contentTypes = removedComments.contentTypes;

    const removedExtended = removePartRegistration(
      parts,
      relationships,
      contentTypes,
      COMMENTS_EXTENDED_PART_PATH,
      COMMENTS_EXTENDED_RELATIONSHIP_TYPE,
    );
    relationships = removedExtended.relationships;
    contentTypes = removedExtended.contentTypes;
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
  if (relationships) {
    for (const rel of relationships) {
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
      writeRelationshipsXml(relationships),
    );
  }

  // 8. Package-level relationships.
  if (bundle.packageRelationships) {
    parts.set('_rels/.rels', writeRelationshipsXml(bundle.packageRelationships));
  }

  // 9. [Content_Types].xml.
  if (contentTypes) {
    parts.set('[Content_Types].xml', writeContentTypesXml(contentTypes));
  }

  // 10. Post-serialization validation (D20 / DXS-19) — catch a broken
  // package here, with a clear, specific error, rather than silently
  // writing a file Word will refuse or complain about.
  validateDocxPackage(parts);

  // 11. Pack.
  const docxParts: DocxPart[] = [];
  for (const [path, content] of parts.entries()) {
    docxParts.push({ path, content });
  }
  return packDocx(docxParts);
}

/**
 * DXS-06 (generalized): registers a part's relationship + content-type entry
 * the first time it's written to a document that never had one before —
 * without this, `saveDocx` can write a part's bytes (as it does for comments
 * and, from D16 on, commentsExtended) into an otherwise OPC-invalid package:
 * nothing declares its content type, and `document.xml` has no relationship
 * Word can use to resolve it.
 *
 * Idempotent: a document whose part was already registered (either from the
 * original load, or from an earlier call in the same save) is returned
 * unchanged.
 */
function ensurePartRegistered(
  relationships: ReadonlyArray<Relationship> | undefined,
  contentTypes: ContentTypes | undefined,
  relationshipType: string,
  relationshipTarget: string,
  partPath: string,
  contentType: string,
): { relationships: ReadonlyArray<Relationship>; contentTypes: ContentTypes } {
  const rels = relationships ?? [];
  const types = contentTypes ?? { defaults: [], overrides: [] };

  const nextRelationships = rels.some((rel) => rel.type === relationshipType)
    ? rels
    : addRelationship(rels, relationshipType, relationshipTarget).rels;

  const partName = `/${partPath}`;
  const nextContentTypes = types.overrides.some((override) => override.partName === partName)
    ? types
    : addOverride(types, partName, contentType);

  return { relationships: nextRelationships, contentTypes: nextContentTypes };
}

/**
 * Wave 1 follow-up: removes a part (if present) along with its own
 * relationship + content-type override entries. Used when the part itself
 * is being dropped from the package — e.g. every comment was deleted — so
 * the saved package never carries a relationship or content-type override
 * pointing at a part that no longer exists in it, and never ships stale
 * bytes the raw-archive passthrough (step 1) would otherwise carry through
 * untouched.
 */
function removePartRegistration(
  parts: Map<string, string | Uint8Array>,
  relationships: ReadonlyArray<Relationship> | undefined,
  contentTypes: ContentTypes | undefined,
  partPath: string,
  relationshipType: string,
): {
  relationships: ReadonlyArray<Relationship> | undefined;
  contentTypes: ContentTypes | undefined;
} {
  parts.delete(partPath);

  const nextRelationships = relationships?.filter((rel) => rel.type !== relationshipType);

  const partName = `/${partPath}`;
  const nextContentTypes = contentTypes
    ? { ...contentTypes, overrides: contentTypes.overrides.filter((o) => o.partName !== partName) }
    : contentTypes;

  return { relationships: nextRelationships, contentTypes: nextContentTypes };
}

/**
 * D16 / DXS-11: folds `commentsExtended.xml`'s resolved/done state onto each
 * Comment (see `resolveCommentExtendedKey` for the join key), so nothing
 * downstream of `loadDocx` needs to know that state lives in a separate part.
 */
function applyResolvedState(
  comments: ReadonlyMap<string, Comment>,
  resolvedStateByKey: ReadonlyMap<string, boolean>,
): ReadonlyMap<string, Comment> {
  if (resolvedStateByKey.size === 0) {
    return comments;
  }

  const result = new Map<string, Comment>();
  for (const [id, comment] of comments) {
    const resolved = resolvedStateByKey.get(resolveCommentExtendedKey(comment));
    result.set(id, resolved !== undefined ? { ...comment, resolved } : comment);
  }
  return result;
}

export { DocxRenderer } from './render';

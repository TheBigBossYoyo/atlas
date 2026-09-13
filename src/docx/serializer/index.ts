export { writeStylesXml } from './stylesWriter'
export { writeNumberingXml } from './numberingWriter'
export { writeHeaderXml } from './headerWriter'
export { writeFooterXml } from './footerWriter'
export { writeFootnotesXml, writeEndnotesXml } from './footnotesWriter'
export { writeCommentsXml } from './commentsWriter'
export {
  writeDocumentXml,
  buildParagraph,
  buildRun,
  buildTable,
  buildTableRow,
  buildTableCell,
  buildSectionProperties,
  buildRunProperties,
  buildParagraphProperties,
} from './documentWriter'

// Wave D.3 — relationships + content types
export {
  writeRelationshipsXml,
  allocateRelationshipId,
  addRelationship,
} from './relsWriter'
export {
  writeContentTypesXml,
  ensureMediaContentType,
  addOverride,
  MEDIA_CONTENT_TYPES,
} from './contentTypesWriter'
export type { ContentTypesPart } from './contentTypesWriter'

// Wave D.4 — zip packager
export { packDocx, sortPartsForWord, WORD_PART_ORDER } from './zipPackager'
export type { DocxPart } from './zipPackager'

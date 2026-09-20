// Atlas — structural validation harness for generated Office/ODF packages.
//
// Checks a .docx/.xlsx/.pptx/.odt/.odp/.ods file against the actual OPC/
// OOXML/ODF package specs, independent of Atlas's own readers/writers:
//
//  - the zip container itself (entry names, no duplicates, ODF's mimetype
//    first-and-stored rule) via `zipReader.mjs`'s own from-scratch PKZIP
//    parser, not JSZip;
//  - every XML/rels part is well-formed and free of characters XML 1.0
//    forbids;
//  - namespace prefixes used in a part are actually declared somewhere in
//    scope;
//  - OPC: `[Content_Types].xml` covers every part (Default or Override, no
//    duplicates), every `.rels` target exists and has a content type, every
//    `r:id`-shaped attribute a part uses resolves to a relationship that
//    part's own `.rels` file actually declares, required parts are present;
//  - ODF: `META-INF/manifest.xml` lists every file (and nothing that isn't
//    one), required parts are present;
//  - schema-mandated element order: `w:pPr` first in `w:p`; `w:tblPr`/
//    `w:tblGrid`/`w:tr` order in `w:tbl`; `a:bodyPr`/`a:lstStyle`/`a:p` order
//    in `p:txBody`; the OOXML `CT_Worksheet` child sequence; and the full
//    ECMA-376 child sequence of every property/container element Atlas's
//    DOCX writers emit (`w:sectPr`, `w:pPr`, `w:rPr`, `w:tblPr`, `w:trPr`,
//    `w:tcPr`, `w:numPr`, `w:style`, `w:settings`).
//
// Used both by `scripts/validate-office-file.mjs` (CLI) and by unit tests
// (`src/**/__tests__/*.test.ts`, which import this module directly — see
// `officeValidator.d.mts` for its type surface there).
import { XMLParser, XMLValidator } from 'fast-xml-parser'

import { readZipArchive } from './zipReader.mjs'

/**
 * @typedef {import('./officeValidator.d.mts').OfficeValidationIssue} Issue
 * @typedef {import('./officeValidator.d.mts').OfficeFormat} OfficeFormat
 * @typedef {import('./officeValidator.d.mts').OfficeValidationResult} OfficeValidationResult
 * @typedef {import('./zipReader.mjs').ZipEntry} ZipEntry
 */

/** The package's part-path -> raw-bytes map, as returned by `readZipArchive`. */
/** @typedef {Map<string, Buffer>} PackageFiles */

/**
 * A `fast-xml-parser` `preserveOrder: true` node: exactly one own key is the
 * tag name (mapping to its element/text children, in document order), plus
 * an optional `:@` key holding the element's attributes.
 * @typedef {{ [tag: string]: OrderedNode[] | string | number | Record<string, string | number | boolean> | undefined, ':@'?: Record<string, string | number | boolean> }} OrderedNode
 */

// ---------------------------------------------------------------------------
// Issue collection
// ---------------------------------------------------------------------------

/**
 * @param {Issue['severity']} severity
 * @param {string} code
 * @param {string | null} part
 * @param {string} message
 * @returns {Issue}
 */
function issue(severity, code, part, message) {
  return { severity, code, part, message }
}

// ---------------------------------------------------------------------------
// XML 1.0 character legality + well-formedness
// ---------------------------------------------------------------------------

// XML 1.0 Char production: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD]
// | [#x10000-#x10FFFF]. Everything else -- including the C0 controls Atlas's
// own `xmlSafeText` strips at write time -- is illegal in a well-formed XML
// document (a permissive parser may accept it; Word/Excel will not).
//
// Expressed as explicit numeric code-unit ranges (checked in a plain loop
// below), not a regex-with-escapes literal: several of these code points
// (the C0 controls, the two noncharacters at the top of the BMP) are
// themselves invisible or easily mis-typed, and this file has already once
// ended up with literal raw control bytes sitting silently inside a regex
// character class here (still ran correctly, but unreadable and fragile to
// the next edit) after a copy/paste through a layer that decodes escape
// sequences before the text reaches disk. Numeric comparisons avoid the
// whole class of problem.
const ILLEGAL_CODE_UNIT_RANGES = [
  [0x00, 0x08], // C0 controls before TAB
  [0x0b, 0x0c], // vertical tab, form feed
  [0x0e, 0x1f], // C0 controls after CR, before space
  [0xfffe, 0xffff], // noncharacters
]

/**
 * @param {number} code
 * @returns {boolean}
 */
function isIllegalCodeUnit(code) {
  return ILLEGAL_CODE_UNIT_RANGES.some(([lo, hi]) => code >= lo && code <= hi)
}

/**
 * @param {string} text
 * @returns {{ index: number, codePoint: number }[]}
 */
function findIllegalXmlChars(text) {
  const found = []
  for (let i = 0; i < text.length && found.length < 5; i++) {
    const code = text.charCodeAt(i)
    if (isIllegalCodeUnit(code)) {
      found.push({ index: i, codePoint: code })
      continue
    }
    // Unpaired surrogates are also illegal (not valid Unicode scalar values).
    const isHighSurrogate = code >= 0xd800 && code <= 0xdbff
    const isLowSurrogate = code >= 0xdc00 && code <= 0xdfff
    if (isHighSurrogate) {
      const next = text.charCodeAt(i + 1)
      if (!(next >= 0xdc00 && next <= 0xdfff)) found.push({ index: i, codePoint: code })
    } else if (isLowSurrogate) {
      const prev = text.charCodeAt(i - 1)
      if (!(prev >= 0xd800 && prev <= 0xdbff)) found.push({ index: i, codePoint: code })
    }
  }
  return found
}

const orderedXmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
})

/**
 * Returns `{ ok: true, tree }` or `{ ok: false, message }`. Never throws.
 * @param {string} xmlText
 * @returns {{ ok: true, tree: OrderedNode[] } | { ok: false, message: string }}
 */
function parseOrdered(xmlText) {
  const validation = XMLValidator.validate(xmlText, { allowBooleanAttributes: true })
  if (validation !== true) {
    return { ok: false, message: `${validation.err.msg} (line ${validation.err.line}, col ${validation.err.col})` }
  }
  try {
    return { ok: true, tree: orderedXmlParser.parse(xmlText) }
  } catch (cause) {
    return { ok: false, message: cause instanceof Error ? cause.message : String(cause) }
  }
}

/**
 * @param {OrderedNode} node
 * @returns {string | undefined}
 */
function nodeName(node) {
  for (const key of Object.keys(node)) {
    if (key !== ':@') return key
  }
  return undefined
}

/**
 * @param {OrderedNode} node
 * @returns {boolean}
 */
function isElementNode(node) {
  const name = nodeName(node)
  return name !== undefined && name !== '#text' && !name.startsWith('?')
}

/**
 * @param {OrderedNode} node
 * @returns {OrderedNode[]}
 */
function elementChildren(node) {
  const name = nodeName(node)
  const children = name === undefined ? undefined : node[name]
  return Array.isArray(children) ? children.filter(isElementNode) : []
}

/**
 * @param {OrderedNode} node
 * @returns {Record<string, string | number | boolean>}
 */
function attributesOf(node) {
  return node[':@'] ?? {}
}

/**
 * Recursively collects every element node anywhere in the tree whose tag equals `tag`.
 * @param {OrderedNode[]} nodes
 * @param {string} tag
 * @param {OrderedNode[]} [out]
 * @returns {OrderedNode[]}
 */
function collectByTag(nodes, tag, out = []) {
  for (const node of nodes) {
    if (!isElementNode(node)) continue
    const name = nodeName(node)
    if (name === undefined) continue
    if (name === tag) out.push(node)
    const children = node[name]
    if (Array.isArray(children)) collectByTag(children, tag, out)
  }
  return out
}

// ---------------------------------------------------------------------------
// Namespace-prefix-declared check
// ---------------------------------------------------------------------------

const RESERVED_PREFIXES = new Set(['xml', 'xmlns'])

/**
 * @param {Record<string, string | number | boolean>} attrs
 * @returns {string[]}
 */
function declaredPrefixesFrom(attrs) {
  const declared = []
  for (const key of Object.keys(attrs)) {
    // attributeNamePrefix is '@_'; xmlns:foo="..." -> '@_xmlns:foo'.
    const bare = key.slice(2)
    if (bare === 'xmlns') declared.push('') // default namespace
    else if (bare.startsWith('xmlns:')) declared.push(bare.slice('xmlns:'.length))
  }
  return declared
}

/**
 * @param {string} qualifiedName
 * @returns {string}
 */
function prefixOf(qualifiedName) {
  const colon = qualifiedName.indexOf(':')
  return colon === -1 ? '' : qualifiedName.slice(0, colon)
}

/**
 * @param {OrderedNode[]} nodes
 * @param {Set<string>} scope
 * @param {string} partPath
 * @param {Issue[]} issues
 * @returns {void}
 */
function checkNamespacesDeclared(nodes, scope, partPath, issues) {
  for (const node of nodes) {
    if (!isElementNode(node)) continue
    const name = nodeName(node)
    if (name === undefined) continue
    const attrs = attributesOf(node)
    const localScope = new Set([...scope, ...declaredPrefixesFrom(attrs)])

    const elementPrefix = prefixOf(name)
    if (elementPrefix !== '' && !RESERVED_PREFIXES.has(elementPrefix) && !localScope.has(elementPrefix)) {
      issues.push(
        issue('error', 'namespace-not-declared', partPath, `Element <${name}> uses undeclared namespace prefix "${elementPrefix}:".`),
      )
    }
    for (const key of Object.keys(attrs)) {
      const attrName = key.slice(2) // strip '@_'
      if (attrName === 'xmlns' || attrName.startsWith('xmlns:')) continue
      const attrPrefix = prefixOf(attrName)
      if (attrPrefix !== '' && !RESERVED_PREFIXES.has(attrPrefix) && !localScope.has(attrPrefix)) {
        issues.push(
          issue(
            'error',
            'namespace-not-declared',
            partPath,
            `Attribute "${attrName}" on <${name}> uses undeclared namespace prefix "${attrPrefix}:".`,
          ),
        )
      }
    }

    const children = node[name]
    if (Array.isArray(children)) checkNamespacesDeclared(children, localScope, partPath, issues)
  }
}

// ---------------------------------------------------------------------------
// Element-order checks
// ---------------------------------------------------------------------------

/**
 * `tag`, if it occurs among `node`'s element children at all, must be the very first one.
 * @param {OrderedNode} node
 * @param {string} tag
 * @param {string} partPath
 * @param {Issue[]} issues
 * @param {string} [schemaNote]
 * @returns {void}
 */
function assertFirstIfPresent(node, tag, partPath, issues, schemaNote) {
  const children = elementChildren(node)
  const index = children.findIndex((child) => nodeName(child) === tag)
  if (index > 0) {
    issues.push(
      issue(
        'error',
        'element-order',
        partPath,
        `<${nodeName(node)}> has <${tag}> at child position ${index}, but the schema requires it first${schemaNote ? ` (${schemaNote})` : ''}.`,
      ),
    )
  }
}

/**
 * Every element child of `node` whose tag is in `canonicalOrder` must appear
 * in that relative order (elements not in the list are ignored, so this
 * only asserts ordering among the recognized set — the right level of
 * strictness for a large closed content model like CT_Worksheet where an
 * unrecognized/future element shouldn't fail the check on its own).
 * @param {OrderedNode} node
 * @param {string[]} canonicalOrder
 * @param {string} partPath
 * @param {Issue[]} issues
 * @param {string} schemaLabel
 * @returns {void}
 */
function assertCanonicalOrder(node, canonicalOrder, partPath, issues, schemaLabel) {
  const rank = new Map(canonicalOrder.map((tag, i) => [tag, i]))
  const children = elementChildren(node)
  let lastRank = -1
  /** @type {string | null} */
  let lastTag = null
  for (const child of children) {
    const tag = nodeName(child)
    if (tag === undefined) continue
    const r = rank.get(tag)
    if (r === undefined) continue
    if (r < lastRank) {
      issues.push(
        issue(
          'error',
          'element-order',
          partPath,
          `<${nodeName(node)}> has <${tag}> after <${lastTag}>, violating the ${schemaLabel} child order.`,
        ),
      )
    }
    lastRank = Math.max(lastRank, r)
    lastTag = tag
  }
}

const CT_WORKSHEET_ORDER = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
]

const TX_BODY_ORDER = ['a:bodyPr', 'a:lstStyle', 'a:p']
const TBL_ORDER = ['w:tblPr', 'w:tblGrid', 'w:tr']

// The remaining sequences below cover every property/container element type
// Atlas's DOCX writers (`src/docx/serializer/**`) actually emit, added after
// a structural-fidelity review found several of those writers emitting
// children out of schema order (the classic cause of Word's "found
// unreadable content" repair prompt, since every one of these content models
// is `xsd:sequence`, not a bag). Each array is the *full* ECMA-376 child
// sequence for that complex type, copied from wml.xsd — not just the subset
// Atlas currently writes — so `assertCanonicalOrder`'s "ignore unrecognized
// tags" behavior stays correct if a writer starts emitting a member it
// doesn't today, and so the list itself doubles as a spec reference instead
// of needing one maintained separately.
const CT_SECT_PR_ORDER = [
  'w:headerReference',
  'w:footerReference',
  'w:footnotePr',
  'w:endnotePr',
  'w:type',
  'w:pgSz',
  'w:pgMar',
  'w:paperSrc',
  'w:pgBorders',
  'w:lnNumType',
  'w:pgNumType',
  'w:cols',
  'w:formProt',
  'w:vAlign',
  'w:noEndnote',
  'w:titlePg',
  'w:textDirection',
  'w:bidi',
  'w:rtlGutter',
  'w:docGrid',
  'w:printerSettings',
  'w:sectPrChange',
]

// CT_PPrBase (§17.3.1.26) followed by CT_PPr's own trailing sequence
// (rPr, sectPr, pPrChange).
const CT_PPR_ORDER = [
  'w:pStyle',
  'w:keepNext',
  'w:keepLines',
  'w:pageBreakBefore',
  'w:framePr',
  'w:widowControl',
  'w:numPr',
  'w:suppressLineNumbers',
  'w:pBdr',
  'w:shd',
  'w:tabs',
  'w:suppressAutoHyphens',
  'w:kinsoku',
  'w:wordWrap',
  'w:overflowPunct',
  'w:topLinePunct',
  'w:autoSpaceDE',
  'w:autoSpaceDN',
  'w:bidi',
  'w:adjustRightInd',
  'w:snapToGrid',
  'w:spacing',
  'w:ind',
  'w:contextualSpacing',
  'w:mirrorIndents',
  'w:suppressOverlap',
  'w:jc',
  'w:textDirection',
  'w:textAlignment',
  'w:textboxTightWrap',
  'w:outlineLvl',
  'w:divId',
  'w:cnfStyle',
  'w:rPr',
  'w:sectPr',
  'w:pPrChange',
]

// CT_RPr / EG_RPrBase (§17.3.2.28) followed by CT_RPr's trailing rPrChange.
const CT_RPR_ORDER = [
  'w:rStyle',
  'w:rFonts',
  'w:b',
  'w:bCs',
  'w:i',
  'w:iCs',
  'w:caps',
  'w:smallCaps',
  'w:strike',
  'w:dstrike',
  'w:outline',
  'w:shadow',
  'w:emboss',
  'w:imprint',
  'w:noProof',
  'w:snapToGrid',
  'w:vanish',
  'w:webHidden',
  'w:color',
  'w:spacing',
  'w:w',
  'w:kern',
  'w:position',
  'w:sz',
  'w:szCs',
  'w:highlight',
  'w:u',
  'w:effect',
  'w:bdr',
  'w:shd',
  'w:fitText',
  'w:vertAlign',
  'w:rtl',
  'w:cs',
  'w:em',
  'w:lang',
  'w:eastAsianLayout',
  'w:specVanish',
  'w:oMath',
  'w:rPrChange',
]

// CT_TblPrBase (§17.4.60) followed by CT_TblPr's trailing tblPrChange.
const CT_TBL_PR_ORDER = [
  'w:tblStyle',
  'w:tblpPr',
  'w:tblOverlap',
  'w:bidiVisual',
  'w:tblStyleRowBandSize',
  'w:tblStyleColBandSize',
  'w:tblW',
  'w:jc',
  'w:tblCellSpacing',
  'w:tblInd',
  'w:tblBorders',
  'w:shd',
  'w:tblLayout',
  'w:tblCellMar',
  'w:tblLook',
  'w:tblCaption',
  'w:tblDescription',
  'w:tblPrChange',
]

// CT_TrPrBase (§17.4.83) followed by CT_TrPr's trailing ins/del/trPrChange.
const CT_TR_PR_ORDER = [
  'w:cnfStyle',
  'w:divId',
  'w:gridBefore',
  'w:gridAfter',
  'w:wBefore',
  'w:wAfter',
  'w:cantSplit',
  'w:trHeight',
  'w:tblHeader',
  'w:tblCellSpacing',
  'w:jc',
  'w:hidden',
  'w:ins',
  'w:del',
  'w:trPrChange',
]

// CT_TcPrBase (§17.4.70) followed by CT_TcPr's trailing members.
const CT_TC_PR_ORDER = [
  'w:cnfStyle',
  'w:tcW',
  'w:gridSpan',
  'w:hMerge',
  'w:vMerge',
  'w:tcBorders',
  'w:shd',
  'w:noWrap',
  'w:tcMar',
  'w:textDirection',
  'w:tcFitText',
  'w:vAlign',
  'w:hideMark',
  'w:headers',
  'w:cellIns',
  'w:cellDel',
  'w:cellMerge',
  'w:tcPrChange',
]

// CT_NumPr (§17.9.11).
const CT_NUM_PR_ORDER = ['w:ilvl', 'w:numId', 'w:numberingChange', 'w:ins']

// CT_Style (§17.7.4.17).
const CT_STYLE_ORDER = [
  'w:name',
  'w:aliases',
  'w:basedOn',
  'w:next',
  'w:link',
  'w:autoRedefine',
  'w:hidden',
  'w:uiPriority',
  'w:semiHidden',
  'w:unhideWhenUsed',
  'w:qFormat',
  'w:locked',
  'w:personal',
  'w:personalCompose',
  'w:personalReply',
  'w:rsid',
  'w:pPr',
  'w:rPr',
  'w:tblPr',
  'w:trPr',
  'w:tcPr',
  'w:tblStylePr',
]

// A conservative prefix of CT_Settings (§17.15.1.32)'s child sequence:
// mirrors `src/docx/serializer/settingsWriter.ts`'s own
// `ELEMENTS_BEFORE_TRACK_CHANGES` list (the elements whose position relative
// to `w:trackChanges` that writer already reasons about when it needs to
// insert one), plus `trackChanges` itself. Intentionally not the full
// ~90-member CT_Settings sequence — Atlas only ever *writes* into this part
// via that narrow trackChanges edit, so this is the slice worth asserting on;
// `assertCanonicalOrder` ignores every element outside the list, so a
// `settings.xml` passed through from Word with other members present in
// between is unaffected.
const CT_SETTINGS_ORDER = [
  'w:writeProtection',
  'w:view',
  'w:zoom',
  'w:removePersonalInformation',
  'w:doNotDisplayPageBoundaries',
  'w:displayBackgroundShape',
  'w:embedTrueTypeFonts',
  'w:embedSystemFonts',
  'w:saveSubsetFonts',
  'w:mirrorMargins',
  'w:hideSpellingErrors',
  'w:hideGrammaticalErrors',
  'w:proofState',
  'w:attachedTemplate',
  'w:linkStyles',
  'w:documentType',
  'w:mailMerge',
  'w:revisionView',
  'w:trackChanges',
]

/**
 * @param {OrderedNode[]} rootNodes
 * @param {string} partPath
 * @param {Issue[]} issues
 * @returns {void}
 */
function checkElementOrder(rootNodes, partPath, issues) {
  for (const p of collectByTag(rootNodes, 'w:p')) {
    assertFirstIfPresent(p, 'w:pPr', partPath, issues, 'w:pPr must precede run content in w:p')
  }
  for (const tbl of collectByTag(rootNodes, 'w:tbl')) {
    const hasGrid = elementChildren(tbl).some((child) => nodeName(child) === 'w:tblGrid')
    if (!hasGrid) {
      issues.push(issue('error', 'missing-tblgrid', partPath, '<w:tbl> has no required <w:tblGrid> child.'))
    }
    assertCanonicalOrder(tbl, TBL_ORDER, partPath, issues, 'CT_Tbl')
  }
  for (const txBody of collectByTag(rootNodes, 'p:txBody')) {
    assertCanonicalOrder(txBody, TX_BODY_ORDER, partPath, issues, 'CT_TextBody')
  }
  for (const worksheet of collectByTag(rootNodes, 'worksheet')) {
    assertCanonicalOrder(worksheet, CT_WORKSHEET_ORDER, partPath, issues, 'CT_Worksheet')
  }
  for (const sectPr of collectByTag(rootNodes, 'w:sectPr')) {
    assertCanonicalOrder(sectPr, CT_SECT_PR_ORDER, partPath, issues, 'CT_SectPr')
  }
  for (const pPr of collectByTag(rootNodes, 'w:pPr')) {
    assertCanonicalOrder(pPr, CT_PPR_ORDER, partPath, issues, 'CT_PPr')
  }
  for (const rPr of collectByTag(rootNodes, 'w:rPr')) {
    assertCanonicalOrder(rPr, CT_RPR_ORDER, partPath, issues, 'CT_RPr')
  }
  for (const tblPr of collectByTag(rootNodes, 'w:tblPr')) {
    assertCanonicalOrder(tblPr, CT_TBL_PR_ORDER, partPath, issues, 'CT_TblPr')
  }
  for (const trPr of collectByTag(rootNodes, 'w:trPr')) {
    assertCanonicalOrder(trPr, CT_TR_PR_ORDER, partPath, issues, 'CT_TrPr')
  }
  for (const tcPr of collectByTag(rootNodes, 'w:tcPr')) {
    assertCanonicalOrder(tcPr, CT_TC_PR_ORDER, partPath, issues, 'CT_TcPr')
  }
  for (const numPr of collectByTag(rootNodes, 'w:numPr')) {
    assertCanonicalOrder(numPr, CT_NUM_PR_ORDER, partPath, issues, 'CT_NumPr')
  }
  for (const style of collectByTag(rootNodes, 'w:style')) {
    assertCanonicalOrder(style, CT_STYLE_ORDER, partPath, issues, 'CT_Style')
  }
  for (const settings of collectByTag(rootNodes, 'w:settings')) {
    assertCanonicalOrder(settings, CT_SETTINGS_ORDER, partPath, issues, 'CT_Settings')
  }
}

// ---------------------------------------------------------------------------
// Attribute datatype checks
// ---------------------------------------------------------------------------
//
// Every check above is a STRUCTURE check (part inventory, element order,
// well-formedness) -- none of them look at whether an attribute's *value*
// actually has the shape its declared simple type requires. Two real
// examples that slipped through before this section existed:
// `<w:color w:val="#ff0000"/>` (ST_HexColor forbids the leading "#" -- Word
// accepts only "auto" or six bare hex digits) and
// `<w:abstractNumId w:val="atlas-list-2"/>` (ST_DecimalNumber is a bare
// integer, not an arbitrary string). This section closes that gap for the
// handful of simple types below, driven by one small table --
// `ELEMENT_ATTRIBUTE_TYPES` -- of (element, attribute) -> type name, so
// adding another checked attribute later is a one-line table entry, not a
// new function.
//
// Every entry was checked against the actual WordprocessingML schema
// (wml.xsd + shared-commonSimpleTypes.xsd from ECMA-376/ISO-29500) --
// deliberately NOT exhaustive. A false positive here (rejecting a file
// that's actually spec-valid) is worse than a missed check, so an attribute
// only goes in the table once its exact type name has been confirmed
// against the schema text; anywhere the sign/unit rules were unclear for a
// specific attribute, it was left out rather than guessed at.

const HEX_COLOR_RE = /^[0-9A-Fa-f]{6}$/
const ON_OFF_VALUES = new Set(['true', 'false', '1', '0', 'on', 'off'])
// ST_UnsignedDecimalNumber (xsd:unsignedLong) and ST_DecimalNumber
// (xsd:integer) lexical spaces: an optional sign (unsigned allows only "+",
// never "-") followed by digits.
const UNSIGNED_DECIMAL_RE = /^\+?\d+$/
const DECIMAL_RE = /^[+-]?\d+$/
// ST_(Positive)UniversalMeasure: a decimal number immediately followed by a
// unit, no space, no "auto". The signed form permits a leading "-"; the
// positive/unsigned form permits no sign at all (not even "+").
const MEASURE_UNIT = '(?:mm|cm|in|pt|pc|pi)'
const POSITIVE_UNIVERSAL_MEASURE_RE = new RegExp(`^\\d+(?:\\.\\d+)?${MEASURE_UNIT}$`)
const SIGNED_UNIVERSAL_MEASURE_RE = new RegExp(`^-?\\d+(?:\\.\\d+)?${MEASURE_UNIT}$`)

/** @type {Record<string, (value: string) => boolean>} */
const SIMPLE_TYPE_VALIDATORS = {
  // ST_HexColor = union(ST_HexColorAuto, ST_HexColorRGB): the literal "auto"
  // or exactly six hex digits (ST_HexColorRGB is xsd:hexBinary, length 3
  // bytes) -- never a leading "#".
  ST_HexColor: (v) => v === 'auto' || HEX_COLOR_RE.test(v),
  // ST_DecimalNumber is xsd:integer everywhere it's used in wml.xsd -- a
  // bare, optionally-signed integer.
  ST_DecimalNumber: (v) => DECIMAL_RE.test(v),
  // ST_OnOff = union(xsd:boolean, ST_OnOff1): xsd:boolean's lexical set
  // (true/false/1/0) plus the literal on/off enumeration. An ABSENT
  // attribute is not checked here at all (see `checkAttributeDatatypes`) --
  // for every CT_OnOff element Atlas or Word emits, omitting `w:val`
  // entirely means "on", which is a presence rule, not a value-format one.
  ST_OnOff: (v) => ON_OFF_VALUES.has(v),
  // ST_TwipsMeasure = union(ST_UnsignedDecimalNumber, ST_PositiveUniversalMeasure):
  // non-negative twips, or a positive measurement with a unit suffix --
  // never signed.
  ST_TwipsMeasure: (v) => UNSIGNED_DECIMAL_RE.test(v) || POSITIVE_UNIVERSAL_MEASURE_RE.test(v),
  // ST_SignedTwipsMeasure = union(xsd:integer, ST_UniversalMeasure): twips
  // that may be negative, or a measurement with a unit suffix that may also
  // be negative.
  ST_SignedTwipsMeasure: (v) => DECIMAL_RE.test(v) || SIGNED_UNIVERSAL_MEASURE_RE.test(v),
}

/** @type {Record<keyof typeof SIMPLE_TYPE_VALIDATORS, string>} */
const SIMPLE_TYPE_DESCRIPTIONS = {
  ST_HexColor: '"auto" or exactly six hex digits (e.g. "FF0000"), never a leading "#"',
  ST_DecimalNumber: 'an integer (e.g. "3" or "-1")',
  ST_OnOff: 'one of "true", "false", "1", "0", "on", "off"',
  ST_TwipsMeasure: 'a non-negative integer in twips, or a positive measurement with a unit, e.g. "720" or "0.5in"',
  ST_SignedTwipsMeasure: 'an integer in twips (may be negative), or a measurement with a unit, e.g. "-720" or "-0.5in"',
}

// Border-position elements shared by CT_PBdr / CT_TblBorders / CT_TcBorders
// (ECMA-376 §17.4.4 / §17.4.38 / §17.4.66) -- each is a CT_Border, so each
// has its own `w:color` (ST_HexColor) and `w:shadow`/`w:frame` (ST_OnOff)
// attributes.
const BORDER_ELEMENTS = ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:start', 'w:end', 'w:between', 'w:bar', 'w:insideH', 'w:insideV']

// CT_OnOff-typed elements' `w:val` -- grouped by the schema group/type each
// was confirmed against (see the module comment above: an element is only
// listed here once its type was actually checked, never assumed from a
// name that merely "sounds" boolean).
const ON_OFF_ELEMENTS = [
  // EG_RPrBase (run properties) -- ECMA-376 §17.3.2.28.
  'w:b',
  'w:bCs',
  'w:i',
  'w:iCs',
  'w:caps',
  'w:smallCaps',
  'w:strike',
  'w:dstrike',
  'w:outline',
  'w:shadow',
  'w:emboss',
  'w:imprint',
  'w:noProof',
  'w:snapToGrid',
  'w:vanish',
  'w:webHidden',
  'w:rtl',
  'w:cs',
  'w:specVanish',
  'w:oMath',
  // CT_PPrBase (paragraph properties) -- §17.3.1.26.
  'w:keepNext',
  'w:keepLines',
  'w:pageBreakBefore',
  'w:widowControl',
  'w:suppressLineNumbers',
  'w:suppressAutoHyphens',
  'w:kinsoku',
  'w:wordWrap',
  'w:overflowPunct',
  'w:topLinePunct',
  'w:autoSpaceDE',
  'w:autoSpaceDN',
  'w:bidi',
  'w:adjustRightInd',
  'w:contextualSpacing',
  'w:mirrorIndents',
  'w:suppressOverlap',
  // CT_TblPrBase -- §17.4.60.
  'w:bidiVisual',
  // CT_TrPrBase -- §17.4.83.
  'w:cantSplit',
  'w:tblHeader',
  'w:hidden',
  // CT_TcPrBase -- §17.4.70.
  'w:noWrap',
  'w:hideMark',
  // CT_Style -- §17.7.4.17.
  'w:autoRedefine',
  'w:semiHidden',
  'w:unhideWhenUsed',
  'w:qFormat',
  'w:locked',
  'w:personal',
  'w:personalCompose',
  'w:personalReply',
  // The CT_Settings slice this file's element-order check already tracks
  // (§17.15.1.32).
  'w:trackChanges',
  'w:removePersonalInformation',
  'w:doNotDisplayPageBoundaries',
  'w:displayBackgroundShape',
  'w:embedTrueTypeFonts',
  'w:embedSystemFonts',
  'w:saveSubsetFonts',
  'w:mirrorMargins',
  'w:hideSpellingErrors',
  'w:hideGrammaticalErrors',
  'w:linkStyles',
]

/**
 * `(element, attribute) -> ST_* type name`. See the module comment above
 * this section for how conservatively this is populated -- this is the
 * "small, obvious" extension point: a new checked attribute is a new entry
 * here (plus, if it's a genuinely new type, a validator/description pair
 * above), never a new special-cased function.
 * @type {{ element: string, attribute: string, type: keyof typeof SIMPLE_TYPE_VALIDATORS }[]}
 */
const ELEMENT_ATTRIBUTE_TYPES = [
  // --- ST_HexColor ---------------------------------------------------------
  { element: 'w:color', attribute: 'w:val', type: 'ST_HexColor' }, // CT_Color (rPr run color, etc.)
  { element: 'w:shd', attribute: 'w:fill', type: 'ST_HexColor' }, // CT_Shd
  { element: 'w:shd', attribute: 'w:color', type: 'ST_HexColor' }, // CT_Shd
  { element: 'w:u', attribute: 'w:color', type: 'ST_HexColor' }, // CT_Underline
  ...BORDER_ELEMENTS.map((element) => ({ element, attribute: 'w:color', type: 'ST_HexColor' })),

  // --- ST_OnOff --------------------------------------------------------------
  ...ON_OFF_ELEMENTS.map((element) => ({ element, attribute: 'w:val', type: 'ST_OnOff' })),
  ...BORDER_ELEMENTS.map((element) => ({ element, attribute: 'w:shadow', type: 'ST_OnOff' })), // CT_Border
  ...BORDER_ELEMENTS.map((element) => ({ element, attribute: 'w:frame', type: 'ST_OnOff' })), // CT_Border

  // --- ST_DecimalNumber -------------------------------------------------------
  // CT_DecimalNumber's own `val`, plus the raw ST_DecimalNumber-typed
  // attributes wml.xsd defines directly on their owning element.
  { element: 'w:abstractNum', attribute: 'w:abstractNumId', type: 'ST_DecimalNumber' },
  { element: 'w:num', attribute: 'w:numId', type: 'ST_DecimalNumber' },
  { element: 'w:abstractNumId', attribute: 'w:val', type: 'ST_DecimalNumber' }, // child of w:num
  { element: 'w:numId', attribute: 'w:val', type: 'ST_DecimalNumber' }, // child of w:numPr
  { element: 'w:ilvl', attribute: 'w:val', type: 'ST_DecimalNumber' }, // child of w:numPr
  { element: 'w:lvl', attribute: 'w:ilvl', type: 'ST_DecimalNumber' },
  { element: 'w:lvlOverride', attribute: 'w:ilvl', type: 'ST_DecimalNumber' },
  { element: 'w:startOverride', attribute: 'w:val', type: 'ST_DecimalNumber' },
  { element: 'w:gridSpan', attribute: 'w:val', type: 'ST_DecimalNumber' },
  { element: 'w:uiPriority', attribute: 'w:val', type: 'ST_DecimalNumber' },
  { element: 'w:tblStyleColBandSize', attribute: 'w:val', type: 'ST_DecimalNumber' },
  { element: 'w:tblStyleRowBandSize', attribute: 'w:val', type: 'ST_DecimalNumber' },
  { element: 'w:outlineLvl', attribute: 'w:val', type: 'ST_DecimalNumber' },
  { element: 'w:divId', attribute: 'w:val', type: 'ST_DecimalNumber' },

  // --- ST_TwipsMeasure / ST_SignedTwipsMeasure --------------------------------
  // CT_Spacing (§17.3.1.33) and CT_Ind (§17.3.1.12), confirmed
  // attribute-by-attribute: before/after/hanging/firstLine never go
  // negative; line/left/right/start/end do (RTL and negative-outdent use,
  // respectively). CT_SignedTwipsMeasure (rPr character spacing, §17.3.2.32)
  // is the *type itself* -- its single `val` attribute is signed.
  { element: 'w:spacing', attribute: 'w:before', type: 'ST_TwipsMeasure' },
  { element: 'w:spacing', attribute: 'w:after', type: 'ST_TwipsMeasure' },
  { element: 'w:spacing', attribute: 'w:line', type: 'ST_SignedTwipsMeasure' },
  { element: 'w:spacing', attribute: 'w:val', type: 'ST_SignedTwipsMeasure' }, // rPr character spacing (CT_SignedTwipsMeasure)
  { element: 'w:ind', attribute: 'w:hanging', type: 'ST_TwipsMeasure' },
  { element: 'w:ind', attribute: 'w:firstLine', type: 'ST_TwipsMeasure' },
  { element: 'w:ind', attribute: 'w:left', type: 'ST_SignedTwipsMeasure' },
  { element: 'w:ind', attribute: 'w:right', type: 'ST_SignedTwipsMeasure' },
  { element: 'w:ind', attribute: 'w:start', type: 'ST_SignedTwipsMeasure' },
  { element: 'w:ind', attribute: 'w:end', type: 'ST_SignedTwipsMeasure' },
]

/** @type {Map<string, Map<string, keyof typeof SIMPLE_TYPE_VALIDATORS>>} */
const ELEMENT_ATTRIBUTE_TYPE_INDEX = new Map()
for (const { element, attribute, type } of ELEMENT_ATTRIBUTE_TYPES) {
  let byAttribute = ELEMENT_ATTRIBUTE_TYPE_INDEX.get(element)
  if (!byAttribute) {
    byAttribute = new Map()
    ELEMENT_ATTRIBUTE_TYPE_INDEX.set(element, byAttribute)
  }
  byAttribute.set(attribute, type)
}

/**
 * Recursively checks every element in the tree against
 * `ELEMENT_ATTRIBUTE_TYPE_INDEX`: for each (element, attribute) pair the
 * table knows about, and that's actually PRESENT on that element (an absent
 * attribute is never wrong here -- e.g. a CT_OnOff element with no `w:val`
 * at all means "on", a presence rule the table doesn't need to special-case),
 * the value must match that simple type's lexical rules.
 * @param {OrderedNode[]} nodes
 * @param {string} partPath
 * @param {Issue[]} issues
 * @returns {void}
 */
function checkAttributeDatatypes(nodes, partPath, issues) {
  for (const node of nodes) {
    if (!isElementNode(node)) continue
    const name = nodeName(node)
    if (name === undefined) continue

    const byAttribute = ELEMENT_ATTRIBUTE_TYPE_INDEX.get(name)
    if (byAttribute) {
      const attrs = attributesOf(node)
      for (const [attribute, type] of byAttribute) {
        const key = `@_${attribute}`
        if (!(key in attrs)) continue
        const value = String(attrs[key])
        if (!SIMPLE_TYPE_VALIDATORS[type](value)) {
          issues.push(
            issue(
              'error',
              'invalid-attribute-value',
              partPath,
              `<${name}> attribute "${attribute}"="${value}" is not a valid ${type} (must be ${SIMPLE_TYPE_DESCRIPTIONS[type]}).`,
            ),
          )
        }
      }
    }

    const children = node[name]
    if (Array.isArray(children)) checkAttributeDatatypes(children, partPath, issues)
  }
}

// ---------------------------------------------------------------------------
// Zip container checks
// ---------------------------------------------------------------------------

/**
 * @param {ZipEntry[]} entries
 * @param {Issue[]} issues
 * @returns {void}
 */
function checkZipEntryNames(entries, issues) {
  const seen = new Map()
  for (const entry of entries) {
    if (entry.name.startsWith('/')) {
      issues.push(issue('error', 'zip-leading-slash', entry.name, 'Zip entry name starts with "/".'))
    }
    if (entry.name.includes('\\')) {
      issues.push(issue('error', 'zip-backslash', entry.name, 'Zip entry name contains a backslash; OPC/ODF require forward slashes.'))
    }
    if (entry.name.split('/').includes('..')) {
      issues.push(issue('error', 'zip-path-traversal', entry.name, 'Zip entry name contains a ".." path segment.'))
    }
    if (!entry.isDirectory) {
      seen.set(entry.name, (seen.get(entry.name) ?? 0) + 1)
    }
  }
  for (const [name, count] of seen) {
    if (count > 1) issues.push(issue('error', 'zip-duplicate-entry', name, `Zip contains ${count} entries with the same name.`))
  }
}

/**
 * @param {ZipEntry[]} entries
 * @param {Issue[]} issues
 * @returns {void}
 */
function checkOdfMimetypeEntry(entries, issues) {
  const mimetype = entries.find((e) => e.name === 'mimetype')
  if (!mimetype) return // caller only invokes this once ODF-ness is already established
  const first = entries.find((e) => !e.isDirectory)
  if (first?.name !== 'mimetype') {
    issues.push(issue('error', 'odf-mimetype-not-first', 'mimetype', `The "mimetype" entry must be the first file in the zip; found "${first?.name}" first.`))
  }
  if (mimetype.method !== 0) {
    issues.push(issue('error', 'odf-mimetype-not-stored', 'mimetype', `The "mimetype" entry must be stored (uncompressed); found compression method ${mimetype.method}.`))
  }
}

// ---------------------------------------------------------------------------
// OPC content-types / relationships
// ---------------------------------------------------------------------------

/**
 * Parses XML the "collapsed" (non-`preserveOrder`) way: one JS object whose
 * shape mirrors the document, used only for the handful of known, fixed
 * OPC/ODF control-file schemas below (`[Content_Types].xml`, `*.rels`,
 * `manifest.xml`) — never for arbitrary document content, which uses
 * `parseOrdered`'s `OrderedNode` tree instead. `fast-xml-parser` itself
 * types `parse()` as `any`; callers narrow the result to the specific
 * `*Tree` shape they expect via a `@type` cast.
 * @param {string} xmlText
 */
function parseSimpleXml(xmlText) {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })
  return parser.parse(xmlText)
}

/**
 * @template T
 * @param {T | T[] | undefined} value
 * @returns {T[]}
 */
function asList(value) {
  if (value === undefined) return []
  return Array.isArray(value) ? value : [value]
}

/** @typedef {{ '@_Extension'?: string }} ContentTypeDefault */
/** @typedef {{ '@_PartName'?: string }} ContentTypeOverride */
/** @typedef {{ Types?: { Default?: ContentTypeDefault | ContentTypeDefault[], Override?: ContentTypeOverride | ContentTypeOverride[] } }} ContentTypesTree */

/** @typedef {{ '@_Id'?: string, '@_Target'?: string, '@_TargetMode'?: string, '@_Type'?: string }} RelationshipEntry */
/** @typedef {{ Relationships?: { Relationship?: RelationshipEntry | RelationshipEntry[] } }} RelationshipsTree */

/** @typedef {{ '@_manifest:full-path'?: string }} ManifestFileEntry */
/** @typedef {{ 'manifest:manifest'?: { 'manifest:file-entry'?: ManifestFileEntry | ManifestFileEntry[] } }} ManifestTree */

/** @typedef {{ defaults: Set<string>, overrides: Set<string> }} ContentTypesIndex */

/**
 * @param {PackageFiles} files
 * @param {Issue[]} issues
 * @returns {ContentTypesIndex}
 */
function checkContentTypes(files, issues) {
  const xml = files.get('[Content_Types].xml')
  if (xml === undefined) {
    issues.push(issue('error', 'missing-content-types', '[Content_Types].xml', 'Package is missing [Content_Types].xml.'))
    return { defaults: new Set(), overrides: new Set() }
  }
  const parsed = parseOrdered(xml.toString('utf8'))
  if (!parsed.ok) {
    issues.push(issue('error', 'malformed-xml', '[Content_Types].xml', `Not well-formed: ${parsed.message}`))
    return { defaults: new Set(), overrides: new Set() }
  }

  /** @type {ContentTypesTree} */
  const tree = parseSimpleXml(xml.toString('utf8'))
  const defaultsSeen = new Map()
  const overridesSeen = new Map()
  for (const d of asList(tree.Types?.Default)) {
    const ext = String(d['@_Extension'] ?? '').toLowerCase()
    defaultsSeen.set(ext, (defaultsSeen.get(ext) ?? 0) + 1)
  }
  for (const o of asList(tree.Types?.Override)) {
    const partName = String(o['@_PartName'] ?? '')
    overridesSeen.set(partName, (overridesSeen.get(partName) ?? 0) + 1)
  }
  for (const [ext, count] of defaultsSeen) {
    if (count > 1) issues.push(issue('error', 'duplicate-content-type-default', '[Content_Types].xml', `Duplicate <Default> entry for extension "${ext}".`))
  }
  for (const [partName, count] of overridesSeen) {
    if (count > 1) issues.push(issue('error', 'duplicate-content-type-override', '[Content_Types].xml', `Duplicate <Override> entry for part "${partName}".`))
  }

  return { defaults: new Set(defaultsSeen.keys()), overrides: new Set(overridesSeen.keys()) }
}

/**
 * @param {ContentTypesIndex} contentTypes
 * @param {string} partPath
 * @returns {boolean}
 */
function hasContentType(contentTypes, partPath) {
  if (contentTypes.overrides.has(`/${partPath}`)) return true
  const dot = partPath.lastIndexOf('.')
  if (dot === -1) return false
  return contentTypes.defaults.has(partPath.slice(dot + 1).toLowerCase())
}

/**
 * @param {PackageFiles} files
 * @param {ContentTypesIndex} contentTypes
 * @param {Issue[]} issues
 * @returns {void}
 */
function checkEveryPartHasContentType(files, contentTypes, issues) {
  for (const partPath of files.keys()) {
    if (partPath === '[Content_Types].xml') continue
    if (!hasContentType(contentTypes, partPath)) {
      issues.push(issue('error', 'part-missing-content-type', partPath, `"${partPath}" has no Default or Override content-type entry.`))
    }
  }
}

const URL_SCHEME_PATTERN = /^[a-zA-Z][a-zA-Z\d+\-.]*:/

/**
 * @param {string} relsPath
 * @returns {string}
 */
function resolveRelsBaseDir(relsPath) {
  const marker = '_rels/'
  const index = relsPath.lastIndexOf(marker)
  return index === -1 ? '' : relsPath.slice(0, index)
}

/**
 * @param {string} p
 * @returns {string}
 */
function normalizePartPath(p) {
  const segments = p.split('/')
  const resolved = []
  for (const segment of segments) {
    if (segment === '' || segment === '.') continue
    if (segment === '..') resolved.pop()
    else resolved.push(segment)
  }
  return resolved.join('/')
}

/**
 * @param {PackageFiles} files
 * @param {ContentTypesIndex} contentTypes
 * @param {Issue[]} issues
 * @returns {Map<string, Set<string>>}
 */
function checkRelationshipFiles(files, contentTypes, issues) {
  const relsByPart = new Map() // owning part path -> Set<relationship id>
  for (const [relsPath, bytes] of files) {
    if (!relsPath.endsWith('.rels')) continue
    const xmlText = bytes.toString('utf8')
    const parsed = parseOrdered(xmlText)
    if (!parsed.ok) {
      issues.push(issue('error', 'malformed-xml', relsPath, `Not well-formed: ${parsed.message}`))
      continue
    }
    /** @type {RelationshipsTree} */
    const tree = parseSimpleXml(xmlText)
    const relationships = asList(tree.Relationships?.Relationship)
    const baseDir = resolveRelsBaseDir(relsPath)
    const ids = new Set()

    for (const rel of relationships) {
      const id = String(rel['@_Id'] ?? '')
      const target = String(rel['@_Target'] ?? '')
      const targetMode = rel['@_TargetMode']
      ids.add(id)
      if (targetMode === 'External' || URL_SCHEME_PATTERN.test(target)) continue

      const resolved = normalizePartPath(target.startsWith('/') ? target.slice(1) : `${baseDir}${target}`)
      if (!files.has(resolved)) {
        issues.push(issue('error', 'dangling-relationship-target', relsPath, `Relationship "${id}" targets "${resolved}", which is not in the package.`))
        continue
      }
      if (!hasContentType(contentTypes, resolved)) {
        issues.push(issue('error', 'relationship-target-missing-content-type', relsPath, `"${resolved}" is targeted by relationship "${id}" but has no content-type entry.`))
      }
    }

    // The owning part is `_rels/`'s sibling: `<baseDir><basename minus .rels>`.
    const basename = relsPath.slice(baseDir.length + '_rels/'.length, -'.rels'.length)
    relsByPart.set(`${baseDir}${basename}`, ids)
  }
  return relsByPart
}

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const R_ID_ATTR_NAMES = ['id', 'embed', 'link', 'cs', 'dm', 'lo', 'qs', 'href', 'pict', 'topLeft']

/**
 * Finds the namespace prefix a part binds to the `r:` relationships namespace (usually, but not necessarily, "r").
 * @param {string} xmlText
 * @returns {string | null}
 */
function relationshipsPrefixIn(xmlText) {
  const match = new RegExp(`xmlns:(\\w+)="${RELATIONSHIPS_NS.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`).exec(xmlText)
  return match ? match[1] : null
}

/**
 * @param {PackageFiles} files
 * @param {Map<string, Set<string>>} relsByPart
 * @param {Issue[]} issues
 * @returns {void}
 */
function checkDanglingRIdReferences(files, relsByPart, issues) {
  for (const [partPath, bytes] of files) {
    if (!partPath.endsWith('.xml') || partPath.endsWith('.rels') || partPath === '[Content_Types].xml') continue
    const xmlText = bytes.toString('utf8')
    const prefix = relationshipsPrefixIn(xmlText)
    if (!prefix) continue // part doesn't reference the relationships namespace at all

    const attrPattern = new RegExp(`\\b${prefix}:(${R_ID_ATTR_NAMES.join('|')})="([^"]*)"`, 'g')
    const referencedIds = new Set()
    let match
    while ((match = attrPattern.exec(xmlText)) !== null) {
      if (match[2] !== '') referencedIds.add(match[2])
    }
    if (referencedIds.size === 0) continue

    const declaredIds = relsByPart.get(partPath) ?? new Set()
    for (const id of referencedIds) {
      if (!declaredIds.has(id)) {
        issues.push(
          issue(
            'error',
            'dangling-rid',
            partPath,
            `References relationship "${id}" (via ${prefix}:*) that is not declared in its .rels file.`,
          ),
        )
      }
    }
  }
}

// ---------------------------------------------------------------------------
// ODF manifest
// ---------------------------------------------------------------------------

/**
 * @param {PackageFiles} files
 * @param {Issue[]} issues
 * @returns {void}
 */
function checkOdfManifest(files, issues) {
  const manifestXml = files.get('META-INF/manifest.xml')
  if (manifestXml === undefined) {
    issues.push(issue('error', 'missing-manifest', 'META-INF/manifest.xml', 'ODF package is missing META-INF/manifest.xml.'))
    return
  }
  const xmlText = manifestXml.toString('utf8')
  const parsed = parseOrdered(xmlText)
  if (!parsed.ok) {
    issues.push(issue('error', 'malformed-xml', 'META-INF/manifest.xml', `Not well-formed: ${parsed.message}`))
    return
  }
  /** @type {ManifestTree} */
  const tree = parseSimpleXml(xmlText)
  const entries = asList(tree['manifest:manifest']?.['manifest:file-entry'])
  const listedPaths = new Set()
  for (const entry of entries) {
    const fullPath = String(entry['@_manifest:full-path'] ?? '')
    if (fullPath !== '/') listedPaths.add(fullPath)
  }

  for (const [partPath] of files) {
    if (partPath === 'mimetype' || partPath === 'META-INF/manifest.xml') continue
    if (!listedPaths.has(partPath)) {
      issues.push(issue('error', 'odf-file-not-in-manifest', partPath, `"${partPath}" is in the package but not listed in META-INF/manifest.xml.`))
    }
  }
  for (const listed of listedPaths) {
    if (!files.has(listed)) {
      issues.push(issue('error', 'odf-manifest-dangling-entry', 'META-INF/manifest.xml', `manifest.xml lists "${listed}", which is not in the package.`))
    }
  }
  const hasRoot = entries.some((entry) => entry['@_manifest:full-path'] === '/')
  if (!hasRoot) {
    issues.push(issue('error', 'odf-manifest-missing-root', 'META-INF/manifest.xml', 'manifest.xml has no root file-entry (manifest:full-path="/").'))
  }
}

// ---------------------------------------------------------------------------
// Format detection
// ---------------------------------------------------------------------------

/** @type {Record<string, 'odt' | 'odp' | 'ods'>} */
const ODF_MIME_TO_KIND = {
  'application/vnd.oasis.opendocument.text': 'odt',
  'application/vnd.oasis.opendocument.presentation': 'odp',
  'application/vnd.oasis.opendocument.spreadsheet': 'ods',
}

const OFFICE_DOCUMENT_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

/**
 * The OPC "start part" for a package is whatever `_rels/.rels` names via an
 * `officeDocument` relationship (ECMA-376 Part 2, §10.2) — determining kind
 * this way, rather than by checking whether `word/document.xml` etc.
 * physically exists, means a package MISSING its main part still gets
 * classified correctly (so `REQUIRED_PARTS` below can flag it as missing,
 * instead of silently falling through to "unknown" and skipping that check
 * entirely).
 * @param {PackageFiles} files
 * @returns {string | null}
 */
function startPartKindFromRels(files) {
  const relsXml = files.get('_rels/.rels')
  if (relsXml === undefined) return null
  /** @type {RelationshipsTree} */
  let tree
  try {
    tree = parseSimpleXml(relsXml.toString('utf8'))
  } catch {
    return null
  }
  const relationships = asList(tree.Relationships?.Relationship)
  const officeDoc = relationships.find((rel) => rel['@_Type'] === OFFICE_DOCUMENT_REL_TYPE)
  const target = officeDoc ? String(officeDoc['@_Target'] ?? '') : ''
  // Exact-match the well-known main-part names rather than just prefix-
  // matching the folder: `xl/workbook.bin` (the binary XLSB variant) starts
  // with `xl/` same as `xl/workbook.xml` does, but requires a different
  // `REQUIRED_PARTS` entry (no `xl/workbook.xml` ever exists in an XLSB
  // package) — conflating the two produced a false `missing-required-part`
  // on every XLSB save.
  if (target === 'word/document.xml') return 'docx'
  if (target === 'ppt/presentation.xml') return 'pptx'
  if (target === 'xl/workbook.xml') return 'xlsx'
  if (target === 'xl/workbook.bin') return 'xlsb'
  return null
}

/**
 * @param {PackageFiles} files
 * @returns {OfficeFormat}
 */
function detectFormat(files) {
  const mimetype = files.get('mimetype')
  if (mimetype !== undefined) {
    const content = mimetype.toString('utf8').trim()
    return { family: 'odf', kind: ODF_MIME_TO_KIND[content] ?? 'odf-unknown', mimetype: content }
  }
  if (files.has('[Content_Types].xml')) {
    const fromRels = startPartKindFromRels(files)
    if (fromRels === 'docx' || fromRels === 'pptx' || fromRels === 'xlsx' || fromRels === 'xlsb') {
      return { family: 'opc', kind: fromRels }
    }
    if (files.has('word/document.xml')) return { family: 'opc', kind: 'docx' }
    if (files.has('ppt/presentation.xml')) return { family: 'opc', kind: 'pptx' }
    if (files.has('xl/workbook.xml')) return { family: 'opc', kind: 'xlsx' }
    return { family: 'opc', kind: 'opc-unknown' }
  }
  return { family: 'unknown', kind: 'unknown' }
}

/** @type {Record<string, string[]>} */
const REQUIRED_PARTS = {
  docx: ['word/document.xml', '_rels/.rels'],
  pptx: ['ppt/presentation.xml', '_rels/.rels'],
  xlsx: ['xl/workbook.xml', '_rels/.rels'],
  xlsb: ['xl/workbook.bin', '_rels/.rels'],
  odt: ['content.xml', 'META-INF/manifest.xml', 'mimetype'],
  odp: ['content.xml', 'META-INF/manifest.xml', 'mimetype'],
  ods: ['content.xml', 'META-INF/manifest.xml', 'mimetype'],
}

// ---------------------------------------------------------------------------
// Top-level entry point
// ---------------------------------------------------------------------------

/**
 * @param {Buffer} buffer
 * @returns {OfficeValidationResult}
 */
export function validateOfficeFile(buffer) {
  /** @type {Issue[]} */
  const issues = []
  /** @type {ZipEntry[]} */
  let entries
  /** @type {PackageFiles} */
  let files
  try {
    ;({ entries, files } = readZipArchive(buffer))
  } catch (cause) {
    issues.push(issue('error', 'not-a-zip', null, cause instanceof Error ? cause.message : String(cause)))
    return { format: { family: 'unknown', kind: 'unknown' }, issues }
  }

  checkZipEntryNames(entries, issues)

  const format = detectFormat(files)

  for (const requiredPart of REQUIRED_PARTS[format.kind] ?? []) {
    if (!files.has(requiredPart)) {
      issues.push(issue('error', 'missing-required-part', requiredPart, `Required part "${requiredPart}" is missing for a .${format.kind} package.`))
    }
  }

  // Every XML/rels part: well-formed, no illegal characters, namespaces declared.
  for (const [partPath, bytes] of files) {
    if (!/\.(xml|rels)$/i.test(partPath)) continue
    const xmlText = bytes.toString('utf8')

    const illegal = findIllegalXmlChars(xmlText)
    if (illegal.length > 0) {
      const sample = illegal.map((c) => `U+${c.codePoint.toString(16).toUpperCase().padStart(4, '0')}@${c.index}`).join(', ')
      issues.push(issue('error', 'illegal-xml-char', partPath, `Contains character(s) XML 1.0 forbids: ${sample}.`))
    }

    const parsed = parseOrdered(xmlText)
    if (!parsed.ok) {
      issues.push(issue('error', 'malformed-xml', partPath, `Not well-formed: ${parsed.message}`))
      continue
    }
    checkNamespacesDeclared(parsed.tree, new Set(), partPath, issues)
    checkElementOrder(parsed.tree, partPath, issues)
    checkAttributeDatatypes(parsed.tree, partPath, issues)
  }

  if (format.family === 'odf') {
    checkOdfMimetypeEntry(entries, issues)
    checkOdfManifest(files, issues)
  } else if (format.family === 'opc') {
    const contentTypes = checkContentTypes(files, issues)
    checkEveryPartHasContentType(files, contentTypes, issues)
    const relsByPart = checkRelationshipFiles(files, contentTypes, issues)
    checkDanglingRIdReferences(files, relsByPart, issues)
  }

  return { format, issues }
}

/**
 * @param {string} filePath
 * @param {OfficeValidationResult} result
 * @returns {string}
 */
export function formatIssuesReport(filePath, result) {
  const lines = [`${filePath}  (${result.format.family}/${result.format.kind})`]
  if (result.issues.length === 0) {
    lines.push('  OK — no violations found.')
    return lines.join('\n')
  }
  for (const problem of result.issues) {
    lines.push(`  [${problem.severity}] ${problem.code} ${problem.part ? `(${problem.part})` : ''}: ${problem.message}`)
  }
  return lines.join('\n')
}

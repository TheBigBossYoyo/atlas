/**
 * USR-16 — shape-level PPTX edits. Each function takes a package and returns
 * a new one with only the touched parts rewritten; an edit that cannot be
 * applied (unknown shape, malformed part) returns the SAME package so callers
 * can detect "nothing happened" by reference.
 */
import {
  CONTENT_TYPE,
  NS,
  REL_TYPE,
  addOverride,
  addRel,
  children,
  descendants,
  escapeXml,
  firstChild,
  importFragment,
  nextPartName,
  parse,
  pxToEmu,
  readRels,
  relsPathFor,
  serialize,
} from './opcXml'
import { readPart, withParts, type OfficePackage } from '../../../../office/officePackage'

export type ShapeBox = { readonly x: number; readonly y: number; readonly w: number; readonly h: number }

const SHAPE_TAGS = new Set(['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'])

function findShape(doc: XMLDocument, sourceId: string): Element | null {
  for (const cNvPr of descendants(doc, 'cNvPr')) {
    if (cNvPr.getAttribute('id') !== sourceId) continue
    const shape = cNvPr.parentElement?.parentElement ?? null
    if (shape && SHAPE_TAGS.has(shape.localName)) return shape
  }
  return null
}

/** Applies `mutate` to the shape's slide DOM; returns the original package when the shape is missing. */
function editShape(
  pkg: OfficePackage,
  slidePath: string,
  sourceId: string,
  mutate: (shape: Element, doc: XMLDocument) => void,
): OfficePackage {
  const xml = readPart(pkg, slidePath)
  if (xml === null) return pkg
  const doc = parse(xml)
  const shape = findShape(doc, sourceId)
  if (!shape) return pkg
  mutate(shape, doc)
  return withParts(pkg, { [slidePath]: serialize(doc) })
}

/**
 * Replaces a text body's paragraphs with `text` (one paragraph per line),
 * keeping each paragraph's own properties and first-run formatting — the
 * formatting of the last existing paragraph carries over to added lines.
 */
export function replaceParagraphs(doc: XMLDocument, txBody: Element, text: string): void {
  const existing = children(txBody, 'p')
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  const created = lines.map((line, i) => {
    const template = existing[i] ?? existing[existing.length - 1] ?? null
    const paragraph = doc.createElementNS(NS.a, 'a:p')
    const pPr = template ? firstChild(template, 'pPr') : null
    if (pPr) paragraph.appendChild(pPr.cloneNode(true))
    const firstRun = template ? firstChild(template, 'r') : null
    const endParaRPr = template ? firstChild(template, 'endParaRPr') : null
    if (line !== '') {
      const run = doc.createElementNS(NS.a, 'a:r')
      const rPr = firstRun ? firstChild(firstRun, 'rPr') : null
      if (rPr) {
        run.appendChild(rPr.cloneNode(true))
      } else if (endParaRPr) {
        const fromEnd = doc.createElementNS(NS.a, 'a:rPr')
        for (const attr of Array.from(endParaRPr.attributes)) fromEnd.setAttribute(attr.name, attr.value)
        run.appendChild(fromEnd)
      }
      const t = doc.createElementNS(NS.a, 'a:t')
      t.textContent = line
      run.appendChild(t)
      paragraph.appendChild(run)
    }
    if (endParaRPr) paragraph.appendChild(endParaRPr.cloneNode(true))
    return paragraph
  })
  for (const paragraph of existing) txBody.removeChild(paragraph)
  for (const paragraph of created) txBody.appendChild(paragraph)
}

export function setShapeText(pkg: OfficePackage, slidePath: string, sourceId: string, text: string): OfficePackage {
  return editShape(pkg, slidePath, sourceId, (shape, doc) => {
    let txBody = firstChild(shape, 'txBody')
    if (!txBody) {
      txBody = importFragment(doc, '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>')
      shape.appendChild(txBody)
    }
    replaceParagraphs(doc, txBody, text)
  })
}

/** Writes the shape's own `a:xfrm` off/ext (creating it when the box was inherited from a placeholder). */
export function setShapeBox(pkg: OfficePackage, slidePath: string, sourceId: string, box: ShapeBox): OfficePackage {
  return editShape(pkg, slidePath, sourceId, (shape, doc) => {
    let xfrm: Element | null
    if (shape.localName === 'graphicFrame') {
      xfrm = firstChild(shape, 'xfrm')
      if (!xfrm) {
        xfrm = doc.createElementNS(NS.p, 'p:xfrm')
        shape.insertBefore(xfrm, firstChild(shape, 'graphic'))
      }
    } else {
      const propsTag = shape.localName === 'grpSp' ? 'grpSpPr' : 'spPr'
      let props = firstChild(shape, propsTag)
      if (!props) {
        props = doc.createElementNS(NS.p, `p:${propsTag}`)
        shape.insertBefore(props, shape.children[1] ?? null)
      }
      xfrm = firstChild(props, 'xfrm')
      if (!xfrm) {
        xfrm = doc.createElementNS(NS.a, 'a:xfrm')
        props.insertBefore(xfrm, props.firstElementChild)
      }
    }
    let off = firstChild(xfrm, 'off')
    let ext = firstChild(xfrm, 'ext')
    if (!off) {
      off = doc.createElementNS(NS.a, 'a:off')
      xfrm.insertBefore(off, xfrm.firstElementChild)
    }
    if (!ext) {
      ext = doc.createElementNS(NS.a, 'a:ext')
      xfrm.insertBefore(ext, off.nextElementSibling)
    }
    off.setAttribute('x', String(pxToEmu(box.x)))
    off.setAttribute('y', String(pxToEmu(box.y)))
    ext.setAttribute('cx', String(Math.max(pxToEmu(box.w), 1)))
    ext.setAttribute('cy', String(Math.max(pxToEmu(box.h), 1)))
  })
}

export function deleteShape(pkg: OfficePackage, slidePath: string, sourceId: string): OfficePackage {
  return editShape(pkg, slidePath, sourceId, (shape) => {
    shape.parentNode?.removeChild(shape)
  })
}

function nextShapeId(doc: XMLDocument): number {
  return descendants(doc, 'cNvPr').reduce((max, el) => Math.max(max, Number(el.getAttribute('id')) || 0), 0) + 1
}

/** Adds a plain text box at `box`; returns the package and the new shape's id. */
export function insertTextBox(
  pkg: OfficePackage,
  slidePath: string,
  box: ShapeBox,
  text: string,
): { readonly pkg: OfficePackage; readonly sourceId: string | null } {
  const xml = readPart(pkg, slidePath)
  if (xml === null) return { pkg, sourceId: null }
  const doc = parse(xml)
  const spTree = descendants(doc, 'spTree')[0]
  if (!spTree) return { pkg, sourceId: null }
  const id = nextShapeId(doc)
  const shape = importFragment(
    doc,
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.w)}" cy="${pxToEmu(box.h)}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
      `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>` +
      `<a:p><a:r><a:rPr lang="en-US" sz="1800" dirty="0"/><a:t>${escapeXml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`,
  )
  const extLst = firstChild(spTree, 'extLst')
  spTree.insertBefore(shape, extLst)
  return { pkg: withParts(pkg, { [slidePath]: serialize(doc) }), sourceId: String(id) }
}

// ---------------------------------------------------------------------------
// Speaker notes
// ---------------------------------------------------------------------------

function notesMasterPath(pkg: OfficePackage): string | null {
  const rels = readRels(readPart(pkg, relsPathFor('ppt/presentation.xml')), 'ppt/presentation.xml')
  return rels.find((rel) => rel.type === REL_TYPE.notesMaster)?.target ?? null
}

function notesSlidePathFor(pkg: OfficePackage, slidePath: string): string | null {
  const rels = readRels(readPart(pkg, relsPathFor(slidePath)), slidePath)
  return rels.find((rel) => rel.type === REL_TYPE.notesSlide)?.target ?? null
}

/** Notes can be edited when the slide already has a notes page or the deck has a notes master to create one from. */
export function canEditNotes(pkg: OfficePackage, slidePath: string): boolean {
  return notesSlidePathFor(pkg, slidePath) !== null || notesMasterPath(pkg) !== null
}

function newNotesSlideXml(text: string): string {
  const paragraphs = text
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => (line ? `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>` : '<a:p/>'))
    .join('')
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>` +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`
  )
}

export function setSlideNotes(pkg: OfficePackage, slidePath: string, text: string): OfficePackage {
  const existingPath = notesSlidePathFor(pkg, slidePath)
  if (existingPath) {
    const xml = readPart(pkg, existingPath)
    if (xml === null) return pkg
    const doc = parse(xml)
    const body = descendants(doc, 'sp').find((sp) => {
      const type = descendants(sp, 'ph')[0]?.getAttribute('type')
      return type === 'body'
    })
    const txBody = body ? firstChild(body, 'txBody') : null
    if (!txBody) return pkg
    replaceParagraphs(doc, txBody, text)
    return withParts(pkg, { [existingPath]: serialize(doc) })
  }

  const master = notesMasterPath(pkg)
  const contentTypes = readPart(pkg, '[Content_Types].xml')
  if (!master || contentTypes === null || text.trim() === '') return pkg
  const notesPath = nextPartName(pkg.parts.keys(), 'ppt/notesSlides', 'notesSlide')
  const notesRels = addRel(null, notesPath, REL_TYPE.notesMaster, master)
  const notesRelsWithSlide = addRel(notesRels.xml, notesPath, REL_TYPE.slide, slidePath)
  const slideRels = addRel(readPart(pkg, relsPathFor(slidePath)), slidePath, REL_TYPE.notesSlide, notesPath)
  return withParts(pkg, {
    [notesPath]: newNotesSlideXml(text),
    [relsPathFor(notesPath)]: notesRelsWithSlide.xml,
    [relsPathFor(slidePath)]: slideRels.xml,
    '[Content_Types].xml': addOverride(contentTypes, notesPath, CONTENT_TYPE.notesSlide),
  })
}

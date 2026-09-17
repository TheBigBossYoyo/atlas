/**
 * USR-16 — slide-level PPTX edits: add, duplicate, delete and reorder.
 * Registers/unregisters slide parts in `ppt/presentation.xml` (`p:sldIdLst`
 * plus any PowerPoint section lists that reference slide ids), the
 * presentation relationships and `[Content_Types].xml`.
 */
import {
  CONTENT_TYPE,
  NS,
  REL_TYPE,
  addOverride,
  addRel,
  children,
  descendants,
  firstChild,
  importFragment,
  nextPartName,
  parse,
  readRels,
  relsPathFor,
  removeOverrides,
  removeRels,
  serialize,
} from './opcXml'
import { readPart, withParts, type OfficePackage } from '../../../../office/officePackage'

const PRESENTATION = 'ppt/presentation.xml'
const PRESENTATION_RELS = relsPathFor(PRESENTATION)
const CONTENT_TYPES = '[Content_Types].xml'
/** Placeholders a new slide does not copy from its layout (they are master-driven). */
const SKIPPED_PLACEHOLDERS = new Set(['dt', 'ftr', 'sldNum'])

export type SlideEntry = { readonly sldId: string; readonly relId: string; readonly path: string }

function isSlideListEntry(el: Element): boolean {
  return el.namespaceURI === NS.p && el.parentElement?.localName === 'sldIdLst'
}

/** The deck's slides in presentation order. */
export function listSlides(pkg: OfficePackage): SlideEntry[] {
  const xml = readPart(pkg, PRESENTATION)
  if (xml === null) return []
  const targets = new Map(readRels(readPart(pkg, PRESENTATION_RELS), PRESENTATION).map((rel) => [rel.id, rel.target]))
  return descendants(parse(xml), 'sldId')
    .filter(isSlideListEntry)
    .flatMap((el) => {
      const relId = el.getAttributeNS(NS.r, 'id') ?? el.getAttribute('r:id')
      const path = relId ? targets.get(relId) : undefined
      return relId && path ? [{ sldId: el.getAttribute('id') ?? '', relId, path }] : []
    })
}

/** Registers `slidePath` in the presentation right after `afterSldId` (or at the end). */
function registerSlide(pkg: OfficePackage, slidePath: string, afterSldId: string | null): Record<string, string> | null {
  const presentationXml = readPart(pkg, PRESENTATION)
  const contentTypes = readPart(pkg, CONTENT_TYPES)
  if (presentationXml === null || contentTypes === null) return null

  const rel = addRel(readPart(pkg, PRESENTATION_RELS), PRESENTATION, REL_TYPE.slide, slidePath)
  const doc = parse(presentationXml)
  const all = descendants(doc, 'sldId')
  const nextId = Math.max(255, ...all.map((el) => Number(el.getAttribute('id')) || 0)) + 1

  let list = descendants(doc, 'sldIdLst').find((el) => el.namespaceURI === NS.p)
  if (!list) {
    list = doc.createElementNS(NS.p, 'p:sldIdLst')
    const anchor = descendants(doc, 'sldSz')[0] ?? descendants(doc, 'notesSz')[0] ?? null
    doc.documentElement.insertBefore(list, anchor)
  }
  const entry = doc.createElementNS(NS.p, 'p:sldId')
  entry.setAttribute('id', String(nextId))
  entry.setAttributeNS(NS.r, 'r:id', rel.id)

  const reference = afterSldId === null ? null : all.find((el) => isSlideListEntry(el) && el.getAttribute('id') === afterSldId)
  list.insertBefore(entry, reference ? reference.nextSibling : null)
  // Keep PowerPoint section lists (p14:sldIdLst in extLst) consistent with the slide list.
  if (afterSldId !== null) {
    for (const sectionRef of all.filter((el) => !isSlideListEntry(el) && el.getAttribute('id') === afterSldId)) {
      const clone = sectionRef.cloneNode(false) as Element
      clone.setAttribute('id', String(nextId))
      sectionRef.parentNode?.insertBefore(clone, sectionRef.nextSibling)
    }
  }

  return {
    [PRESENTATION]: serialize(doc),
    [PRESENTATION_RELS]: rel.xml,
    [CONTENT_TYPES]: addOverride(contentTypes, slidePath, CONTENT_TYPE.slide),
  }
}

function layoutPathOf(pkg: OfficePackage, slidePath: string): string | null {
  return readRels(readPart(pkg, relsPathFor(slidePath)), slidePath).find((rel) => rel.type === REL_TYPE.slideLayout)?.target ?? null
}

/** PowerPoint's "New Slide": the current slide's layout, except a title slide is followed by "Title and Content". */
function layoutForNewSlide(pkg: OfficePackage, referencePath: string | null): string | null {
  const layouts = [...pkg.parts.keys()].filter((path) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)).sort()
  const reference = referencePath ? layoutPathOf(pkg, referencePath) : null
  const typeOf = (path: string): string | null => {
    const xml = readPart(pkg, path)
    return xml ? parse(xml).documentElement.getAttribute('type') : null
  }
  if (reference && typeOf(reference) !== 'title') return reference
  return layouts.find((path) => typeOf(path) === 'obj') ?? reference ?? layouts[0] ?? null
}

function newSlideXml(pkg: OfficePackage, layoutPath: string): string {
  const layoutXml = readPart(pkg, layoutPath)
  const doc = parse(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">` +
      `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
  )
  const spTree = descendants(doc, 'spTree')[0]
  if (!layoutXml) return serialize(doc)

  let nextId = 2
  for (const placeholder of descendants(parse(layoutXml), 'sp')) {
    const ph = descendants(placeholder, 'ph')[0]
    const nvSpPr = firstChild(placeholder, 'nvSpPr')
    if (!ph || !nvSpPr || SKIPPED_PLACEHOLDERS.has(ph.getAttribute('type') ?? '')) continue
    const type = ph.getAttribute('type')
    const idx = ph.getAttribute('idx')
    const name = firstChild(nvSpPr, 'cNvPr')?.getAttribute('name') ?? `Placeholder ${nextId}`
    const phXml = `<p:ph${type ? ` type="${type}"` : ''}${idx ? ` idx="${idx}"` : ''}/>`
    spTree.appendChild(
      importFragment(
        doc,
        `<p:sp><p:nvSpPr><p:cNvPr id="${nextId}" name="${name.replace(/[<>&"]/g, '')}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
          `<p:nvPr>${phXml}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`,
      ),
    )
    nextId += 1
  }
  return serialize(doc)
}

/** Inserts a new slide after `afterIndex` (-1 for the start); returns the package and the new slide's position. */
export function addSlide(pkg: OfficePackage, afterIndex: number): { readonly pkg: OfficePackage; readonly index: number } {
  const slides = listSlides(pkg)
  const reference = slides[Math.min(Math.max(afterIndex, 0), slides.length - 1)] ?? null
  const layout = layoutForNewSlide(pkg, reference?.path ?? null)
  if (!layout) return { pkg, index: afterIndex }

  const slidePath = nextPartName(pkg.parts.keys(), 'ppt/slides', 'slide')
  const registration = registerSlide(pkg, slidePath, afterIndex < 0 ? null : (reference?.sldId ?? null))
  if (!registration) return { pkg, index: afterIndex }
  const rels = addRel(null, slidePath, REL_TYPE.slideLayout, layout)
  const next = withParts(pkg, { ...registration, [slidePath]: newSlideXml(pkg, layout), [relsPathFor(slidePath)]: rels.xml })
  return { pkg: next, index: afterIndex < 0 ? slides.length : afterIndex + 1 }
}

export function duplicateSlide(pkg: OfficePackage, index: number): OfficePackage {
  const source = listSlides(pkg)[index]
  const xml = source ? readPart(pkg, source.path) : null
  if (!source || xml === null) return pkg

  const slidePath = nextPartName(pkg.parts.keys(), 'ppt/slides', 'slide')
  const registration = registerSlide(pkg, slidePath, source.sldId)
  if (!registration) return pkg
  // The copy shares media/layout relationships but not the original's notes page.
  const sourceRels = readPart(pkg, relsPathFor(source.path))
  const rels = sourceRels
    ? removeRels(sourceRels, (rel) => rel.getAttribute('Type') === REL_TYPE.notesSlide)
    : null
  return withParts(pkg, { ...registration, [slidePath]: xml, ...(rels ? { [relsPathFor(slidePath)]: rels } : {}) })
}

export function deleteSlide(pkg: OfficePackage, index: number): OfficePackage {
  const slides = listSlides(pkg)
  const target = slides[index]
  const presentationXml = readPart(pkg, PRESENTATION)
  const presentationRels = readPart(pkg, PRESENTATION_RELS)
  const contentTypes = readPart(pkg, CONTENT_TYPES)
  if (slides.length <= 1 || !target || presentationXml === null || presentationRels === null || contentTypes === null) {
    return pkg
  }

  const doc = parse(presentationXml)
  for (const el of descendants(doc, 'sldId').filter((el) => el.getAttribute('id') === target.sldId)) {
    el.parentNode?.removeChild(el)
  }
  const notesPath = readRels(readPart(pkg, relsPathFor(target.path)), target.path).find(
    (rel) => rel.type === REL_TYPE.notesSlide,
  )?.target
  const removed = [target.path, ...(notesPath ? [notesPath] : [])]
  return withParts(pkg, {
    [PRESENTATION]: serialize(doc),
    [PRESENTATION_RELS]: removeRels(presentationRels, (rel) => rel.getAttribute('Id') === target.relId),
    [CONTENT_TYPES]: removeOverrides(contentTypes, removed),
    ...Object.fromEntries(removed.flatMap((path) => [[path, null], [relsPathFor(path), null]])),
  })
}

export function moveSlide(pkg: OfficePackage, from: number, to: number): OfficePackage {
  const xml = readPart(pkg, PRESENTATION)
  if (xml === null || from === to) return pkg
  const doc = parse(xml)
  const list = descendants(doc, 'sldIdLst').find((el) => el.namespaceURI === NS.p)
  const entries = list ? children(list, 'sldId') : []
  if (!list || from < 0 || to < 0 || from >= entries.length || to >= entries.length) return pkg
  const [moved] = entries.splice(from, 1)
  entries.splice(to, 0, moved)
  for (const entry of entries) list.appendChild(entry)
  return withParts(pkg, { [PRESENTATION]: serialize(doc) })
}

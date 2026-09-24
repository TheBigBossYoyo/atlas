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
  escapeXml,
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

// ---------------------------------------------------------------------------
// SHELL-5 — fallback slideMaster/slideLayout/theme for a package that has
// none at all (OOXML schema makes them effectively mandatory, but a lossy
// converter or hand-built file can still omit every one). Shaped the same
// way as the test fixture in `editing/__tests__/editableDeck.ts`, which
// `officeFileValidation.test.ts` proves is spec-clean.
// ---------------------------------------------------------------------------

/** Real PowerPoint's own id space for masters starts here (well above any `p:sldId`). */
const FALLBACK_MASTER_ID = 2147483648

const FALLBACK_THEME_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements>` +
  `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
  `<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
  `<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`

const FALLBACK_SLIDE_MASTER_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">` +
  `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
  `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
  `</p:sldMaster>`

const FALLBACK_SLIDE_LAYOUT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="blank">` +
  `<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
  `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`

/**
 * SHELL-5 — synthesizes a minimal, blank slideMaster + slideLayout + theme
 * and registers them (presentation.xml's `p:sldMasterIdLst`, presentation
 * rels, `[Content_Types].xml`) exactly like a real save from a real tool
 * would. Only called when the package has NO slideLayout at all — see
 * `addSlide` below. Returns `null` when `ppt/presentation.xml` or
 * `[Content_Types].xml` is itself missing (an even more broken package
 * `addSlide` already bails out of).
 */
function ensureFallbackLayout(pkg: OfficePackage): { readonly pkg: OfficePackage; readonly layoutPath: string } | null {
  const presentationXml = readPart(pkg, PRESENTATION)
  const contentTypes = readPart(pkg, CONTENT_TYPES)
  if (presentationXml === null || contentTypes === null) return null

  const themePath = nextPartName(pkg.parts.keys(), 'ppt/theme', 'theme')
  const masterPath = nextPartName(pkg.parts.keys(), 'ppt/slideMasters', 'slideMaster')
  const layoutPath = nextPartName(pkg.parts.keys(), 'ppt/slideLayouts', 'slideLayout')

  const masterRels = addRel(null, masterPath, REL_TYPE.theme, themePath)
  const layoutRels = addRel(null, layoutPath, REL_TYPE.slideMaster, masterPath)
  const presentationRels = addRel(readPart(pkg, PRESENTATION_RELS), PRESENTATION, REL_TYPE.slideMaster, masterPath)

  const doc = parse(presentationXml)
  const entry = doc.createElementNS(NS.p, 'p:sldMasterId')
  entry.setAttribute('id', String(FALLBACK_MASTER_ID))
  entry.setAttributeNS(NS.r, 'r:id', presentationRels.id)
  let list = descendants(doc, 'sldMasterIdLst').find((el) => el.namespaceURI === NS.p)
  if (!list) {
    list = doc.createElementNS(NS.p, 'p:sldMasterIdLst')
    // CT_Presentation's sequence: sldMasterIdLst comes before sldIdLst/sldSz.
    const anchor = descendants(doc, 'sldIdLst')[0] ?? descendants(doc, 'sldSz')[0] ?? null
    doc.documentElement.insertBefore(list, anchor)
  }
  list.appendChild(entry)

  const withOverrides = [
    [themePath, CONTENT_TYPE.theme],
    [masterPath, CONTENT_TYPE.slideMaster],
    [layoutPath, CONTENT_TYPE.slideLayout],
  ].reduce((ct, [path, type]) => addOverride(ct, path, type), contentTypes)

  return {
    pkg: withParts(pkg, {
      [themePath]: FALLBACK_THEME_XML,
      [masterPath]: FALLBACK_SLIDE_MASTER_XML,
      [relsPathFor(masterPath)]: masterRels.xml,
      [layoutPath]: FALLBACK_SLIDE_LAYOUT_XML,
      [relsPathFor(layoutPath)]: layoutRels.xml,
      [PRESENTATION]: serialize(doc),
      [PRESENTATION_RELS]: presentationRels.xml,
      [CONTENT_TYPES]: withOverrides,
    }),
    layoutPath,
  }
}

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
    // type/idx come from the (untrusted) layout part: escape them like any text.
    const phXml = `<p:ph${type ? ` type="${escapeXml(type)}"` : ''}${idx ? ` idx="${escapeXml(idx)}"` : ''}/>`
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

  let workingPkg = pkg
  let layout = layoutForNewSlide(workingPkg, reference?.path ?? null)
  if (!layout) {
    // SHELL-5 — a package with no slideLayout at all (non-conforming, but a
    // real file) used to make "New Slide" a silent no-op here. Synthesize a
    // minimal one instead, exactly once, rather than doing nothing.
    const fallback = ensureFallbackLayout(workingPkg)
    if (!fallback) return { pkg, index: afterIndex }
    workingPkg = fallback.pkg
    layout = fallback.layoutPath
  }

  const slidePath = nextPartName(workingPkg.parts.keys(), 'ppt/slides', 'slide')
  const registration = registerSlide(workingPkg, slidePath, afterIndex < 0 ? null : (reference?.sldId ?? null))
  if (!registration) return { pkg, index: afterIndex }
  const rels = addRel(null, slidePath, REL_TYPE.slideLayout, layout)
  const next = withParts(workingPkg, { ...registration, [slidePath]: newSlideXml(workingPkg, layout), [relsPathFor(slidePath)]: rels.xml })
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

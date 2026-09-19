/** USR-16 — a small but structurally complete PPTX (layouts, notes master, notes, sections) for edit tests. */
import JSZip from 'jszip'

const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const rels = (items: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`

const PLACEHOLDERS = `
  <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="476250"/><a:ext cx="9525000" cy="952500"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:p/></p:txBody></p:sp>
  <p:sp><p:nvSpPr><p:cNvPr id="3" name="Content 2"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="1905000"/><a:ext cx="9525000" cy="3810000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>
  <p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer 3"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`

const layout = (type: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout ${NS} type="${type}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${PLACEHOLDERS}</p:spTree></p:cSld></p:sldLayout>`

const SLIDE_1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
  <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4000" b="1"/><a:t>Quarterly review</a:t></a:r><a:endParaRPr lang="en-US" sz="4000"/></a:p></p:txBody></p:sp>
  <p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="600000"><a:off x="190500" y="5715000"/><a:ext cx="1905000" cy="381000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" i="1"/><a:t>Draft</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`

const SLIDE_2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
  <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Agenda</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>`

const NOTES_1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
  <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
  <p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Welcome everyone</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:notes>`

export async function buildEditableDeck(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    // `Default Extension="png"` matters, not just for realism: OPC requires
    // every part to resolve to SOME content type (§10.1.2.2.1) — omitting
    // it left `ppt/media/image1.png` below with none at all, a genuine
    // spec violation `scripts/validate-office-file.mjs` catches.
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
      `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
      `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
      `<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`))
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId4"/></p:notesMasterIdLst>` +
      `<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>` +
      `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:section name="Main" id="{00000000-0000-0000-0000-000000000001}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>`,
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    rels(
      `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL}/slide" Target="slides/slide1.xml"/>` +
        `<Relationship Id="rId3" Type="${REL}/slide" Target="slides/slide2.xml"/><Relationship Id="rId4" Type="${REL}/notesMaster" Target="notesMasters/notesMaster1.xml"/>`,
    ),
  )
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${PLACEHOLDERS}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>`,
  )
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels(''))
  zip.file('ppt/slideLayouts/slideLayout1.xml', layout('title'))
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels(`<Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`))
  zip.file('ppt/slideLayouts/slideLayout2.xml', layout('obj'))
  zip.file('ppt/slideLayouts/_rels/slideLayout2.xml.rels', rels(`<Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`))
  zip.file('ppt/notesMasters/notesMaster1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notesMaster>`)
  zip.file('ppt/slides/slide1.xml', SLIDE_1)
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    rels(`<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>`),
  )
  zip.file('ppt/slides/slide2.xml', SLIDE_2)
  zip.file('ppt/slides/_rels/slide2.xml.rels', rels(`<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>`))
  zip.file('ppt/notesSlides/notesSlide1.xml', NOTES_1)
  zip.file(
    'ppt/notesSlides/_rels/notesSlide1.xml.rels',
    rels(`<Relationship Id="rId1" Type="${REL}/notesMaster" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId2" Type="${REL}/slide" Target="../slides/slide1.xml"/>`),
  )
  zip.file('ppt/media/image1.png', new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]))
  return zip.generateAsync({ type: 'arraybuffer' })
}

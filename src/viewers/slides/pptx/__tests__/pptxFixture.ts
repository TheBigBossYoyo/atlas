/**
 * S22 — an in-memory PPTX fixture covering S1/S3/S4/S5/S6/S7/S8/S9/S10/S11/S12/S14/S17/S19/S20/S21
 * in one small deck: a real `JSZip` works directly as the parser's `ZipArchive`
 * (no need to round-trip through bytes), so this builds the archive contents
 * as XML strings and hands the zip straight to `parsePptxSlides`.
 *
 * All EMU offsets below are exact multiples of 9525 (1px) so pixel
 * assertions in the test file can use plain equality instead of tolerances.
 */

import JSZip from 'jszip'

const PRESENTATION_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
    <p:sldId id="258" r:id="rId4"/>
  </p:sldIdLst>
</p:presentation>`

const PRESENTATION_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
</Relationships>`

const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TestTheme">
  <a:themeElements>
    <a:clrScheme name="Test">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`

const SLIDE_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="285750"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
        <p:txBody><a:p/></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1524000"/><a:ext cx="8229600" cy="4286250"/></a:xfrm></p:spPr>
        <p:txBody><a:p/></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" b="1"/></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr>
      <a:lvl2pPr><a:defRPr sz="2000"/></a:lvl2pPr>
    </p:bodyStyle>
    <p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`

const SLIDE_MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`

const SLIDE_LAYOUT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:p/></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:p/></p:txBody></p:sp>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="476250" y="476250"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr>
        <p:txBody><a:p/></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`

const SLIDE_LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`

// Slide 1 exercises: S1 (title has no own xfrm -> inherits master's, since the
// layout's own placeholder also has none; the idx=7 placeholder's own xfrm
// sets only rotation, regression-testing that its position/size still merge
// in from the layout's box field-by-field instead of collapsing to zero),
// S3 (accent1 theme color + master txStyles size/bold fallback), S4 (buChar
// bullets at two levels), S5 (a:br soft break inside the title), S6 (a
// table), S7 (a group's off/ext -> chOff/chExt composition), S8 (solid fill
// + roundRect geometry with a "no line" border that must not shadow the
// shape's own fill, plus the master's own background), S10 (srcRect crop +
// rotation), S11 (a chart graphicFrame), S19 (alt text), S20 (normAutofit),
// S21 (spcBef).
const SLIDE1_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr><a:normAutofit fontScale="92500" lnSpcReduction="10000"/></a:bodyPr>
          <a:p>
            <a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>Quarterly Report</a:t></a:r>
            <a:br/>
            <a:r><a:t>Q3 2024</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1524000"/><a:ext cx="3810000" cy="1905000"/></a:xfrm></p:spPr>
        <p:txBody>
          <a:p>
            <a:pPr lvl="0"><a:buChar char="-"/><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>
            <a:r><a:t>Revenue up 12%</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="1"><a:buChar char="-"/></a:pPr>
            <a:r><a:t>Driven by EMEA</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:nvPr id="9" name="Banner"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="3810000"/><a:ext cx="1905000" cy="476250"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
        <p:txBody><a:p/></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm rot="5400000"/></p:spPr>
        <p:txBody><a:p><a:r><a:t>Rotated placeholder</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="10" name="Table1"/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="457200" y="4286250"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tr h="370840">
                <a:tc><a:txBody><a:p><a:r><a:t>Region</a:t></a:r></a:p></a:txBody></a:tc>
                <a:tc><a:txBody><a:p><a:r><a:t>Growth</a:t></a:r></a:p></a:txBody></a:tc>
              </a:tr>
              <a:tr h="370840">
                <a:tc><a:txBody><a:p><a:r><a:t>EMEA</a:t></a:r></a:p></a:txBody></a:tc>
                <a:tc><a:txBody><a:p><a:r><a:t>18%</a:t></a:r></a:p></a:txBody></a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="11" name="Chart1"/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="5715000" y="4286250"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="12" name="Pic1" descr="A quarterly chart"/></p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rIdImg"/>
          <a:srcRect l="10000" t="20000" r="5000" b="15000"/>
        </p:blipFill>
        <p:spPr><a:xfrm rot="2700000"><a:off x="6667500" y="952500"/><a:ext cx="1428750" cy="1428750"/></a:xfrm></p:spPr>
      </p:pic>
      <p:grpSp>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="952500" y="5715000"/><a:ext cx="1905000" cy="1905000"/>
            <a:chOff x="0" y="0"/><a:chExt cx="3810000" cy="3810000"/>
          </a:xfrm>
        </p:grpSpPr>
        <p:sp>
          <p:nvSpPr><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></a:xfrm>
            <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
          </p:spPr>
          <p:txBody><a:p/></p:txBody>
        </p:sp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
</p:sld>`

const SLIDE1_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>`

const NOTES_SLIDE1_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:p><a:r><a:t>Remember to mention EMEA growth drivers.</a:t></a:r><a:br/><a:r><a:t>Follow up with APAC next.</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`

// S14 — hidden via show="0". Slide 3 (referenced in presentation.xml.rels but
// never added to the archive) covers S9's per-slide error isolation instead.
const SLIDE2_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" show="0">
  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Hidden slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>`

export function buildPptxFixtureZip(): JSZip {
  const zip = new JSZip()

  zip.file('ppt/presentation.xml', PRESENTATION_XML)
  zip.file('ppt/_rels/presentation.xml.rels', PRESENTATION_RELS)
  zip.file('ppt/theme/theme1.xml', THEME_XML)
  zip.file('ppt/slideMasters/slideMaster1.xml', SLIDE_MASTER_XML)
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', SLIDE_MASTER_RELS)
  zip.file('ppt/slideLayouts/slideLayout1.xml', SLIDE_LAYOUT_XML)
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', SLIDE_LAYOUT_RELS)
  zip.file('ppt/slides/slide1.xml', SLIDE1_XML)
  zip.file('ppt/slides/_rels/slide1.xml.rels', SLIDE1_RELS)
  zip.file('ppt/notesSlides/notesSlide1.xml', NOTES_SLIDE1_XML)
  zip.file('ppt/media/image1.png', 'FAKEPNGDATA')
  zip.file('ppt/slides/slide2.xml', SLIDE2_XML)
  // slide3.xml intentionally omitted — S9 test.

  return zip
}

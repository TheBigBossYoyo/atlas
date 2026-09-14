/**
 * S22 — an in-memory ODP fixture covering S2/S3/S4/S5/S6/S7/S8/S11/S12/S14/S17/S19
 * in one small deck. As with the PPTX fixture, a real `JSZip` works directly
 * as the parser's `ZipArchive` — no byte round-trip needed.
 *
 * Every length below is written in bare `px` so pixel assertions can use
 * plain equality; `styles.xml`'s title style uses `pt` once to prove unit
 * conversion still works.
 */

import JSZip from 'jszip'

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  office:version="1.2">
  <office:styles>
    <style:style style:name="TitleText" style:family="text">
      <style:text-properties fo:font-size="32pt" fo:font-weight="bold" fo:color="#4472C4"/>
    </style:style>
    <text:list-style style:name="OutlineList">
      <text:list-level-style-bullet text:level="1" text:bullet-char="&#8226;"/>
      <text:list-level-style-bullet text:level="2" text:bullet-char="&#9702;"/>
    </text:list-style>
  </office:styles>
  <office:automatic-styles>
    <style:page-layout style:name="PM1">
      <style:page-layout-properties fo:page-width="1280px" fo:page-height="720px"/>
    </style:page-layout>
    <style:style style:name="Rect1" style:family="graphic">
      <style:graphic-properties draw:fill="solid" draw:fill-color="#FF0000" draw:stroke="solid" svg:stroke-color="#000000" svg:stroke-width="2px"/>
    </style:style>
    <style:style style:name="DP1" style:family="drawing-page">
      <style:drawing-page-properties draw:fill="solid" draw:fill-color="#FFFFFF"/>
    </style:style>
  </office:automatic-styles>
  <office:master-styles>
    <style:master-page style:name="Main" style:page-layout-name="PM1" draw:style-name="DP1"/>
  </office:master-styles>
</office:document-styles>`

const CONTENT_XML = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  office:version="1.2">
  <office:automatic-styles>
    <style:style style:name="P1" style:family="paragraph">
      <style:text-properties fo:font-size="24px"/>
    </style:style>
  </office:automatic-styles>
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1" draw:master-page-name="Main">
        <draw:frame presentation:class="title" svg:x="40px" svg:y="20px" svg:width="800px" svg:height="100px">
          <draw:text-box><text:p><text:span text:style-name="TitleText">Quarterly Report</text:span><text:line-break/><text:span>Q3 2024</text:span></text:p></draw:text-box>
        </draw:frame>
        <draw:frame presentation:class="outline" svg:x="40px" svg:y="140px" svg:width="600px" svg:height="200px">
          <draw:text-box>
            <text:list text:style-name="OutlineList">
              <text:list-item><text:p text:style-name="P1">Revenue up 12%</text:p></text:list-item>
              <text:list-item>
                <text:list>
                  <text:list-item><text:p>Driven by EMEA</text:p></text:list-item>
                </text:list>
              </text:list-item>
            </text:list>
          </draw:text-box>
        </draw:frame>
        <draw:frame svg:x="40px" svg:y="360px" svg:width="400px" svg:height="100px">
          <table:table>
            <table:table-row>
              <table:table-cell><text:p>Region</text:p></table:table-cell>
              <table:table-cell><text:p>Growth</text:p></table:table-cell>
            </table:table-row>
            <table:table-row>
              <table:table-cell><text:p>EMEA</text:p></table:table-cell>
              <table:table-cell><text:p>18%</text:p></table:table-cell>
            </table:table-row>
          </table:table>
        </draw:frame>
        <draw:g draw:transform="translate(100px 500px)">
          <draw:rect svg:x="0px" svg:y="0px" svg:width="100px" svg:height="60px" draw:style-name="Rect1"/>
        </draw:g>
        <draw:frame svg:x="500px" svg:y="360px" svg:width="150px" svg:height="150px">
          <draw:image xlink:href="Pictures/image1.png"/>
          <svg:title>A quarterly chart</svg:title>
        </draw:frame>
        <draw:frame svg:x="500px" svg:y="20px" svg:width="200px" svg:height="80px">
          <draw:object xlink:href="Object 1"/>
        </draw:frame>
        <presentation:notes>
          <draw:frame presentation:class="notes" svg:x="0px" svg:y="0px" svg:width="400px" svg:height="200px">
            <draw:text-box><text:p>Remember to mention EMEA growth drivers.</text:p></draw:text-box>
          </draw:frame>
        </presentation:notes>
      </draw:page>
      <draw:page draw:name="Slide 2" draw:master-page-name="Main" presentation:visibility="hidden">
        <draw:frame svg:x="0px" svg:y="0px" svg:width="100px" svg:height="50px">
          <draw:text-box><text:p>Hidden slide</text:p></draw:text-box>
        </draw:frame>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>`

export function buildOdpFixtureZip(): JSZip {
  const zip = new JSZip()

  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation')
  zip.file('styles.xml', STYLES_XML)
  zip.file('content.xml', CONTENT_XML)
  zip.file('Pictures/image1.png', 'FAKEPNGDATA')

  return zip
}

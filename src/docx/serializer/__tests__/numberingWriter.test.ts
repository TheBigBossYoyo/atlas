import { describe, expect, it } from 'vitest'

import { halfPoint, twip } from '../../model'
import type { NumberingPart } from '../../parser/numbering'
import { writeNumberingXml } from '../numberingWriter'

describe('writeNumberingXml', () => {
  it('writes a single bullet list definition', () => {
    const xml = writeNumberingXml({
      abstractNums: new Map([
        [
          '7',
          {
            abstractNumId: '7',
            levels: new Map([
              [
                0,
                {
                  level: 0,
                  format: 'bullet',
                  text: { value: '•', placeholders: [] },
                  paragraph: {
                    ind: {
                      left: twip(720),
                      hanging: twip(360),
                    },
                  },
                  run: {
                    sz: halfPoint(20),
                  },
                },
              ],
            ]),
          },
        ],
      ]),
      nums: new Map([['9', { numId: '9', abstractNumId: '7' }]]),
    } satisfies NumberingPart)

    expect(xml.startsWith('<?xml')).toBe(true)
    expect(xml).toContain('<w:numbering ')
    expect(xml).toContain('<w:abstractNum w:abstractNumId="7">')
    expect(xml).toContain('<w:numFmt w:val="bullet"/>')
    expect(xml).toContain('<w:lvlText w:val="•"/>')
  })

  it('writes a single ordered list definition', () => {
    const xml = writeNumberingXml({
      abstractNums: new Map([
        [
          '0',
          {
            abstractNumId: '0',
            levels: new Map([
              [
                0,
                {
                  level: 0,
                  start: 1,
                  format: 'decimal',
                  text: { value: '%1.', placeholders: [1] },
                  justification: 'start',
                },
              ],
            ]),
          },
        ],
      ]),
      nums: new Map(),
    } satisfies NumberingPart)

    expect(xml).toContain('<w:numFmt w:val="decimal"/>')
    expect(xml).toContain('<w:start w:val="1"/>')
    expect(xml).toContain('<w:lvlJc w:val="start"/>')
  })

  it('writes multi-level numbering with three levels', () => {
    const xml = writeNumberingXml({
      abstractNums: new Map([
        [
          '1',
          {
            abstractNumId: '1',
            levels: new Map([
              [0, { level: 0, format: 'decimal', text: { value: '%1.', placeholders: [1] } }],
              [1, { level: 1, format: 'lowerLetter', text: { value: '%2)', placeholders: [2] } }],
              [2, { level: 2, format: 'upperRoman', text: { value: '%1.%2', placeholders: [1, 2] } }],
            ]),
          },
        ],
      ]),
      nums: new Map(),
    } satisfies NumberingPart)

    expect(xml).toContain('<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>')
    expect(xml).toContain('<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>')
    expect(xml).toContain('<w:lvl w:ilvl="2"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1.%2"/></w:lvl>')
  })

  it('writes abstractNum to num linkage and level overrides', () => {
    const xml = writeNumberingXml({
      abstractNums: new Map([
        [
          '2',
          {
            abstractNumId: '2',
            styleLink: 'ListBullet',
            numberStyleLink: 'ListBulletStyle',
            levels: new Map([[0, { level: 0, format: 'decimal', text: { value: '%1)', placeholders: [1] } }]]),
          },
        ],
      ]),
      nums: new Map([
        [
          '10',
          {
            numId: '10',
            abstractNumId: '2',
            levelOverrides: new Map([
              [
                0,
                {
                  level: 0,
                  startOverride: 5,
                },
              ],
            ]),
          },
        ],
      ]),
    } satisfies NumberingPart)

    expect(xml).toContain('<w:styleLink w:val="ListBullet"/>')
    expect(xml).toContain('<w:numStyleLink w:val="ListBulletStyle"/>')
    expect(xml).toContain('<w:num w:numId="10"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>')
  })
})

import { XMLBuilder } from 'fast-xml-parser'

import type { LvlDef, LvlOverride } from '../model'
import type { AbstractNum, NumberingPart, NumInstance } from '../parser/numbering'
import { buildParagraphPropertiesXml, buildRunPropertiesXml } from './stylesWriter'

type XmlPrimitive = string | number | boolean
type XmlValue = XmlPrimitive | XmlNode | XmlValue[]

interface XmlNode {
  readonly [key: string]: XmlValue | undefined
}

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

export function writeNumberingXml(numberingPart: NumberingPart): string {
  const abstractNums = Array.from(numberingPart.abstractNums.values()).map(buildAbstractNumXml)
  const nums = Array.from(numberingPart.nums.values()).map(buildNumXml)

  const root: XmlNode = {
    '@_xmlns:w': WORD_NAMESPACE,
    ...(abstractNums.length > 0 ? { 'w:abstractNum': abstractNums } : {}),
    ...(nums.length > 0 ? { 'w:num': nums } : {}),
  }

  return `${XML_DECLARATION}${xmlBuilder.build({ 'w:numbering': root })}`
}

function buildAbstractNumXml(abstractNum: AbstractNum): XmlNode {
  const levels = Array.from(abstractNum.levels.values()).map(buildLevelXml)

  return {
    '@_w:abstractNumId': abstractNum.abstractNumId,
    ...buildValElement('w:styleLink', abstractNum.styleLink),
    ...buildValElement('w:numStyleLink', abstractNum.numberStyleLink),
    ...(levels.length > 0 ? { 'w:lvl': levels } : {}),
  }
}

function buildNumXml(num: NumInstance): XmlNode {
  const levelOverrides = Array.from(num.levelOverrides?.values() ?? []).map(buildLevelOverrideXml)

  return {
    '@_w:numId': num.numId,
    ...buildValElement('w:abstractNumId', num.abstractNumId),
    ...(levelOverrides.length > 0 ? { 'w:lvlOverride': levelOverrides } : {}),
  }
}

function buildLevelXml(level: LvlDef): XmlNode {
  return {
    '@_w:ilvl': String(level.level),
    ...withAttribute('@_w:tentative', buildOnOffAttribute(level.tentative)),
    ...buildValElement('w:start', level.start),
    ...buildValElement('w:lvlRestart', level.restart),
    ...buildValElement('w:numFmt', level.format),
    ...buildValElement('w:lvlText', level.text?.value),
    ...buildValElement('w:suff', level.suffix),
    ...withElement('w:isLgl', buildOnOffElement(level.legal)),
    ...buildValElement('w:lvlJc', level.justification),
    ...buildValElement('w:pStyle', level.pStyle),
    ...withElement('w:pPr', buildParagraphPropertiesXml(level.paragraph)),
    ...withElement('w:rPr', buildRunPropertiesXml(level.run)),
  }
}

function buildLevelOverrideXml(levelOverride: LvlOverride): XmlNode {
  return {
    '@_w:ilvl': String(levelOverride.level),
    ...buildValElement('w:startOverride', levelOverride.startOverride),
    ...withElement('w:lvl', levelOverride.levelDefinition === undefined ? undefined : buildLevelXml(levelOverride.levelDefinition)),
  }
}

function buildOnOffElement(value: boolean | undefined): XmlNode | undefined {
  if (value === undefined) return undefined
  return value ? {} : { '@_w:val': '0' }
}

function buildOnOffAttribute(value: boolean | undefined): string | undefined {
  if (value === undefined) return undefined
  return value ? '1' : '0'
}

function buildValElement(name: string, value: string | number | undefined): XmlNode {
  return value === undefined ? {} : { [name]: { '@_w:val': String(value) } }
}

function withElement(name: string, value: Record<string, unknown> | XmlNode | undefined): XmlNode {
  return value === undefined ? {} : { [name]: value as XmlNode }
}

function withAttribute(name: string, value: string | undefined): XmlNode {
  return value === undefined ? {} : { [name]: value }
}

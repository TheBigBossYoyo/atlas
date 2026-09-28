import { XMLBuilder } from 'fast-xml-parser'

import type { LvlDef, LvlOverride } from '../model'
import type { AbstractNum, NumberingPart, NumInstance } from '../parser/numbering'
import {
  createRawPassthroughState,
  restoreRawPassthrough,
  type RawPassthroughState,
} from './rawPassthrough'
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
  // RPR-STYLES-1 — a `w:lvl`'s `w:rPr` is parsed by `parser/styles.ts`'s
  // `parseRunPropsNode`, so it carries the same `rPrUnknown` passthrough that
  // styles do, and this part has to restore the placeholders as well. The root
  // attributes come from the source for the same reason `styles.xml`'s do: a
  // re-emitted fragment can use a prefix only the source declared.
  const state = createRawPassthroughState()
  const abstractNums = Array.from(numberingPart.abstractNums.values()).map((abstractNum) =>
    buildAbstractNumXml(abstractNum, state),
  )
  const nums = Array.from(numberingPart.nums.values()).map((num) => buildNumXml(num, state))

  const root: XmlNode = {
    ...buildRootAttributes(numberingPart.rootAttributes),
    ...(abstractNums.length > 0 ? { 'w:abstractNum': abstractNums } : {}),
    ...(nums.length > 0 ? { 'w:num': nums } : {}),
  }

  return restoreRawPassthrough(`${XML_DECLARATION}${xmlBuilder.build({ 'w:numbering': root })}`, state)
}

function buildRootAttributes(attributes: NumberingPart['rootAttributes']): XmlNode {
  if (attributes === undefined || attributes.size === 0) {
    return { '@_xmlns:w': WORD_NAMESPACE }
  }

  const node: Record<string, string> = {}
  for (const [name, value] of attributes) {
    node[`@_${name}`] = value
  }
  if (node['@_xmlns:w'] === undefined) node['@_xmlns:w'] = WORD_NAMESPACE
  return node
}

function buildAbstractNumXml(abstractNum: AbstractNum, state: RawPassthroughState): XmlNode {
  const levels = Array.from(abstractNum.levels.values()).map((level) => buildLevelXml(level, state))

  return {
    '@_w:abstractNumId': abstractNum.abstractNumId,
    ...buildValElement('w:multiLevelType', abstractNum.multiLevelType),
    ...buildValElement('w:styleLink', abstractNum.styleLink),
    ...buildValElement('w:numStyleLink', abstractNum.numberStyleLink),
    ...(levels.length > 0 ? { 'w:lvl': levels } : {}),
  }
}

function buildNumXml(num: NumInstance, state: RawPassthroughState): XmlNode {
  const levelOverrides = Array.from(num.levelOverrides?.values() ?? []).map((levelOverride) =>
    buildLevelOverrideXml(levelOverride, state),
  )

  return {
    '@_w:numId': num.numId,
    ...buildValElement('w:abstractNumId', num.abstractNumId),
    ...(levelOverrides.length > 0 ? { 'w:lvlOverride': levelOverrides } : {}),
  }
}

// Child order follows ECMA-376 `CT_Lvl` (§17.9.6), restricted to the members
// `LvlDef` models. An element-order review found `lvlRestart`/`numFmt`
// swapped and `pStyle`/`isLgl`/`suff`/`lvlText` all out of sequence.
function buildLevelXml(level: LvlDef, state: RawPassthroughState): XmlNode {
  return {
    '@_w:ilvl': String(level.level),
    ...withAttribute('@_w:tentative', buildOnOffAttribute(level.tentative)),
    ...buildValElement('w:start', level.start),
    ...buildValElement('w:numFmt', level.format),
    ...buildValElement('w:lvlRestart', level.restart),
    ...buildValElement('w:pStyle', level.pStyle),
    ...withElement('w:isLgl', buildOnOffElement(level.legal)),
    ...buildValElement('w:suff', level.suffix),
    ...buildValElement('w:lvlText', level.text?.value),
    ...buildValElement('w:lvlJc', level.justification),
    ...withElement('w:pPr', buildParagraphPropertiesXml(level.paragraph)),
    ...withElement('w:rPr', buildRunPropertiesXml(level.run, state)),
  }
}

function buildLevelOverrideXml(levelOverride: LvlOverride, state: RawPassthroughState): XmlNode {
  return {
    '@_w:ilvl': String(levelOverride.level),
    ...buildValElement('w:startOverride', levelOverride.startOverride),
    ...withElement(
      'w:lvl',
      levelOverride.levelDefinition === undefined ? undefined : buildLevelXml(levelOverride.levelDefinition, state),
    ),
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

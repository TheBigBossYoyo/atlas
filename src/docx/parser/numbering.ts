/**
 * Atlas — DOCX numbering.xml parser (Wave A.4)
 *
 * Parses abstract numbering definitions plus concrete numbering instances.
 */

import { XMLParser } from 'fast-xml-parser'

import type { JustifyContent, LvlDef, LvlOverride, NumberingSuffix } from '../model'
import { parseParaPropsNode, parseRunPropsNode } from './styles'
import { assertXmlPartSizeWithinLimit } from './xmlSizeGuard'

type XmlScalar = string | number | boolean
type XmlValue = XmlScalar | XmlNode | XmlValue[]

interface XmlNode {
  readonly [key: string]: XmlValue | undefined
}

interface RawNumberingDocument {
  readonly 'w:numbering'?: XmlValue
}

export interface AbstractNum {
  readonly abstractNumId: string
  readonly styleLink?: string
  readonly numberStyleLink?: string
  readonly levels: ReadonlyMap<number, LvlDef>
}

export interface NumInstance {
  readonly numId: string
  readonly abstractNumId?: string
  readonly levelOverrides?: ReadonlyMap<number, LvlOverride>
}

export interface NumberingPart {
  readonly abstractNums: ReadonlyMap<string, AbstractNum>
  readonly nums: ReadonlyMap<string, NumInstance>
}

const xmlParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })

export function parseNumbering(xml: string): NumberingPart {
  assertXmlPartSizeWithinLimit(xml, 'word/numbering.xml')
  const raw = xmlParser.parse(xml) as RawNumberingDocument
  const numberingRoot = asXmlNode(raw['w:numbering'])

  const abstractNums = new Map<string, AbstractNum>()
  const nums = new Map<string, NumInstance>()

  for (const abstractNumNode of getNodes(numberingRoot, 'w:abstractNum')) {
    const abstractNum = parseAbstractNumNode(abstractNumNode)
    if (abstractNum === undefined) continue
    abstractNums.set(abstractNum.abstractNumId, abstractNum)
  }

  for (const numNode of getNodes(numberingRoot, 'w:num')) {
    const num = parseNumNode(numNode)
    if (num === undefined) continue
    nums.set(num.numId, num)
  }

  return { abstractNums, nums }
}

function parseAbstractNumNode(node: XmlNode): AbstractNum | undefined {
  const abstractNumId = getAttr(node, 'w:abstractNumId')?.trim()
  if (abstractNumId === undefined || abstractNumId === '') return undefined

  const levels = new Map<number, LvlDef>()

  for (const levelNode of getNodes(node, 'w:lvl')) {
    const level = parseLevelNode(levelNode)
    if (level === undefined) continue
    levels.set(level.level, level)
  }

  return {
    abstractNumId,
    ...(withValue('styleLink', getValAttr(getNode(node, 'w:styleLink')))),
    ...(withValue('numberStyleLink', getValAttr(getNode(node, 'w:numStyleLink')))),
    levels,
  }
}

function parseNumNode(node: XmlNode): NumInstance | undefined {
  const numId = getAttr(node, 'w:numId')?.trim()
  if (numId === undefined || numId === '') return undefined

  const levelOverrides = new Map<number, LvlOverride>()

  for (const overrideNode of getNodes(node, 'w:lvlOverride')) {
    const levelOverride = parseLevelOverrideNode(overrideNode)
    if (levelOverride === undefined) continue
    levelOverrides.set(levelOverride.level, levelOverride)
  }

  const abstractNumId = getValAttr(getNode(node, 'w:abstractNumId'))

  return {
    numId,
    ...(withValue('abstractNumId', abstractNumId)),
    ...(levelOverrides.size > 0 ? { levelOverrides } : {}),
  }
}

function parseLevelNode(node: XmlNode): LvlDef | undefined {
  const level = parseInteger(getAttr(node, 'w:ilvl'))
  if (level === undefined) return undefined

  return {
    level,
    ...(withValue('start', parseInteger(getValAttr(getNode(node, 'w:start'))))),
    ...(withValue('restart', parseInteger(getValAttr(getNode(node, 'w:lvlRestart'))))),
    ...(withValue('format', getValAttr(getNode(node, 'w:numFmt')))),
    ...(withValue('text', parseLevelText(getValAttr(getNode(node, 'w:lvlText'))))),
    ...(withValue('suffix', parseNumberingSuffix(getValAttr(getNode(node, 'w:suff'))))),
    ...(withValue('tentative', parseOnOffAttr(getAttr(node, 'w:tentative')))),
    ...(withValue('legal', parseOnOffElement(node['w:isLgl']))),
    ...(withValue('justification', parseJustifyContent(getValAttr(getNode(node, 'w:lvlJc'))))),
    ...(withValue('pStyle', getValAttr(getNode(node, 'w:pStyle')))),
    ...(withValue('paragraph', parseParaPropsNode(getNode(node, 'w:pPr')))),
    ...(withValue('run', parseRunPropsNode(getNode(node, 'w:rPr')))),
  }
}

function parseLevelOverrideNode(node: XmlNode): LvlOverride | undefined {
  const level = parseInteger(getAttr(node, 'w:ilvl'))
  if (level === undefined) return undefined

  const startOverride = parseInteger(getValAttr(getNode(node, 'w:startOverride')))
  const levelDefinition = parseLevelNode(getNode(node, 'w:lvl') ?? {})

  if (startOverride === undefined && levelDefinition === undefined) {
    return { level }
  }

  return {
    level,
    ...(withValue('startOverride', startOverride)),
    ...(withValue('levelDefinition', levelDefinition)),
  }
}

function parseLevelText(value: string | undefined): LvlDef['text'] | undefined {
  if (value === undefined) return undefined

  const placeholders: number[] = []
  const matches = value.matchAll(/%(\d+)/g)

  for (const match of matches) {
    const placeholder = Number.parseInt(match[1], 10)
    if (Number.isFinite(placeholder)) {
      placeholders.push(placeholder)
    }
  }

  return {
    value,
    placeholders,
  }
}

function parseNumberingSuffix(value: string | undefined): NumberingSuffix | undefined {
  switch (value) {
    case 'tab':
    case 'space':
    case 'nothing':
      return value
    default:
      return undefined
  }
}

function parseJustifyContent(value: string | undefined): JustifyContent | undefined {
  switch (value) {
    case 'left':
    case 'start':
      return 'start'
    case 'center':
      return 'center'
    case 'right':
    case 'end':
      return 'end'
    case 'both':
      return 'both'
    case 'distribute':
      return 'distribute'
    default:
      return undefined
  }
}

function parseOnOffElement(value: XmlValue | undefined): boolean | undefined {
  if (value === undefined) return undefined
  if (Array.isArray(value)) {
    return value.length === 0 ? true : parseOnOffElement(value[0])
  }
  if (isXmlNode(value)) {
    return parseOnOffAttr(getAttr(value, 'w:val')) ?? true
  }
  return parseOnOffAttr(toOptionalString(value)) ?? true
}

function parseOnOffAttr(value: string | undefined): boolean | undefined {
  if (value === undefined) return undefined

  switch (value.toLowerCase()) {
    case '0':
    case 'false':
    case 'off':
      return false
    case '1':
    case 'true':
    case 'on':
    case '':
      return true
    default:
      return true
  }
}

function getNodes(parent: XmlNode | undefined, key: string): XmlNode[] {
  return toArray(parent?.[key]).filter((value): value is XmlNode => isXmlNode(value))
}

function getNode(parent: XmlNode | undefined, key: string): XmlNode | undefined {
  return asXmlNode(parent?.[key])
}

function getValAttr(node: XmlNode | undefined): string | undefined {
  return getAttr(node, 'w:val')
}

function getAttr(node: XmlNode | undefined, key: string): string | undefined {
  if (node === undefined) return undefined
  return toOptionalString(node[`@_${key}`])
}

function parseInteger(value: string | undefined): number | undefined {
  if (value === undefined || value.trim() === '') return undefined
  const parsed = Number.parseInt(value, 10)
  return Number.isFinite(parsed) ? parsed : undefined
}

function toArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? [...value] : [value]
}

function isXmlNode(value: XmlValue | unknown): value is XmlNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function asXmlNode(value: XmlValue | unknown): XmlNode | undefined {
  return isXmlNode(value) ? value : undefined
}

function toOptionalString(value: XmlValue | undefined): string | undefined {
  if (typeof value === 'string') return value
  if (typeof value === 'number' || typeof value === 'boolean') return String(value)
  return undefined
}

function withValue<K extends string, V>(
  key: K,
  value: V | undefined,
) {
  if (value === undefined) return {}
  return { [key]: value }
}

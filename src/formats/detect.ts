import type { FormatId } from './types'

const EXTENSION_MAP: Readonly<Record<string, FormatId>> = {
  md: 'markdown',
  markdown: 'markdown',
  mdx: 'markdown',
  mkd: 'markdown',
  docx: 'docx',
  xlsx: 'xlsx',
  xlsm: 'xlsx',
  xlsb: 'xlsx',
  pptx: 'pptx',
  pptm: 'pptx',
  pdf: 'pdf',
  csv: 'csv',
  tsv: 'tsv',
  tab: 'tsv',
  txt: 'text',
  log: 'text',
  ts: 'code',
  tsx: 'code',
  js: 'code',
  jsx: 'code',
  mjs: 'code',
  cjs: 'code',
  py: 'code',
  pyw: 'code',
  rb: 'code',
  go: 'code',
  rs: 'code',
  java: 'code',
  kt: 'code',
  kts: 'code',
  swift: 'code',
  c: 'code',
  h: 'code',
  cpp: 'code',
  hpp: 'code',
  cc: 'code',
  cs: 'code',
  php: 'code',
  pl: 'code',
  lua: 'code',
  r: 'code',
  scala: 'code',
  clj: 'code',
  ex: 'code',
  exs: 'code',
  erl: 'code',
  hs: 'code',
  ml: 'code',
  dart: 'code',
  vue: 'code',
  svelte: 'code',
  sh: 'code',
  bash: 'code',
  zsh: 'code',
  fish: 'code',
  ps1: 'code',
  bat: 'code',
  cmd: 'code',
  json: 'code',
  jsonc: 'code',
  yaml: 'code',
  yml: 'code',
  toml: 'code',
  xml: 'code',
  html: 'code',
  htm: 'code',
  css: 'code',
  scss: 'code',
  sass: 'code',
  less: 'code',
  sql: 'code',
  graphql: 'code',
  proto: 'code',
  dockerfile: 'code',
  ini: 'code',
  cfg: 'code',
  conf: 'code',
  odt: 'odt',
  ods: 'ods',
  odp: 'odp',
  rtf: 'rtf',
}

function getExtension(path: string): string | null {
  const fileName = path.split(/[\\/]/).pop() ?? path
  const dotIndex = fileName.lastIndexOf('.')

  if (dotIndex <= 0 || dotIndex === fileName.length - 1) {
    return null
  }

  return fileName.slice(dotIndex + 1).toLowerCase()
}

function decodeAscii(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer)
  let decoded = ''

  for (const byte of bytes) {
    decoded += String.fromCharCode(byte)
  }

  return decoded
}

function detectByMagicMarker(buffer: ArrayBuffer): FormatId {
  const bytes = new Uint8Array(buffer)

  if (bytes.length >= 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) {
    return 'pdf'
  }

  if (bytes.length >= 5 && bytes[0] === 0x7b && bytes[1] === 0x5c && bytes[2] === 0x72 && bytes[3] === 0x74 && bytes[4] === 0x66) {
    return 'rtf'
  }

  if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) {
    const ascii = decodeAscii(buffer)

    if (ascii.includes('word/document.xml')) return 'docx'
    if (ascii.includes('xl/workbook.xml')) return 'xlsx'
    if (ascii.includes('ppt/presentation.xml')) return 'pptx'
    if (ascii.includes('mimetypeapplication/vnd.oasis.opendocument.text')) return 'odt'
    if (ascii.includes('mimetypeapplication/vnd.oasis.opendocument.spreadsheet')) return 'ods'
    if (ascii.includes('mimetypeapplication/vnd.oasis.opendocument.presentation')) return 'odp'
  }

  return 'unknown'
}

export function detectByExtension(path: string): FormatId {
  const extension = getExtension(path)

  if (!extension) {
    return 'unknown'
  }

  return EXTENSION_MAP[extension] ?? 'unknown'
}

export function detectByMagic(buffer: ArrayBuffer): FormatId {
  return detectByMagicMarker(buffer)
}

export function detectFormat(path: string, buffer?: ArrayBuffer): FormatId {
  const extensionFormat = detectByExtension(path)

  if (!buffer) {
    return extensionFormat
  }

  const magicFormat = detectByMagic(buffer)

  if (magicFormat === 'unknown') {
    return extensionFormat
  }

  if (extensionFormat === 'text' || extensionFormat === 'code') {
    return extensionFormat
  }

  if (extensionFormat !== 'unknown' && extensionFormat === magicFormat) {
    return extensionFormat
  }

  return magicFormat
}

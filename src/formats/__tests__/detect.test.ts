import { describe, expect, it } from 'vitest'

import { detectByExtension, detectByMagic, detectFormat } from '../detect'
import type { FormatId } from '../types'

function bytesToBuffer(bytes: ReadonlyArray<number>): ArrayBuffer {
  return new Uint8Array(bytes).buffer
}

function asciiBuffer(prefix: ReadonlyArray<number>, text: string): ArrayBuffer {
  const encoded = new TextEncoder().encode(text)
  const bytes = new Uint8Array(prefix.length + encoded.length)

  bytes.set(prefix, 0)
  bytes.set(encoded, prefix.length)

  return bytes.buffer
}

function zipWithAscii(marker: string): ArrayBuffer {
  return asciiBuffer([0x50, 0x4b, 0x03, 0x04], marker)
}

const extensionCases = [
  ['README.md', 'markdown'],
  ['report.docx', 'docx'],
  ['budget.xlsx', 'xlsx'],
  ['deck.pptx', 'pptx'],
  ['paper.pdf', 'pdf'],
  ['data.csv', 'csv'],
  ['data.tsv', 'tsv'],
  ['notes.txt', 'text'],
  ['script.ts', 'code'],
  ['writer.odt', 'odt'],
  ['sheets.ods', 'ods'],
  ['slides.odp', 'odp'],
  ['legacy.rtf', 'rtf'],
  ['mystery.xyz', 'unknown'],
] as const satisfies ReadonlyArray<readonly [string, FormatId]>

const codeExtensionCases = [
  'js',
  'jsx',
  'ts',
  'tsx',
  'mjs',
  'cjs',
  'py',
  'pyw',
  'rb',
  'go',
  'rs',
  'java',
  'kt',
  'kts',
  'swift',
  'c',
  'h',
  'cpp',
  'hpp',
  'cc',
  'cs',
  'php',
  'pl',
  'lua',
  'r',
  'scala',
  'clj',
  'ex',
  'exs',
  'erl',
  'hs',
  'ml',
  'dart',
  'vue',
  'svelte',
  'sh',
  'bash',
  'zsh',
  'fish',
  'ps1',
  'bat',
  'cmd',
  'json',
  'jsonc',
  'yaml',
  'yml',
  'toml',
  'xml',
  'html',
  'htm',
  'css',
  'scss',
  'sass',
  'less',
  'sql',
  'graphql',
  'proto',
  'dockerfile',
  'ini',
  'cfg',
  'conf',
] as const

describe('detectByExtension', () => {
  it.each(extensionCases)('detects %s as %s', (path, expected) => {
    expect(detectByExtension(path)).toBe(expected)
  })

  it.each([
    ['guide.markdown', 'markdown'],
    ['post.mdx', 'markdown'],
    ['draft.mkd', 'markdown'],
    ['macro.xlsm', 'xlsx'],
    ['binary.xlsb', 'xlsx'],
    ['slides.pptm', 'pptx'],
    ['table.tab', 'tsv'],
    ['app.log', 'text'],
  ] as const)('supports alternate extension %s -> %s', (path, expected) => {
    expect(detectByExtension(path)).toBe(expected)
  })

  it.each(codeExtensionCases)('maps .%s to code', (extension) => {
    expect(detectByExtension(`source.${extension}`)).toBe('code')
  })

  it('matches extensions case-insensitively', () => {
    expect(detectByExtension('REPORT.PDF')).toBe('pdf')
    expect(detectByExtension('component.TSX')).toBe('code')
  })

  it('returns unknown for files without an extension', () => {
    expect(detectByExtension('README')).toBe('unknown')
  })
})

describe('detectByMagic', () => {
  it('detects PDF by %PDF header', () => {
    expect(detectByMagic(bytesToBuffer([0x25, 0x50, 0x44, 0x46, 0x2d]))).toBe('pdf')
  })

  it.each([
    ['word/document.xml', 'docx'],
    ['xl/workbook.xml', 'xlsx'],
    ['ppt/presentation.xml', 'pptx'],
  ] as const)('distinguishes OOXML payload %s as %s', (marker, expected) => {
    expect(detectByMagic(zipWithAscii(marker))).toBe(expected)
  })

  it.each([
    ['application/vnd.oasis.opendocument.text', 'odt'],
    ['application/vnd.oasis.opendocument.spreadsheet', 'ods'],
    ['application/vnd.oasis.opendocument.presentation', 'odp'],
  ] as const)('distinguishes ODF mimetype %s as %s', (mimeType, expected) => {
    expect(detectByMagic(zipWithAscii(`mimetype${mimeType}`))).toBe(expected)
  })

  it('returns unknown for ZIP without a recognized marker', () => {
    expect(detectByMagic(zipWithAscii('plain/archive.bin'))).toBe('unknown')
  })

  it('detects RTF by {\\rtf prefix', () => {
    expect(detectByMagic(bytesToBuffer([0x7b, 0x5c, 0x72, 0x74, 0x66, 0x31]))).toBe('rtf')
  })

  it('returns unknown for non-matching buffers and short inputs', () => {
    expect(detectByMagic(bytesToBuffer([0x00, 0x01, 0x02, 0x03]))).toBe('unknown')
    expect(detectByMagic(bytesToBuffer([0x25, 0x50, 0x44]))).toBe('unknown')
  })
})

describe('detectFormat', () => {
  it('returns extension detection immediately when buffer is omitted', () => {
    expect(detectFormat('report.docx')).toBe('docx')
    expect(detectFormat('mystery.bin')).toBe('unknown')
  })

  it('uses magic detection when extension is unknown', () => {
    expect(detectFormat('archive.bin', zipWithAscii('word/document.xml'))).toBe('docx')
  })

  it('keeps extension for text-class formats when magic disagrees', () => {
    expect(detectFormat('notes.txt', zipWithAscii('word/document.xml'))).toBe('text')
    expect(detectFormat('script.ts', bytesToBuffer([0x25, 0x50, 0x44, 0x46]))).toBe('code')
  })

  it('trusts magic for document-class mismatches', () => {
    expect(detectFormat('report.pdf', bytesToBuffer([0x7b, 0x5c, 0x72, 0x74, 0x66]))).toBe('rtf')
    expect(detectFormat('report.docx', zipWithAscii('xl/workbook.xml'))).toBe('xlsx')
    expect(detectFormat('slides.odp', zipWithAscii('mimetypeapplication/vnd.oasis.opendocument.text'))).toBe('odt')
  })

  it('keeps extension when magic is unknown', () => {
    expect(detectFormat('report.pdf', zipWithAscii('plain/archive.bin'))).toBe('pdf')
  })
})

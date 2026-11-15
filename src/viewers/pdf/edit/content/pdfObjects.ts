/**
 * Small helpers over pdf-lib's object model (PDF-EDIT-2). pdf-lib is passed
 * in rather than imported so the viewer can keep loading it on demand.
 */

import type { PDFContext, PDFDict, PDFObject } from 'pdf-lib'

import type { PdfLib } from '../writePdf'
import { bytesToBinaryString } from './lexer'

export function resolve(lib: PdfLib, context: PDFContext, value: PDFObject | undefined): PDFObject | undefined {
  return value instanceof lib.PDFRef ? context.lookup(value) : value
}

export function asDict(lib: PdfLib, context: PDFContext, value: PDFObject | undefined): PDFDict | undefined {
  const resolved = resolve(lib, context, value)
  if (resolved instanceof lib.PDFDict) return resolved
  if (resolved instanceof lib.PDFStream) return resolved.dict
  return undefined
}

export function dictGet(lib: PdfLib, context: PDFContext, dict: PDFDict | undefined, key: string): PDFObject | undefined {
  return dict ? resolve(lib, context, dict.get(lib.PDFName.of(key))) : undefined
}

export function nameOf(lib: PdfLib, value: PDFObject | undefined): string | undefined {
  return value instanceof lib.PDFName ? value.decodeText() : undefined
}

export function numberOf(lib: PdfLib, value: PDFObject | undefined): number | undefined {
  return value instanceof lib.PDFNumber ? value.asNumber() : undefined
}

export function numbersOf(lib: PdfLib, context: PDFContext, value: PDFObject | undefined): number[] | undefined {
  const resolved = resolve(lib, context, value)
  if (!(resolved instanceof lib.PDFArray)) return undefined
  return resolved.asArray().map((item) => numberOf(lib, resolve(lib, context, item)) ?? 0)
}

/** A stream's decoded bytes as a binary string, or null for a filter pdf-lib cannot decode. */
export function streamText(lib: PdfLib, stream: PDFObject | undefined): string | null {
  try {
    if (stream instanceof lib.PDFRawStream) return bytesToBinaryString(lib.decodePDFRawStream(stream).decode())
    if (stream && 'getUnencodedContents' in stream && typeof stream.getUnencodedContents === 'function') {
      return bytesToBinaryString((stream as { getUnencodedContents(): Uint8Array }).getUnencodedContents())
    }
    if (stream instanceof lib.PDFStream) return bytesToBinaryString(stream.getContents())
  } catch {
    return null
  }
  return null
}

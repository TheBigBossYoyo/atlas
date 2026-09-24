export function decodeTextBuffer(buffer: Buffer): string;
export function isValidUtf8(buffer: Buffer): boolean;
export function decodeWindows1252(buffer: Buffer): string;

export type TextEncodingName = 'utf-8' | 'utf-16le' | 'utf-16be' | 'windows-1252';
export type NewlineStyle = 'crlf' | 'lf';

export interface TextFileMeta {
  readonly encoding: TextEncodingName;
  readonly bom: boolean;
  readonly newline: NewlineStyle;
}

export function decodeTextBufferWithMeta(buffer: Buffer): { content: string; meta: TextFileMeta };
export function encodeTextBuffer(
  content: string,
  meta: TextFileMeta,
): { buffer: Buffer; encodingFallback: boolean };
export function detectNewline(content: string): NewlineStyle;
export function normalizeNewlines(content: string): string;

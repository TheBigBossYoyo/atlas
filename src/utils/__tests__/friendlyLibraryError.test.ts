import { describe, expect, it } from 'vitest';
import { toFriendlyError } from '../friendlyLibraryError';

describe('toFriendlyError', () => {
  it.each([
    ["Can't find end of central directory", 'the file could not be read as a valid Office document (it may be corrupted or not a real Office file)'],
    ['File is not a valid zip file', 'the file could not be read as a valid Office document (it may be corrupted or not a real Office file)'],
    ['Corrupted zip: bad CRC', 'the file could not be read as a valid Office document (it may be corrupted or not a real Office file)'],
    ['Invalid signature detected', 'the file is not in a format this app recognizes'],
    ['Unsupported compression method', 'the file is not in a format this app recognizes'],
    ['Out of memory', 'the document is too large to process'],
    ['Allocation failed while parsing', 'the document is too large to process'],
    ['Maximum call stack size exceeded', 'the document is too large to process'],
    ['Invalid PDF structure', 'the PDF file appears to be corrupted or malformed'],
    ['Bad XREF entry', 'the PDF file appears to be corrupted or malformed'],
    ['This document is password protected', 'the file is password-protected and cannot be opened'],
    ['The document is encrypted', 'the file is password-protected and cannot be opened'],
  ])('maps "%s" to a plain-language explanation', (raw, expectedFriendly) => {
    const err = toFriendlyError(new Error(raw), 'DOCX export failed');
    expect(err.message).toBe(`DOCX export failed: ${expectedFriendly}`);
  });

  it('is case-insensitive when matching known patterns', () => {
    const err = toFriendlyError(new Error('CENTRAL DIRECTORY not found'), 'Opening this file');
    expect(err.message).toContain('valid Office document');
  });

  it('falls back to the raw message, unaltered, for an unrecognized error', () => {
    const err = toFriendlyError(new Error('some totally novel failure'), 'Export failed');
    expect(err.message).toBe('Export failed: some totally novel failure');
  });

  it('stringifies a non-Error throw instead of hiding it', () => {
    const err = toFriendlyError('a plain string was thrown', 'Export failed');
    expect(err.message).toBe('Export failed: a plain string was thrown');
  });

  it('always returns a real Error instance', () => {
    expect(toFriendlyError(new Error('x'), 'ctx')).toBeInstanceOf(Error);
    expect(toFriendlyError('x', 'ctx')).toBeInstanceOf(Error);
  });
});

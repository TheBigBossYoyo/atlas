import type { DocStats } from '../types';

const WORDS_PER_MINUTE = 220;

/** Compute lightweight document statistics. Pure & memo-friendly. */
export function computeStats(markdown: string): DocStats {
  const characters = markdown.length;
  const charactersNoSpaces = markdown.replace(/\s+/g, '').length;
  const words = markdown.trim() === '' ? 0 : markdown.trim().split(/\s+/).length;
  const lines = markdown === '' ? 0 : markdown.split(/\r?\n/).length;
  const readingMinutes = Math.max(1, Math.ceil(words / WORDS_PER_MINUTE));
  return { characters, charactersNoSpaces, words, lines, readingMinutes };
}

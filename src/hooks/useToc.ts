import { useMemo } from 'react';
import type { TocItem } from '../types';

function slugify(text: string): string {
  return text
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .trim();
}

export function useToc(markdown: string): TocItem[] {
  return useMemo(() => {
    if (!markdown) return [];

    const headingRegex = /^(#{1,6})\s+(.+)$/gm;
    const items: TocItem[] = [];
    let match: RegExpExecArray | null;

    while ((match = headingRegex.exec(markdown)) !== null) {
      const level = match[1].length;
      const raw = match[2].trim();
      // Strip inline markdown formatting
      const text = raw
        .replace(/\*\*(.*?)\*\*/g, '$1')
        .replace(/\*(.*?)\*/g, '$1')
        .replace(/`(.*?)`/g, '$1')
        .replace(/\[(.*?)\]\(.*?\)/g, '$1')
        .replace(/~~(.*?)~~/g, '$1');

      items.push({
        id: slugify(text),
        text,
        level,
      });
    }

    return items;
  }, [markdown]);
}

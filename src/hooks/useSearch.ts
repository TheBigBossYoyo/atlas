import { useState, useCallback, useEffect, useRef } from 'react';
import { useShellShortcut } from './useShortcutManager';
import type { ShortcutHandler } from './shortcutManagerContext';

/**
 * `enabled` gates in-app search entirely (P2.1/X4/DAT-15/RUN-08): pass
 * `true` for any format whose DOM structure this TreeWalker-based search can
 * actually walk (markdown, Text/Code/RTF/ODT), and `false` for everything
 * else (PDF, spreadsheets, slide decks) — the search shortcut simply isn't
 * registered at all in that case, so Ctrl+F is never intercepted for those
 * formats and falls through to whatever (if anything) the viewer or the
 * platform does with it, instead of the old behavior of unconditionally
 * calling `preventDefault()` with no feature behind it (RUN-08).
 */
export function useSearch(contentRef: React.RefObject<HTMLElement | null>, enabled = true) {
  const [isOpen, setIsOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [matchCount, setMatchCount] = useState(0);
  const [currentMatch, setCurrentMatch] = useState(0);
  const markedNodes = useRef<HTMLElement[]>([]);

  const clearHighlights = useCallback(() => {
    for (const mark of markedNodes.current) {
      const parent = mark.parentNode;
      if (parent) {
        parent.replaceChild(document.createTextNode(mark.textContent || ''), mark);
        parent.normalize();
      }
    }
    markedNodes.current = [];
    setMatchCount(0);
    setCurrentMatch(0);
  }, []);

  const highlightMatches = useCallback((searchQuery: string) => {
    clearHighlights();
    if (!searchQuery || !contentRef.current) return;

    const walker = document.createTreeWalker(
      contentRef.current,
      NodeFilter.SHOW_TEXT,
      null,
    );

    const textNodes: Text[] = [];
    let node: Node | null;
    while ((node = walker.nextNode())) {
      textNodes.push(node as Text);
    }

    const lowerQuery = searchQuery.toLowerCase();
    const marks: HTMLElement[] = [];

    for (const textNode of textNodes) {
      const text = textNode.textContent || '';
      const lowerText = text.toLowerCase();
      let startIdx = 0;
      const positions: number[] = [];

      while (true) {
        const idx = lowerText.indexOf(lowerQuery, startIdx);
        if (idx === -1) break;
        positions.push(idx);
        startIdx = idx + 1;
      }

      if (positions.length === 0) continue;

      const fragment = document.createDocumentFragment();
      let lastEnd = 0;

      for (const pos of positions) {
        if (pos > lastEnd) {
          fragment.appendChild(document.createTextNode(text.slice(lastEnd, pos)));
        }
        const mark = document.createElement('mark');
        mark.className = 'search-highlight';
        mark.textContent = text.slice(pos, pos + searchQuery.length);
        fragment.appendChild(mark);
        marks.push(mark);
        lastEnd = pos + searchQuery.length;
      }

      if (lastEnd < text.length) {
        fragment.appendChild(document.createTextNode(text.slice(lastEnd)));
      }

      textNode.parentNode?.replaceChild(fragment, textNode);
    }

    markedNodes.current = marks;
    setMatchCount(marks.length);
    if (marks.length > 0) {
      setCurrentMatch(0);
      marks[0].classList.add('search-highlight--active');
      marks[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }, [clearHighlights, contentRef]);

  const goToMatch = useCallback((direction: 'next' | 'prev') => {
    const marks = markedNodes.current;
    if (marks.length === 0) return;

    marks[currentMatch]?.classList.remove('search-highlight--active');

    let next: number;
    if (direction === 'next') {
      next = (currentMatch + 1) % marks.length;
    } else {
      next = (currentMatch - 1 + marks.length) % marks.length;
    }

    setCurrentMatch(next);
    marks[next]?.classList.add('search-highlight--active');
    marks[next]?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }, [currentMatch]);

  const open = useCallback(() => setIsOpen(true), []);

  const close = useCallback(() => {
    setIsOpen(false);
    setQuery('');
    clearHighlights();
  }, [clearHighlights]);

  // P2.1/X4/RUN-08 — registered only while `enabled`; the dispatcher simply
  // never calls this handler for a disabled format, so Ctrl+F is never
  // preventDefault()-ed with nothing behind it (the old "disabled branch"
  // bug).
  const handler = useCallback<ShortcutHandler>(
    (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
        e.preventDefault();
        if (isOpen) {
          close();
        } else {
          open();
        }
        return true;
      }
      if (e.key === 'Escape' && isOpen) {
        close();
        return true;
      }
      return false;
    },
    [close, isOpen, open],
  );

  useShellShortcut(handler, enabled);

  // Re-highlight on query change
  useEffect(() => {
    if (!enabled) return;

    const timer = setTimeout(() => {
      highlightMatches(query);
    }, 200);
    return () => clearTimeout(timer);
  }, [enabled, highlightMatches, query]);

  return {
    isOpen,
    query,
    setQuery,
    matchCount,
    currentMatch,
    goToMatch,
    open,
    close,
  } as const;
}

/**
 * useToc — characterization tests (P0.4 / QA-07).
 *
 * Locks in today's heading-extraction + slug-id behavior against the same
 * fixture set MarkdownRenderer.characterization.test.tsx uses, so a later
 * wave's shared-hook refactor can prove it left this hook's output
 * byte-identical (see the improvement plan's Section 7 guardrail).
 *
 * Note (frozen quirk, not a bug to fix here): `useToc`'s own `slugify` never
 * de-duplicates repeated slugs, unlike `rehype-slug` in MarkdownRenderer
 * (which appends "-1", "-2", ...). Two headings with identical text
 * therefore get the *same* TOC id today — asserted explicitly below so a
 * future fix to this inconsistency is a deliberate, reviewed change rather
 * than an invisible one.
 */

import { describe, expect, it } from 'vitest';
import { renderHook } from '@testing-library/react';
import { useToc } from '../useToc';

import headingsFixture from '../../__tests__/fixtures/markdown/headings.md?raw';
import gfmTableFixture from '../../__tests__/fixtures/markdown/gfm-table.md?raw';
import kitchenSinkFixture from '../../__tests__/fixtures/markdown/kitchen-sink.md?raw';

describe('useToc characterization', () => {
  it('returns an empty array for empty markdown', () => {
    const { result } = renderHook(() => useToc(''));
    expect(result.current).toEqual([]);
  });

  it('extracts headings h1-h6 from the "headings" fixture, stripping inline formatting from the display text', () => {
    const { result } = renderHook(() => useToc(headingsFixture));

    expect(result.current).toEqual([
      { id: 'heading-one', text: 'Heading One', level: 1 },
      { id: 'heading-two-with-bold', text: 'Heading Two with Bold', level: 2 },
      { id: 'heading-three-with-code', text: 'Heading Three with Code', level: 3 },
      { id: 'heading-four-with-italic', text: 'Heading Four with Italic', level: 4 },
      { id: 'heading-five-with-link', text: 'Heading Five with Link', level: 5 },
      { id: 'heading-six-with-strikethrough', text: 'Heading Six with Strikethrough', level: 6 },
      { id: 'heading-two-with-bold', text: 'Heading Two with Bold', level: 2 },
    ]);
  });

  it('does not de-duplicate repeated slugs (unlike rehype-slug in MarkdownRenderer)', () => {
    const { result } = renderHook(() => useToc(headingsFixture));
    const duplicates = result.current.filter((item) => item.id === 'heading-two-with-bold');
    expect(duplicates).toHaveLength(2);
  });

  it('only extracts actual headings from GFM content (table rows/cells are not headings)', () => {
    const { result } = renderHook(() => useToc(gfmTableFixture));
    expect(result.current).toEqual([{ id: 'team-roster', text: 'Team Roster', level: 2 }]);
  });

  it('matches a full snapshot of heading extraction over the "kitchen-sink" fixture', () => {
    const { result } = renderHook(() => useToc(kitchenSinkFixture));
    expect(result.current).toMatchSnapshot();
  });

  it('is memoized: the same markdown string reference returns the same array instance', () => {
    const { result, rerender } = renderHook(({ md }: { md: string }) => useToc(md), {
      initialProps: { md: headingsFixture },
    });
    const first = result.current;
    rerender({ md: headingsFixture });
    expect(result.current).toBe(first);
  });

  it('recomputes when the markdown content changes', () => {
    const { result, rerender } = renderHook(({ md }: { md: string }) => useToc(md), {
      initialProps: { md: '# One' },
    });
    expect(result.current).toEqual([{ id: 'one', text: 'One', level: 1 }]);

    rerender({ md: '# One\n## Two' });
    expect(result.current).toEqual([
      { id: 'one', text: 'One', level: 1 },
      { id: 'two', text: 'Two', level: 2 },
    ]);
  });
});

/**
 * Found by driving the real app with a broken mermaid diagram: mermaid's own
 * `render()` builds its error graphic inside a temporary `div#d<id>` it
 * appends directly to `document.body` for a live DOM context, and only
 * removes that temp element on the SUCCESS path — a rejected `render()`
 * (a parse error) leaves it behind, visible and full-sized, as a sibling of
 * the whole app for the rest of the window's lifetime (outside React's own
 * tree, so no remount/unmount ever clears it). These tests fake mermaid's
 * real "leave a `div#d<id>` behind on failure" behavior (rather than
 * depending on the real, slow `mermaid` package) to pin down that `Mermaid`
 * itself cleans it up.
 */
import { act, render, screen } from '@testing-library/react';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { Mermaid } from '../Mermaid';

vi.mock('mermaid', () => {
  return {
    default: {
      initialize: vi.fn(),
      render: vi.fn(),
    },
  };
});

afterEach(() => {
  vi.restoreAllMocks();
  document.body.querySelectorAll('[id^="d"]').forEach((el) => {
    if (el.tagName === 'DIV' && el.id.startsWith('dtest-')) el.remove();
  });
});

describe('Mermaid', () => {
  it('renders the resolved SVG on success', async () => {
    const mermaid = (await import('mermaid')).default;
    vi.mocked(mermaid.render).mockResolvedValue({ svg: '<svg data-testid="ok-svg"></svg>', diagramType: 'flowchart' } as never);

    const { container } = render(<Mermaid code="graph TD; A-->B" id="test-ok" />);
    await act(async () => {});

    expect(container.querySelector('[data-testid="ok-svg"]')).not.toBeNull();
  });

  it('shows the inline error AND removes mermaid\'s own leftover temp DOM node on a parse failure', async () => {
    const mermaid = (await import('mermaid')).default;
    vi.mocked(mermaid.render).mockImplementation(async (id: string) => {
      // Mimics the real library: a temp host div is appended to
      // document.body while it works, and is still there when it rejects.
      const leftover = document.createElement('div');
      leftover.id = `d${id}`;
      leftover.textContent = 'mermaid error graphic';
      document.body.appendChild(leftover);
      throw new Error('Parse error on line 1');
    });

    render(<Mermaid code="not valid mermaid" id="test-broken" />);
    await act(async () => {});

    expect(screen.getByText(/Parse error on line 1/)).toBeInTheDocument();
    expect(document.getElementById('dtest-broken')).toBeNull();
  });
});

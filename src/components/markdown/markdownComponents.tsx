/**
 * Builds the `components` map both `MarkdownRenderer` branches (the inline
 * `<ReactMarkdown>` path and `WorkerMarkdownBody`'s `renderHastTree` path)
 * render with. Kept in its own module — not inline in `MarkdownRenderer.tsx`
 * — for two reasons: `react-refresh/only-export-components` requires a
 * component file to export only components, and, more importantly, both
 * rendering branches (and `renderHastTree.parity.test.tsx`, which compares
 * them) need to build the exact same `components` value with no chance of
 * the two drifting apart.
 */
import type { Components } from 'react-markdown';
import { Mermaid } from '../Mermaid';
import { HighlightedText } from './HighlightedText';

export function buildMarkdownComponents(searchQuery: string | undefined, baseId: string): Partial<Components> {
  const wrapText = (children: React.ReactNode): React.ReactNode => {
    if (!searchQuery) return children;
    if (typeof children === 'string') {
      return <HighlightedText text={children} query={searchQuery} />;
    }
    if (Array.isArray(children)) {
      return children.map((child, i) => (
        <span key={i}>{wrapText(child)}</span>
      ));
    }
    return children;
  };

  const components: Partial<Components> = {
    code({ className, children, ...props }) {
      const match = /language-(\w+)/.exec(className ?? '');
      const lang = match?.[1];
      const raw = String(children ?? '').replace(/\n$/, '');
      if (lang === 'mermaid') {
        // Stable id per mermaid block based on content hash & baseId
        const hashed = `mermaid-${baseId.replace(/[^a-zA-Z0-9]/g, '')}-${Math.abs(
          raw.split('').reduce((acc, ch) => ((acc << 5) - acc + ch.charCodeAt(0)) | 0, 0)
        )}`;
        return <Mermaid code={raw} id={hashed} />;
      }
      return (
        <code className={className} {...props}>
          {children}
        </code>
      );
    },
  };

  if (searchQuery) {
    components.p = ({ children, ...props }) => <p {...props}>{wrapText(children)}</p>;
    components.li = ({ children, ...props }) => <li {...props}>{wrapText(children)}</li>;
    components.td = ({ children, ...props }) => <td {...props}>{wrapText(children)}</td>;
    components.th = ({ children, ...props }) => <th {...props}>{wrapText(children)}</th>;
  }

  return components;
}

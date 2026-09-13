import { forwardRef, useCallback, useMemo, useId } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import rehypeKatex from 'rehype-katex';
import rehypeHighlight from 'rehype-highlight';
import rehypeSlug from 'rehype-slug';
import rehypeRaw from 'rehype-raw';
import type { Components } from 'react-markdown';
import { Mermaid } from './Mermaid';

interface MarkdownRendererProps {
  markdown: string;
  searchQuery?: string;
}

function HighlightedText({ text, query }: { text: string; query: string }) {
  if (!query) return <>{text}</>;

  const parts: { text: string; highlighted: boolean }[] = [];
  const lowerText = text.toLowerCase();
  const lowerQuery = query.toLowerCase();
  let lastEnd = 0;

  let startIdx = 0;
  while (true) {
    const idx = lowerText.indexOf(lowerQuery, startIdx);
    if (idx === -1) break;
    if (idx > lastEnd) {
      parts.push({ text: text.slice(lastEnd, idx), highlighted: false });
    }
    parts.push({ text: text.slice(idx, idx + query.length), highlighted: true });
    lastEnd = idx + query.length;
    startIdx = idx + 1;
  }

  if (lastEnd < text.length) {
    parts.push({ text: text.slice(lastEnd), highlighted: false });
  }

  if (parts.length === 0) return <>{text}</>;

  return (
    <>
      {parts.map((part, i) =>
        part.highlighted ? (
          <mark key={i} className="search-highlight">{part.text}</mark>
        ) : (
          <span key={i}>{part.text}</span>
        )
      )}
    </>
  );
}

export const MarkdownRenderer = forwardRef<HTMLDivElement, MarkdownRendererProps>(
  function MarkdownRenderer({ markdown, searchQuery }, ref) {
    const baseId = useId();
    const remarkPlugins = useMemo(() => [remarkGfm, remarkMath], []);
    const rehypePlugins = useMemo(() => [rehypeKatex, rehypeHighlight, rehypeSlug, rehypeRaw], []);

    const createComponents = useCallback((): Partial<Components> => {
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
    }, [searchQuery, baseId]);

    return (
      <div ref={ref} className="markdown-body" id="markdown-content">
        <ReactMarkdown
          remarkPlugins={remarkPlugins}
          rehypePlugins={rehypePlugins}
          components={createComponents()}
        >
          {markdown}
        </ReactMarkdown>
      </div>
    );
  }
);

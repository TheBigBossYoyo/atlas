/**
 * Wraps every case-insensitive occurrence of `query` in `text` with a
 * `<mark>`. Split out of `markdownComponents.tsx` so that file can export
 * only its `buildMarkdownComponents` factory without also defining a
 * component inline (`react-refresh/only-export-components` flags a file
 * that does both, even when the component itself isn't exported).
 */
export function HighlightedText({ text, query }: { text: string; query: string }) {
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

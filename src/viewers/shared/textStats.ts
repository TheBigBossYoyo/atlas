/**
 * Shared word-count helper for RtfViewer/OdtViewer.
 *
 * `innerText` is layout-aware (collapses whitespace the way a real reader
 * would see it) but jsdom — this project's test environment — does not
 * implement it at all (no layout engine, so the property is `undefined`
 * rather than an empty string). `textContent` is always present but
 * includes text a real browser's `innerText` would hide (e.g.
 * `display:none` content), which doesn't matter here since neither
 * viewer's sanitized output includes hidden text. Falling back to
 * `textContent` keeps real-browser behavior unchanged while making word
 * counting exercisable under jsdom.
 */
export function countWords(container: Pick<HTMLElement, 'innerText' | 'textContent'>): number {
  const text = container.innerText ?? container.textContent ?? ''
  return text.trim().split(/\s+/).filter(Boolean).length
}

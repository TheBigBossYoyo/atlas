/**
 * A11Y-pass-3 — `Element.scrollIntoView({ behavior: 'smooth' })` is a
 * per-call option, not a CSS property: it overrides the element's CSS
 * `scroll-behavior` outright, so the app-wide
 * `@media (prefers-reduced-motion: reduce) { *, *::before, *::after {
 * scroll-behavior: auto !important } }` rule in `index.css` never actually
 * reaches any of the several `scrollIntoView({ behavior: 'smooth' })` call
 * sites across the app (Find in markdown/PDF, jumping to a PDF/DOCX page,
 * a DOCX bookmark/heading nav item). A user with reduced motion enabled at
 * the OS level got full animated scrolling on every one of those regardless.
 *
 * `scrollIntoViewRespectingMotionPreference` is a drop-in replacement for a
 * bare `element.scrollIntoView(options)` call: identical behavior, except it
 * downgrades `behavior: 'smooth'` to `'auto'` (instant) whenever the media
 * query matches.
 */

/** Re-read on every call (rather than cached) since the user can flip the OS
 * setting while the app is open, and this is never called often enough
 * (only on an explicit scroll/jump) to make a fresh `matchMedia` call a
 * measurable cost. */
export function prefersReducedMotion(): boolean {
  return (
    typeof window !== 'undefined' &&
    typeof window.matchMedia === 'function' &&
    window.matchMedia('(prefers-reduced-motion: reduce)').matches
  );
}

/**
 * Same signature as `Element.scrollIntoView`, but honors the OS-level
 * reduced-motion preference by forcing `behavior: 'auto'` when it's set,
 * regardless of what the caller asked for.
 */
export function scrollIntoViewRespectingMotionPreference(
  element: Element,
  options?: ScrollIntoViewOptions,
): void {
  if (!prefersReducedMotion() || options === undefined) {
    element.scrollIntoView(options);
    return;
  }
  element.scrollIntoView({ ...options, behavior: 'auto' });
}

/**
 * sanitizeExportHtml.ts — X5. The concrete acceptance test the plan asks
 * for: "test that <script> and on* handlers in markdown raw HTML don't
 * survive export."
 */
import { describe, expect, it } from 'vitest';
import { sanitizeExportHtml } from '../sanitizeExportHtml';

describe('sanitizeExportHtml', () => {
  it('strips a <script> tag and its content', () => {
    const out = sanitizeExportHtml('<p>hi</p><script>window.pwned = true</script>');
    expect(out).not.toContain('<script');
    expect(out).not.toContain('pwned');
    expect(out).toContain('<p>hi</p>');
  });

  it('strips onerror/onclick/onload event handler attributes', () => {
    const out = sanitizeExportHtml(
      '<img src="x" onerror="pwn()"><button onclick="pwn()">go</button><body onload="pwn()">',
    );
    expect(out).not.toContain('onerror');
    expect(out).not.toContain('onclick');
    expect(out).not.toContain('onload');
  });

  it('strips a javascript: href on a hyperlink (e.g. from a converted RTF file)', () => {
    const out = sanitizeExportHtml('<a href="javascript:alert(1)">click</a>');
    expect(out).not.toContain('javascript:');
  });

  it('strips an embedded <style>/<iframe>/<object>/<embed>/<link>/<base>', () => {
    const out = sanitizeExportHtml(
      '<style>body{display:none}</style><iframe src="evil"></iframe><object data="evil"></object>' +
        '<embed src="evil"><link rel="stylesheet" href="evil"><base href="evil">',
    );
    expect(out).not.toContain('<style');
    expect(out).not.toContain('<iframe');
    expect(out).not.toContain('<object');
    expect(out).not.toContain('<embed');
    expect(out).not.toContain('<link');
    expect(out).not.toContain('<base');
  });

  it('preserves benign HTML content and structure (raw-html.md fixture-shaped content)', () => {
    const out = sanitizeExportHtml('<div class="callout"><strong>Note:</strong> hi</div>');
    expect(out).toContain('class="callout"');
    expect(out).toContain('<strong>Note:</strong>');
  });

  it('preserves inline styles and data attributes needed for DOCX/table layout fidelity', () => {
    const out = sanitizeExportHtml('<div class="docx-page" style="width: 816px;" data-size-key="816x1056">x</div>');
    expect(out).toContain('style="width: 816px;"');
    expect(out).toContain('data-size-key="816x1056"');
  });

  it('preserves a Mermaid-rendered inline <svg> and its <foreignObject> element', () => {
    // DOMPurify strips HTML *content* nested inside an SVG `foreignObject`
    // regardless of `ADD_TAGS` (a defense against a known mutation-XSS
    // bypass class that historically abused exactly this namespace-crossing
    // pattern) — an accepted trade-off, not a bug: the `<svg>`/plain SVG
    // shape markup Mermaid renders for the vast majority of diagram types
    // survives untouched (see the plain `<svg>` case covered elsewhere via
    // the "math"/"mermaid" markdown-export snapshot fixtures).
    const out = sanitizeExportHtml('<svg><foreignObject><div class="label">flow step</div></foreignObject></svg>');
    expect(out).toContain('<svg');
    expect(out).toContain('foreignObject');
  });

  it('preserves KaTeX MathML, including <semantics>/<annotation> (screen-reader/copy-as-LaTeX source)', () => {
    const out = sanitizeExportHtml(
      '<math><semantics><mrow><mi>x</mi></mrow><annotation encoding="application/x-tex">x</annotation></semantics></math>',
    );
    expect(out).toContain('<semantics>');
    expect(out).toContain('<annotation encoding="application/x-tex">x</annotation>');
  });

  it('does NOT allow the historically-dangerous annotation-xml tag through', () => {
    const out = sanitizeExportHtml('<math><annotation-xml encoding="text/html"><script>pwn()</script></annotation-xml></math>');
    expect(out).not.toContain('annotation-xml');
    expect(out).not.toContain('pwn()');
  });
});

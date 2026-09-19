// Content-Security-Policy for the main window (ELEC-04).
//
// Dev needs 'unsafe-eval' (Vite/React Fast Refresh) and the HMR websocket;
// prod is stricter but must still allow what the viewers genuinely need:
//   - the pdf.js module worker (worker-src 'self'/blob:)
//   - shiki's WASM tokenizer for CodeViewer ('wasm-unsafe-eval')
//   - blob:/data: images (DOCX inline images, canvas-rendered pages)
//   - blob:/data: fonts (bundled DOCX substitute fonts, embedded fonts)
//   - inline styles (KaTeX/mermaid inject <style>/style attributes)
//   - https: images/fonts — markdown's `rehypeRaw` pass-through and plain
//     `![]()` syntax both regularly reference remote-hosted images (README
//     badges, hotlinked screenshots) and, occasionally, remote @font-face
//     URLs. The owner's hard constraint is that markdown's behavior must not
//     change; blocking these here would silently turn previously-loading
//     remote images into broken-image icons in production, which the
//     fixture-based verification (all-local fixtures) cannot catch. `https:`
//     only — not `http:` — to avoid reopening a plaintext-network/SSRF-style
//     probe surface for internal-network URLs a malicious document could
//     embed, which the pre-CSP baseline already carried but which a hard
//     CSP add is a natural place to not re-widen.

/**
 * @param {boolean} isDev
 * @returns {string}
 */
function buildContentSecurityPolicy(isDev) {
  const scriptSrc = isDev
    ? ["'self'", "'unsafe-inline'", "'unsafe-eval'"]
    : ["'self'", "'wasm-unsafe-eval'"];

  const connectSrc = isDev
    ? ["'self'", 'ws://localhost:5173', 'http://localhost:5173']
    : ["'self'"];

  /** @type {Array<[string, string[]]>} */
  const directives = [
    ['default-src', ["'self'"]],
    ['script-src', scriptSrc],
    ['worker-src', ["'self'", 'blob:']],
    ['style-src', ["'self'", "'unsafe-inline'"]],
    ['img-src', ["'self'", 'data:', 'blob:', 'https:']],
    ['font-src', ["'self'", 'data:', 'blob:', 'https:']],
    ['connect-src', connectSrc],
    ['object-src', ["'none'"]],
    ['base-uri', ["'self'"]],
  ];

  return directives.map(([name, values]) => `${name} ${values.join(' ')}`).join('; ');
}

module.exports = { buildContentSecurityPolicy };

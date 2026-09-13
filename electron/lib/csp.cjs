// Content-Security-Policy for the main window (ELEC-04).
//
// Dev needs 'unsafe-eval' (Vite/React Fast Refresh) and the HMR websocket;
// prod is stricter but must still allow what the viewers genuinely need:
//   - the pdf.js module worker (worker-src 'self'/blob:)
//   - shiki's WASM tokenizer for CodeViewer ('wasm-unsafe-eval')
//   - blob:/data: images (DOCX inline images, canvas-rendered pages)
//   - blob:/data: fonts (bundled DOCX substitute fonts, embedded fonts)
//   - inline styles (KaTeX/mermaid inject <style>/style attributes)

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

  const directives = [
    ['default-src', ["'self'"]],
    ['script-src', scriptSrc],
    ['worker-src', ["'self'", 'blob:']],
    ['style-src', ["'self'", "'unsafe-inline'"]],
    ['img-src', ["'self'", 'data:', 'blob:']],
    ['font-src', ["'self'", 'data:', 'blob:']],
    ['connect-src', connectSrc],
    ['object-src', ["'none'"]],
    ['base-uri', ["'self'"]],
  ];

  return directives.map(([name, values]) => `${name} ${values.join(' ')}`).join('; ');
}

module.exports = { buildContentSecurityPolicy };

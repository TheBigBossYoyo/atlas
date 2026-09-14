/**
 * X2 / UX-05 — rasterizes rendered math (KaTeX) and Mermaid diagrams into PNG
 * images for DOCX export. Word has no live LaTeX/Mermaid renderer, so the
 * only way to make these show up as anything but raw source text is to embed
 * a picture of the already-rendered output — matching what the live preview
 * (`MarkdownRenderer` + `Mermaid.tsx`, both out of scope to touch per the
 * plan's markdown guardrails) already shows on screen.
 *
 * Both `katex` and `mermaid` are dynamically imported so this rasterization
 * path only runs when a document actually contains math or a Mermaid fence.
 * `mermaid` is genuinely lazy-loaded this way (Vite code-splits it into its
 * own chunk, per the build output). `katex` is a transitive dependency (via
 * `rehype-katex`, used by the live preview) that's already statically
 * bundled into the main chunk regardless — the dynamic import here doesn't
 * additionally defer *loading* it, but it renders synchronously to a plain
 * HTML string with no engine of its own, so importing it directly for this
 * one rasterization step is still simple and safe.
 */
import html2canvas from 'html2canvas-pro';

export interface RasterImage {
  readonly bytes: Uint8Array;
  readonly width: number;
  readonly height: number;
}

const RASTER_SCALE = 2;

function dataUrlToBytes(dataUrl: string): Uint8Array | null {
  const commaIndex = dataUrl.indexOf(',');
  if (commaIndex === -1) return null;
  const base64 = dataUrl.slice(commaIndex + 1);
  if (!base64) return null;
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/** Mounts `html` off-screen (white background, so transparent glyphs stay legible), rasterizes it, then tears the scratch element down. */
async function rasterizeHtml(html: string): Promise<RasterImage | null> {
  const host = document.createElement('div');
  host.style.position = 'fixed';
  host.style.left = '-99999px';
  host.style.top = '0';
  host.style.background = '#ffffff';
  host.innerHTML = html;
  document.body.appendChild(host);
  try {
    const canvas = await html2canvas(host, {
      scale: RASTER_SCALE,
      backgroundColor: '#ffffff',
      logging: false,
    });
    if (!canvas.width || !canvas.height) return null;
    const dataUrl = canvas.toDataURL('image/png');
    const bytes = dataUrlToBytes(dataUrl);
    if (!bytes) return null;
    return {
      bytes,
      width: Math.max(1, Math.round(canvas.width / RASTER_SCALE)),
      height: Math.max(1, Math.round(canvas.height / RASTER_SCALE)),
    };
  } catch {
    return null;
  } finally {
    host.remove();
  }
}

/**
 * Renders a LaTeX expression to a PNG via KaTeX for DOCX embedding.
 * Resolves `null` on render failure so the caller can fall back to plain text
 * instead of failing the whole export over one bad equation.
 */
export async function renderMathImage(tex: string, displayMode: boolean): Promise<RasterImage | null> {
  try {
    const katex = (await import('katex')).default;
    const html = katex.renderToString(tex, { throwOnError: false, displayMode, output: 'html' });
    const fontSize = displayMode ? 24 : 18;
    return await rasterizeHtml(
      `<span style="font-size:${fontSize}px;display:inline-block;padding:6px;color:#000000;">${html}</span>`,
    );
  } catch {
    return null;
  }
}

let mermaidRenderCounter = 0;

/**
 * Renders Mermaid diagram source to a PNG for DOCX embedding.
 * Resolves `null` on render failure so the caller can fall back to a plain
 * code block instead of failing the whole export over one bad diagram.
 */
export async function renderMermaidImage(code: string): Promise<RasterImage | null> {
  try {
    const mermaid = (await import('mermaid')).default;
    mermaid.initialize({ startOnLoad: false, theme: 'default', securityLevel: 'loose' });
    mermaidRenderCounter += 1;
    const { svg } = await mermaid.render(`docx-export-mermaid-${mermaidRenderCounter}`, code);
    return await rasterizeHtml(svg);
  } catch {
    return null;
  }
}

/** Caps an image's width for a standard A4 page (with 1in margins) at 96dpi, scaling height proportionally. */
export function scaleToMaxWidth(
  width: number,
  height: number,
  maxWidth = 550,
): { readonly width: number; readonly height: number } {
  if (width <= maxWidth) return { width, height };
  const ratio = maxWidth / width;
  return { width: maxWidth, height: Math.max(1, Math.round(height * ratio)) };
}

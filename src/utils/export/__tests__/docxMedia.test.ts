import { afterEach, describe, expect, it, vi } from 'vitest';
import { scaleToMaxWidth } from '../docxMedia';

const { html2canvasMock } = vi.hoisted(() => ({
  html2canvasMock: vi.fn(),
}));

vi.mock('html2canvas-pro', () => ({ default: html2canvasMock }));

// docxMedia.ts dynamically `import('mermaid')`s — vi.mock still intercepts
// dynamic imports since it hoists above them.
const { mermaidRenderMock } = vi.hoisted(() => ({
  mermaidRenderMock: vi.fn(),
}));

vi.mock('mermaid', () => ({
  default: { initialize: vi.fn(), render: mermaidRenderMock },
}));

// Imported after the mocks above (mirrors export.markdown.test.tsx's ordering).
import { renderMathImage, renderMermaidImage } from '../docxMedia';

describe('scaleToMaxWidth', () => {
  it('returns the original size unchanged when already within the max width', () => {
    expect(scaleToMaxWidth(200, 100)).toEqual({ width: 200, height: 100 });
  });

  it('scales width down to the max and height proportionally', () => {
    expect(scaleToMaxWidth(1100, 550, 550)).toEqual({ width: 550, height: 275 });
  });

  it('uses the default 550 max width when none is given', () => {
    expect(scaleToMaxWidth(1650, 300)).toEqual({ width: 550, height: 100 });
  });

  it('never rounds height down to 0 for an extreme aspect ratio', () => {
    const { height } = scaleToMaxWidth(100_000, 1, 550);
    expect(height).toBeGreaterThanOrEqual(1);
  });
});

describe('renderMathImage / renderMermaidImage', () => {
  afterEach(() => {
    vi.restoreAllMocks();
    html2canvasMock.mockReset();
    mermaidRenderMock.mockReset();
  });

  it('renderMathImage resolves a rasterized image for valid LaTeX', async () => {
    html2canvasMock.mockResolvedValue({ width: 100, height: 40, toDataURL: () => 'data:image/png;base64,AAAA' });

    const image = await renderMathImage('E = mc^2', false);

    expect(image).not.toBeNull();
    expect(image!.width).toBe(50);
    expect(image!.height).toBe(20);
    expect(image!.bytes).toBeInstanceOf(Uint8Array);
  });

  it('renderMathImage resolves null when rasterization produces an empty canvas', async () => {
    html2canvasMock.mockResolvedValue({ width: 0, height: 0, toDataURL: () => '' });

    await expect(renderMathImage('E = mc^2', false)).resolves.toBeNull();
  });

  it('renderMathImage resolves null (never throws) when html2canvas itself rejects', async () => {
    html2canvasMock.mockRejectedValue(new Error('canvas blew up'));

    await expect(renderMathImage('E = mc^2', false)).resolves.toBeNull();
  });

  it('renderMermaidImage passes the exact diagram source to mermaid.render and rasterizes the result', async () => {
    mermaidRenderMock.mockResolvedValue({ svg: '<svg><text>diagram</text></svg>' });
    html2canvasMock.mockResolvedValue({ width: 200, height: 100, toDataURL: () => 'data:image/png;base64,AAAA' });

    const image = await renderMermaidImage('graph TD; A-->B;');

    expect(mermaidRenderMock).toHaveBeenCalledWith(expect.stringContaining('docx-export-mermaid-'), 'graph TD; A-->B;');
    expect(image).not.toBeNull();
  });

  it('renderMermaidImage resolves null (never throws) when mermaid.render itself rejects', async () => {
    mermaidRenderMock.mockRejectedValue(new Error('invalid diagram syntax'));

    await expect(renderMermaidImage('not a real diagram')).resolves.toBeNull();
  });
});

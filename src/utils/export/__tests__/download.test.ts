import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { sanitizeFileName, saveBinaryOutput, saveTextOutput } from '../download';

describe('sanitizeFileName', () => {
  it('strips an existing extension and appends the requested one', () => {
    expect(sanitizeFileName('notes.md', 'html')).toBe('notes.html');
  });

  it('appends the extension when the name has none', () => {
    expect(sanitizeFileName('notes', 'docx')).toBe('notes.docx');
  });

  it('only strips the final extension, keeping dots elsewhere in the name', () => {
    expect(sanitizeFileName('My v1.2 Notes.md', 'pdf')).toBe('My v1.2 Notes.pdf');
  });
});

describe('saveTextOutput', () => {
  let createObjectURLMock: ReturnType<typeof vi.spyOn>;
  let anchorClicks: Array<{ download: string; href: string }>;

  beforeEach(() => {
    anchorClicks = [];
    createObjectURLMock = vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:mock-url');
    vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
      anchorClicks.push({ download: this.download, href: this.href });
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
    delete (window as { electronAPI?: unknown }).electronAPI;
  });

  it('falls back to a browser download outside Electron', async () => {
    await saveTextOutput('hello', 'notes.md', 'text/markdown');

    expect(anchorClicks).toEqual([{ download: 'notes.md', href: 'blob:mock-url' }]);
    const blob = createObjectURLMock.mock.calls[0]![0] as Blob;
    expect(blob.type).toBe('text/markdown');
    await expect(blob.text()).resolves.toBe('hello');
  });

  it('routes through window.electronAPI.saveFile with the given filters, skipping the browser download', async () => {
    const saveFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveFile: saveFileMock } as unknown as typeof window.electronAPI;

    await saveTextOutput('hello', 'notes.md', 'text/markdown', [{ name: 'Markdown', extensions: ['md'] }]);

    expect(saveFileMock).toHaveBeenCalledWith({
      content: 'hello',
      suggestedName: 'notes.md',
      filters: [{ name: 'Markdown', extensions: ['md'] }],
    });
    expect(anchorClicks).toHaveLength(0);
    expect(createObjectURLMock).not.toHaveBeenCalled();
  });

  it('resolves false when the user cancels the Electron dialog (no error, no fallback download)', async () => {
    window.electronAPI = {
      saveFile: vi.fn().mockResolvedValue({ saved: false }),
    } as unknown as typeof window.electronAPI;

    await expect(saveTextOutput('hello', 'notes.md', 'text/markdown')).resolves.toBe(false);
    expect(anchorClicks).toHaveLength(0);
  });

  it('throws when the Electron dialog reports a real failure', async () => {
    window.electronAPI = {
      saveFile: vi.fn().mockResolvedValue({ saved: false, error: 'Permission denied' }),
    } as unknown as typeof window.electronAPI;

    await expect(saveTextOutput('hello', 'notes.md', 'text/markdown')).rejects.toThrow('Permission denied');
  });
});

describe('saveBinaryOutput', () => {
  let createObjectURLMock: ReturnType<typeof vi.spyOn>;
  let anchorClicks: Array<{ download: string; href: string }>;

  beforeEach(() => {
    anchorClicks = [];
    createObjectURLMock = vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:mock-url');
    vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
      anchorClicks.push({ download: this.download, href: this.href });
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
    delete (window as { electronAPI?: unknown }).electronAPI;
  });

  it('falls back to a browser download outside Electron for a Uint8Array payload', async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    await saveBinaryOutput(bytes, 'doc.docx', 'application/octet-stream');

    expect(anchorClicks).toEqual([{ download: 'doc.docx', href: 'blob:mock-url' }]);
    const blob = createObjectURLMock.mock.calls[0]![0] as Blob;
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(bytes);
  });

  it('accepts a plain ArrayBuffer payload too', async () => {
    const buffer = new Uint8Array([9, 9, 9]).buffer;
    await saveBinaryOutput(buffer, 'doc.pdf', 'application/pdf');

    const blob = createObjectURLMock.mock.calls[0]![0] as Blob;
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(new Uint8Array([9, 9, 9]));
  });

  it('routes through window.electronAPI.saveBinaryFile, skipping the browser download', async () => {
    const saveBinaryFileMock = vi.fn().mockResolvedValue({ saved: true });
    window.electronAPI = { saveBinaryFile: saveBinaryFileMock } as unknown as typeof window.electronAPI;

    await saveBinaryOutput(new Uint8Array([1]), 'doc.docx', 'application/octet-stream', [
      { name: 'Word Document', extensions: ['docx'] },
    ]);

    expect(saveBinaryFileMock).toHaveBeenCalledWith(
      expect.objectContaining({ suggestedName: 'doc.docx', filters: [{ name: 'Word Document', extensions: ['docx'] }] }),
    );
    expect(anchorClicks).toHaveLength(0);
  });

  it('throws when the Electron dialog reports a real failure', async () => {
    window.electronAPI = {
      saveBinaryFile: vi.fn().mockResolvedValue({ saved: false, error: 'Disk is full' }),
    } as unknown as typeof window.electronAPI;

    await expect(saveBinaryOutput(new Uint8Array([1]), 'doc.docx', 'application/octet-stream')).rejects.toThrow(
      'Disk is full',
    );
  });
});

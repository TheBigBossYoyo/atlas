/**
 * FIELD-02 — CodeMirror keymap precedence regression coverage: `basicSetup`
 * registers its own default `searchKeymap` (Mod-g -> findNext) before this
 * editor's own `editorKeys` (Mod-g -> gotoLine) in the extensions array, so
 * without an explicit `Prec.high` on `editorKeys`, Ctrl+G silently did
 * whatever `findNext` does with no active query (opens the *search* panel,
 * never the goto-line one) instead of opening "Go to line". Ctrl+F must keep
 * working exactly as before — it isn't touched by `editorKeys` at all.
 */
import { fireEvent, render, waitFor } from '@testing-library/react';
import { describe, expect, it } from 'vitest';

import { CodeEditor } from '../CodeEditor';

function renderEditor(initialText: string) {
  return render(
    <CodeEditor
      initialText={initialText}
      fileName="sample.ts"
      dark={false}
      wrap={false}
      onDocChange={() => {}}
      onReady={() => {}}
    />,
  );
}

describe('CodeEditor — keymap precedence (FIELD-02)', () => {
  it('Ctrl+G opens "Go to line", not basicSetup\'s default find-next/search panel', async () => {
    const { container } = renderEditor('line one\nline two\nline three\n');
    await waitFor(() => expect(container.querySelector('.cm-content')).not.toBeNull());
    const content = container.querySelector('.cm-content') as HTMLElement;

    fireEvent.keyDown(content, { key: 'g', ctrlKey: true });

    await waitFor(() => expect(container.querySelector('.cm-panel.cm-goto-line')).not.toBeNull());
    // The pre-fix behavior: basicSetup's Mod-g -> findNext with no active
    // query falls back to opening the *search* panel instead.
    expect(container.querySelector('.cm-panel.cm-search')).toBeNull();
  });

  it('Ctrl+F still opens the combined find/replace panel (unaffected by the Mod-g fix)', async () => {
    const { container } = renderEditor('const a = 1\n');
    await waitFor(() => expect(container.querySelector('.cm-content')).not.toBeNull());
    const content = container.querySelector('.cm-content') as HTMLElement;

    fireEvent.keyDown(content, { key: 'f', ctrlKey: true });

    await waitFor(() => expect(container.querySelector('.cm-panel.cm-search')).not.toBeNull());
    expect(container.querySelector('input[name="replace"]')).not.toBeNull();
  });
});

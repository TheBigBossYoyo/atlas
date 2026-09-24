import { useEffect } from 'react';

const KEY = 'atlas-draft';

export interface Draft {
  markdown: string;
  fileName: string | null;
  /** SHELL-6 — which file the draft belongs to, so reopening that file can offer it. Absent in drafts written before this field existed. */
  filePath?: string | null;
  savedAt: number;
}

function isDraft(value: unknown): value is Draft {
  if (typeof value !== 'object' || value === null) return false;
  const candidate = value as Partial<Draft>;
  return (
    typeof candidate.markdown === 'string' &&
    (typeof candidate.fileName === 'string' || candidate.fileName === null) &&
    (candidate.filePath === undefined || candidate.filePath === null || typeof candidate.filePath === 'string') &&
    typeof candidate.savedAt === 'number'
  );
}

/**
 * @param enabled P2.6/SHELL-12 — gate autosave to markdown (pass
 * `isMarkdownDocument`) so a stale draft is never written/paired with an
 * unrelated binary file's name. When `false`, this hook does nothing at all
 * (it neither writes nor clears any existing draft — restoring a draft is
 * `App.tsx`'s job, via `loadDraft`/`clearDraft`). SHELL-6 — App also keeps it
 * `false` while there are no unsaved changes (loading a file is not an edit,
 * and used to overwrite a crashed session's draft ~800ms after relaunch) and
 * while a recovered draft is still being offered.
 * @param filePath SHELL-6 — recorded in the draft so reopening that file offers it.
 */
export function useAutosave(
  markdown: string,
  fileName: string | null,
  enabled = true,
  debounceMs = 800,
  filePath: string | null = null,
): void {
  useEffect(() => {
    if (!enabled) return;

    const trimmed = markdown.trim();
    if (trimmed.length === 0) {
      localStorage.removeItem(KEY);
      return;
    }

    const timer = window.setTimeout(() => {
      const draft: Draft = { markdown, fileName, filePath, savedAt: Date.now() };
      localStorage.setItem(KEY, JSON.stringify(draft));
    }, debounceMs);

    return () => {
      window.clearTimeout(timer);
    };
  }, [debounceMs, enabled, fileName, filePath, markdown]);
}

export function loadDraft(): Draft | null {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw == null) return null;
    const parsed: unknown = JSON.parse(raw);
    return isDraft(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

export function clearDraft(): void {
  localStorage.removeItem(KEY);
}

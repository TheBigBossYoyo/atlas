import { History } from 'lucide-react';
import type { Draft } from '../hooks/useAutosave';

interface DraftRecoveryBannerProps {
  draft: Draft | null;
  onRestore: () => void;
  onDiscard: () => void;
}

function formatSavedAt(savedAt: number): string {
  try {
    return new Date(savedAt).toLocaleString(undefined, {
      dateStyle: 'medium',
      timeStyle: 'short',
    });
  } catch {
    return new Date(savedAt).toLocaleString();
  }
}

/**
 * P2.6/SHELL-11/LOAD-20 — `useAutosave` writes a markdown draft every 800ms
 * but nothing ever read it back, so it protected nothing despite being
 * advertised. Shown on launch when a leftover draft is found (e.g. the app
 * was force-quit or crashed with unsaved markdown edits still pending).
 */
export function DraftRecoveryBanner({ draft, onRestore, onDiscard }: DraftRecoveryBannerProps) {
  if (!draft) return null;

  return (
    <div className="draft-recovery" role="alert">
      <History size={16} aria-hidden="true" />
      <span className="draft-recovery__message">
        {draft.fileName
          ? `Unsaved changes to "${draft.fileName}" from ${formatSavedAt(draft.savedAt)} were found.`
          : `An unsaved draft from ${formatSavedAt(draft.savedAt)} was found.`}
      </span>
      <div className="draft-recovery__actions">
        <button className="draft-recovery__btn draft-recovery__btn--discard" onClick={onDiscard} type="button">
          Discard
        </button>
        <button className="draft-recovery__btn draft-recovery__btn--restore" onClick={onRestore} type="button">
          Restore
        </button>
      </div>
    </div>
  );
}

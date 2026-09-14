import { FileText, Upload, BookOpen, Clock, X } from 'lucide-react';
import type { RecentFile } from '../types';

interface WelcomeScreenProps {
  onOpenFile: () => void;
  onLoadSample: () => void;
  recent: readonly RecentFile[];
  onOpenRecent?: (file: RecentFile) => void;
  onRemoveRecent?: (key: string) => void;
}

function formatRelative(ts: number): string {
  const diff = Date.now() - ts;
  const min = Math.floor(diff / 60000);
  if (min < 1) return 'just now';
  if (min < 60) return `${min}m ago`;
  const hr = Math.floor(min / 60);
  if (hr < 24) return `${hr}h ago`;
  const d = Math.floor(hr / 24);
  if (d < 7) return `${d}d ago`;
  return new Date(ts).toLocaleDateString();
}

export function WelcomeScreen({
  onOpenFile,
  onLoadSample,
  recent,
  onOpenRecent,
  onRemoveRecent,
}: WelcomeScreenProps) {
  return (
    <div className="welcome">
      <div className="welcome__content">
        <div className="welcome__icon">
          <BookOpen size={48} strokeWidth={1.5} />
        </div>
        <h1 className="welcome__title">Atlas</h1>
        <p className="welcome__subtitle">
          A universal document viewer for Word, Excel, PowerPoint, PDF, code, markdown, and more.
        </p>

        <div className="welcome__actions">
          <button className="welcome__btn welcome__btn--primary" onClick={onOpenFile}>
            <FileText size={20} />
            <span>Open File</span>
          </button>
          <button className="welcome__btn welcome__btn--secondary" onClick={onLoadSample}>
            <Upload size={20} />
            <span>Load Sample Document</span>
          </button>
        </div>

        <div className="welcome__hint">
          <p>or drag & drop any document anywhere</p>
        </div>

        {recent.length > 0 && (
          <div className="recent">
            <div className="recent__header">
              <Clock size={14} />
              <span>Recent files</span>
            </div>
            <ul className="recent__list">
              {recent.map(file => {
                const key = file.path || file.name;
                return (
                  <li key={key} className="recent__item">
                    <button
                      className="recent__open"
                      onClick={() => onOpenRecent?.(file)}
                      title={file.path || file.name}
                    >
                      <span className="recent__name">{file.name}</span>
                      <span className="recent__time">{formatRelative(file.openedAt)}</span>
                    </button>
                    {onRemoveRecent && (
                      <button
                        className="recent__remove"
                        onClick={() => onRemoveRecent(key)}
                        aria-label={`Remove ${file.name} from recent`}
                        title="Remove from recent"
                      >
                        <X size={13} />
                      </button>
                    )}
                  </li>
                );
              })}
            </ul>
          </div>
        )}

        <div className="welcome__features">
          {[
            { emoji: '📄', text: 'Word, Excel, PowerPoint, PDF' },
            { emoji: '📝', text: 'Markdown + code highlighting' },
            { emoji: '📐', text: 'Math + Mermaid diagrams' },
            { emoji: '🎨', text: '5 beautiful themes' },
            { emoji: '🔍', text: 'In-document search' },
            // RUN-13 — scoped to what's actually true today: every format
            // exports to PDF, but DOCX/HTML export is Markdown-only until
            // per-format export (Task X1) ships.
            { emoji: '💾', text: 'PDF export, plus DOCX/HTML for Markdown' },
          ].map((f) => (
            <div key={f.text} className="welcome__feature">
              <span>{f.emoji}</span>
              <span>{f.text}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

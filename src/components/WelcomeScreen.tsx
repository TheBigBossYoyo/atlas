import { FileText, Upload, BookOpen, Clock, X } from 'lucide-react';
import type { RecentFile } from '../types';
import { useTranslate, type TranslateFn } from '../i18n';

interface WelcomeScreenProps {
  onOpenFile: () => void;
  onLoadSample: () => void;
  recent: readonly RecentFile[];
  onOpenRecent?: (file: RecentFile) => void;
  onRemoveRecent?: (key: string) => void;
}

function formatRelative(ts: number, t: TranslateFn): string {
  const diff = Date.now() - ts;
  const min = Math.floor(diff / 60000);
  if (min < 1) return t('welcome.relative.justNow');
  if (min < 60) return t('welcome.relative.minutesAgo', { count: min });
  const hr = Math.floor(min / 60);
  if (hr < 24) return t('welcome.relative.hoursAgo', { count: hr });
  const d = Math.floor(hr / 24);
  if (d < 7) return t('welcome.relative.daysAgo', { count: d });
  return new Date(ts).toLocaleDateString();
}

const FEATURE_KEYS: readonly { emoji: string; key: string }[] = [
  { emoji: '📄', key: 'welcome.feature.office' },
  { emoji: '📝', key: 'welcome.feature.markdown' },
  { emoji: '📐', key: 'welcome.feature.math' },
  { emoji: '🎨', key: 'welcome.feature.themes' },
  { emoji: '🔍', key: 'welcome.feature.search' },
  // RUN-13 — scoped to what's actually true today: every format exports to
  // PDF, but DOCX/HTML export is Markdown-only until per-format export
  // (Task X1) ships.
  { emoji: '💾', key: 'welcome.feature.export' },
];

export function WelcomeScreen({
  onOpenFile,
  onLoadSample,
  recent,
  onOpenRecent,
  onRemoveRecent,
}: WelcomeScreenProps) {
  const t = useTranslate();
  return (
    <div className="welcome">
      <div className="welcome__content">
        <div className="welcome__icon">
          <BookOpen size={48} strokeWidth={1.5} />
        </div>
        <h1 className="welcome__title">Atlas</h1>
        <p className="welcome__subtitle">
          {t('welcome.subtitle')}
        </p>

        <div className="welcome__actions">
          <button className="welcome__btn welcome__btn--primary" onClick={onOpenFile}>
            <FileText size={20} />
            <span>{t('welcome.openFile')}</span>
          </button>
          <button className="welcome__btn welcome__btn--secondary" onClick={onLoadSample}>
            <Upload size={20} />
            <span>{t('welcome.loadSample')}</span>
          </button>
        </div>

        <div className="welcome__hint">
          <p>{t('welcome.dragDropHint')}</p>
        </div>

        {recent.length > 0 && (
          <div className="recent">
            <div className="recent__header">
              <Clock size={14} />
              <span>{t('welcome.recentFiles')}</span>
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
                      <span className="recent__time">{formatRelative(file.openedAt, t)}</span>
                    </button>
                    {onRemoveRecent && (
                      <button
                        className="recent__remove"
                        onClick={() => onRemoveRecent(key)}
                        aria-label={t('welcome.removeRecentAria', { name: file.name })}
                        title={t('welcome.removeRecentTitle')}
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
          {FEATURE_KEYS.map((f) => (
            <div key={f.key} className="welcome__feature">
              <span>{f.emoji}</span>
              <span>{t(f.key)}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

import { useTranslate } from '../i18n';

interface RawEditorProps {
  markdown: string;
  onChange: (value: string) => void;
}

export function RawEditor({ markdown, onChange }: RawEditorProps) {
  const t = useTranslate();

  return (
    <div className="editor-panel">
      <div className="editor-panel__header">
        <span>{t('rawEditor.title')}</span>
      </div>
      <textarea
        className="editor-panel__textarea"
        value={markdown}
        onChange={(e) => onChange(e.target.value)}
        spellCheck={false}
        placeholder={t('rawEditor.placeholder')}
        // I18N-1 — the placeholder disappears the instant there is text, so
        // it can't be this textarea's only name (WCAG 3.3.2). aria-label
        // gives it a real accessible name that survives non-empty content.
        aria-label={t('rawEditor.title')}
      />
    </div>
  );
}

interface RawEditorProps {
  markdown: string;
  onChange: (value: string) => void;
}

export function RawEditor({ markdown, onChange }: RawEditorProps) {
  return (
    <div className="editor-panel">
      <div className="editor-panel__header">
        <span>Markdown Source</span>
      </div>
      <textarea
        className="editor-panel__textarea"
        value={markdown}
        onChange={(e) => onChange(e.target.value)}
        spellCheck={false}
        placeholder="Type or paste markdown here..."
      />
    </div>
  );
}

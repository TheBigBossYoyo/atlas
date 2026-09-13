import { FileDown } from 'lucide-react';

interface DropZoneProps {
  isVisible: boolean;
}

export function DropZone({ isVisible }: DropZoneProps) {
  if (!isVisible) return null;

  return (
    <div className="dropzone">
      <div className="dropzone__content">
        <FileDown size={48} strokeWidth={1.5} />
        <h3>Drop your Markdown file</h3>
        <p>.md, .markdown, or .txt</p>
      </div>
    </div>
  );
}

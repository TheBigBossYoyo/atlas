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
        <h3>Drop a file to open it</h3>
        <p>Markdown, DOCX, PDF, spreadsheets, slides, and more</p>
      </div>
    </div>
  );
}

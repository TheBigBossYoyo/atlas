import { FileDown } from 'lucide-react';

interface DropZoneProps {
  isVisible: boolean;
}

// UX-16 / LOAD-22 — this used to claim "Markdown only" while the app has
// detected and viewed 13 real formats since format detection landed
// (src/formats/detect.ts); kept short/representative rather than listing
// all 13 extensions.
export function DropZone({ isVisible }: DropZoneProps) {
  if (!isVisible) return null;

  return (
    <div className="dropzone">
      <div className="dropzone__content">
        <FileDown size={48} strokeWidth={1.5} />
        <h3>Drop your document</h3>
        <p>Word, Excel, PowerPoint, PDF, OpenDocument, Markdown, CSV, RTF, and code files</p>
      </div>
    </div>
  );
}

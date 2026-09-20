import { useEffect, useState, useCallback } from 'react';

interface MermaidProps {
  code: string;
  id: string;
}

const detectTheme = () => {
  const theme = document.documentElement.getAttribute('data-theme');
  if (theme === 'dark' || theme === 'dracula' || theme === 'nord') {
    return 'dark';
  }
  return 'default';
};

export function Mermaid({ code, id }: MermaidProps) {
  const [svg, setSvg] = useState<string>('');
  const [error, setError] = useState<string | null>(null);

  const render = useCallback(async () => {
    try {
      const m = (await import('mermaid')).default;
      m.initialize({ startOnLoad: false, theme: detectTheme(), securityLevel: 'loose' });
      const { svg: renderedSvg } = await m.render(id, code);
      setSvg(renderedSvg);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      // Found by driving the real app on a broken diagram: mermaid's own
      // `render()` builds its error SVG inside a temporary `div#d<id>` it
      // appends directly to `document.body` (never inside this component's
      // own tree — see node_modules/mermaid's `render()`: it serializes that
      // div's contents into the string it resolves/rejects with, but only
      // calls its own `removeTempElements()` cleanup on the SUCCESS path;
      // the reject-with-parse-error path skips straight past it). Left
      // alone, that visible, full-sized error graphic (mermaid's bomb icon +
      // "Syntax error in text") stays attached below the entire app for the
      // rest of the window's lifetime — outside React's tree, so no
      // remount/unmount ever clears it — duplicating the inline
      // `.mermaid--error` message this component already renders in place.
      document.getElementById(`d${id}`)?.remove();
    }
  }, [code, id]);

  useEffect(() => {
    render();
  }, [render]);

  useEffect(() => {
    const observer = new MutationObserver(() => {
      render();
    });
    observer.observe(document.documentElement, {
      attributes: true,
      attributeFilter: ['data-theme']
    });
    return () => {
      observer.disconnect();
    };
  }, [render]);

  if (error) {
    return <pre className="mermaid mermaid--error">{`${error}\n\n${code}`}</pre>;
  }

  return <div className="mermaid" dangerouslySetInnerHTML={{ __html: svg }} />;
}

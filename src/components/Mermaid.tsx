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

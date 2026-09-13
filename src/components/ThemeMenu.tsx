import { useState, useRef, useEffect } from 'react';
import { Palette, Check } from 'lucide-react';
import type { Theme, ThemeMeta } from '../types';

interface ThemeMenuProps {
  current: Theme;
  themes: readonly ThemeMeta[];
  onSelect: (t: Theme) => void;
}

export function ThemeMenu({ current, themes, onSelect }: ThemeMenuProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    const handleEsc = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setOpen(false);
      }
    };
    if (open) {
      window.addEventListener('mousedown', handleOutsideClick);
      window.addEventListener('keydown', handleEsc);
    }
    return () => {
      window.removeEventListener('mousedown', handleOutsideClick);
      window.removeEventListener('keydown', handleEsc);
    };
  }, [open]);

  return (
    <div className="dropdown" ref={ref}>
      <button
        className="toolbar__btn toolbar__btn--icon toolbar__nodrag"
        title="Theme"
        aria-haspopup="menu"
        aria-expanded={open}
        onClick={() => setOpen(!open)}
      >
        <Palette size={18} />
      </button>
      {open && (
        <ul className="dropdown__menu" role="menu">
          {themes.map(t => (
            <li key={t.id}>
              <button
                className={`dropdown__item ${current === t.id ? 'dropdown__item--active' : ''}`}
                role="menuitemradio"
                aria-checked={current === t.id}
                onClick={() => {
                  onSelect(t.id);
                  setOpen(false);
                }}
              >
                <span className="dropdown__item-swatch" style={{ backgroundColor: t.overlayBg }} />
                <span className="dropdown__item-label">{t.label}</span>
                {current === t.id && <Check size={14} />}
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

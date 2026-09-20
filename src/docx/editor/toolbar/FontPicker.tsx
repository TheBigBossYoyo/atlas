import React, { useCallback, useEffect, useId, useMemo, useRef, useState } from 'react';

import { useTranslate } from '../../../i18n';

/**
 * USR-10/USR-11 — searchable font combobox. The previous native <select>
 * offered 8 hardcoded fonts and rendered blank whenever the document's font
 * wasn't one of them. This shows the effective font as-is, filters as you
 * type, lists document fonts first and then every installed font (each
 * rendered in its own face), and applies a typed name on Enter.
 */

interface FontPickerProps {
  value: string | null;
  fonts: ReadonlyArray<string>;
  documentFonts?: ReadonlyArray<string>;
  onSelect: (family: string) => void;
}

const MAX_VISIBLE = 200;

export const FontPicker: React.FC<FontPickerProps> = ({ value, fonts, documentFonts = [], onSelect }) => {
  const t = useTranslate();
  const listId = useId();
  const containerRef = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState<string | null>(null);
  const [activeIndex, setActiveIndex] = useState(0);

  const options = useMemo(() => {
    const documentSet = new Set(documentFonts);
    const rest = fonts.filter((font) => !documentSet.has(font));
    const all = [...documentFonts.map((font) => ({ font, group: 'document' as const })), ...rest.map((font) => ({ font, group: 'all' as const }))];
    const needle = (query ?? '').trim().toLowerCase();
    const filtered = needle.length === 0 ? all : all.filter((option) => option.font.toLowerCase().includes(needle));
    return filtered.slice(0, MAX_VISIBLE);
  }, [documentFonts, fonts, query]);

  const close = useCallback(() => {
    setOpen(false);
    setQuery(null);
  }, []);

  useEffect(() => {
    if (!open) return undefined;
    const handleMouseDown = (event: MouseEvent): void => {
      if (containerRef.current !== null && !containerRef.current.contains(event.target as Node)) {
        close();
      }
    };
    document.addEventListener('mousedown', handleMouseDown);
    return () => document.removeEventListener('mousedown', handleMouseDown);
  }, [close, open]);

  const apply = useCallback(
    (family: string) => {
      const trimmed = family.trim();
      if (trimmed.length > 0) {
        onSelect(trimmed);
      }
      close();
    },
    [close, onSelect],
  );

  const handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>): void => {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      setOpen(true);
      setActiveIndex((index) => Math.min(index + 1, Math.max(options.length - 1, 0)));
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      setActiveIndex((index) => Math.max(index - 1, 0));
    } else if (event.key === 'Enter') {
      event.preventDefault();
      const highlighted = open ? options[activeIndex] : undefined;
      apply(highlighted !== undefined && query !== null ? highlighted.font : query ?? value ?? '');
    } else if (event.key === 'Escape') {
      event.preventDefault();
      close();
    }
  };

  return (
    <div className="docx-font-picker" ref={containerRef}>
      <input
        className="docx-toolbar__font-select docx-font-picker__input"
        aria-label={t('docx.fontPicker.inputAria')}
        role="combobox"
        aria-expanded={open}
        aria-controls={listId}
        aria-autocomplete="list"
        placeholder={t('docx.fontPicker.placeholder')}
        spellCheck={false}
        value={query ?? value ?? ''}
        style={value !== null && query === null ? { fontFamily: `"${value}", system-ui, sans-serif` } : undefined}
        onFocus={(event) => {
          setOpen(true);
          event.currentTarget.select();
        }}
        onChange={(event) => {
          setQuery(event.target.value);
          setActiveIndex(0);
          setOpen(true);
        }}
        onKeyDown={handleKeyDown}
      />
      {open && options.length > 0 && (
        <ul className="docx-font-picker__list" role="listbox" id={listId} aria-label={t('docx.fontPicker.listAria')}>
          {options.map((option, index) => {
            const firstOfAll = option.group === 'all' && index > 0 && options[index - 1].group === 'document';
            return (
              <li
                key={`${option.group}-${option.font}`}
                role="option"
                aria-selected={index === activeIndex}
                className={`docx-font-picker__option${index === activeIndex ? ' docx-font-picker__option--active' : ''}${firstOfAll ? ' docx-font-picker__option--group-start' : ''}`}
                style={{ fontFamily: `"${option.font}", system-ui, sans-serif` }}
                onMouseDown={(event) => {
                  event.preventDefault();
                  apply(option.font);
                }}
                onMouseEnter={() => setActiveIndex(index)}
              >
                {option.font}
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
};

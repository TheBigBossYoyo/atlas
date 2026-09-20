import { ChevronRight, Hash, List } from 'lucide-react';
import { useNavItems } from '../viewers/shared/useViewerContext';
import { useTranslate } from '../i18n';

interface SidebarProps {
  isOpen: boolean;
}

export function Sidebar({ isOpen }: SidebarProps) {
  const items = useNavItems();
  const t = useTranslate();

  if (!isOpen) return null;

  return (
    <aside className="sidebar">
      <div className="sidebar__header">
        <List size={16} />
        <span>{t('sidebar.title')}</span>
      </div>
      <nav className="sidebar__nav">
        {items.length === 0 ? (
          <p className="sidebar__empty">{t('sidebar.empty')}</p>
        ) : (
          <ul className="sidebar__list">
            {items.map((item, idx) => (
              <li
                key={`${item.id}-${idx}`}
                className="sidebar__item"
                style={{ paddingLeft: `calc(var(--nav-indent, 12px) * ${item.level ?? 0})` }}
              >
                <button
                  className={`sidebar__link sidebar__link--h${item.level ?? 1}`}
                  onClick={item.onSelect}
                  title={item.label}
                >
                  {item.icon ? <item.icon size={14} /> : (item.level ?? 1) <= 1 ? <Hash size={14} /> : <ChevronRight size={14} />}
                  <span>{item.label}</span>
                </button>
              </li>
            ))}
          </ul>
        )}
      </nav>
    </aside>
  );
}

import { Hash } from 'lucide-react';
import { useEffect } from 'react';
import { fireEvent, render, screen } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import type { NavItem } from '../../formats/types';
import { ViewerProvider } from '../../viewers/shared/ViewerContext';
import { useSetNavItems } from '../../viewers/shared/useViewerContext';
import { Sidebar } from '../Sidebar';

function SidebarHarness({ items }: { items: readonly NavItem[] }) {
  const setNavItems = useSetNavItems();

  useEffect(() => {
    setNavItems(items);
  }, [items, setNavItems]);

  return <Sidebar isOpen />;
}

describe('Sidebar', () => {
  it('renders empty state when no nav items exist', () => {
    render(
      <ViewerProvider filePath="/tmp/readme.md">
        <SidebarHarness items={[]} />
      </ViewerProvider>
    );

    expect(screen.getByText('No outline available')).toBeInTheDocument();
  });

  it('renders contextual nav items with icons, labels, indentation, and click handlers', async () => {
    const onSelect = vi.fn();

    render(
      <ViewerProvider filePath="/tmp/readme.md">
        <SidebarHarness
          items={[
            { id: 'intro', label: 'Introduction', level: 1, icon: Hash, onSelect },
            { id: 'details', label: 'Details', level: 2, onSelect: vi.fn() },
          ]}
        />
      </ViewerProvider>
    );

    const introButton = screen.getByRole('button', { name: 'Introduction' });
    const detailsButton = screen.getByRole('button', { name: 'Details' });

    expect(introButton).toBeInTheDocument();
    expect(detailsButton).toBeInTheDocument();
    expect(detailsButton.closest('.sidebar__item')).toHaveStyle({
      paddingLeft: 'calc(var(--nav-indent, 12px) * 2)',
    });

    fireEvent.click(introButton);
    expect(onSelect).toHaveBeenCalledTimes(1);
  });
});

import { useEffect, type RefObject } from 'react';

const FOCUSABLE_SELECTOR = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';

interface UseFocusTrapOptions {
  /**
   * Move focus to the first focusable element inside the container when it
   * opens. Default true. Pass false when the caller already manages its own
   * (e.g. delayed-for-animation) autofocus and only needs Tab containment +
   * focus restore from this hook.
   */
  focusOnOpen?: boolean;
}

/**
 * A11Y-2 — shared Tab-containment + focus-restore behavior, factored out of
 * the pattern ShortcutsModal/UnsavedChangesDialog each hand-rolled (UX-13):
 * on open, remember whatever was focused, optionally move focus into the
 * container, keep Tab/Shift+Tab cycling within it while open, and restore
 * focus to whatever triggered it once it closes.
 */
export function useFocusTrap(
  containerRef: RefObject<HTMLElement | null>,
  isOpen: boolean,
  options: UseFocusTrapOptions = {},
): void {
  const { focusOnOpen = true } = options;

  useEffect(() => {
    if (!isOpen) return undefined;

    const previouslyFocused = document.activeElement as HTMLElement | null;
    const container = containerRef.current;
    if (focusOnOpen) {
      container?.querySelector<HTMLElement>(FOCUSABLE_SELECTOR)?.focus();
    }

    const handleTab = (e: KeyboardEvent) => {
      if (e.key !== 'Tab' || !container) return;
      const focusables = Array.from(container.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR)).filter(
        el => !el.hasAttribute('disabled'),
      );
      if (focusables.length === 0) return;
      const first = focusables[0]!;
      const last = focusables[focusables.length - 1]!;
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    };

    window.addEventListener('keydown', handleTab);
    return () => {
      window.removeEventListener('keydown', handleTab);
      previouslyFocused?.focus();
    };
  }, [isOpen, containerRef, focusOnOpen]);
}

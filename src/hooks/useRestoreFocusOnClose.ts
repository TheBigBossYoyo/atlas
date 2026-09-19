import { useEffect, useRef, type RefObject } from 'react';

/**
 * UX — the toolbar's non-modal dropdown menus (ThemeMenu/ExportMenu/
 * NewDocumentMenu) intentionally skip a full focus trap (UX-14 — plain
 * labeled buttons, not the `role="menu"`/arrow-key-nav ARIA pattern), but
 * that also meant closing one via Escape or picking an item never returned
 * focus anywhere: the closing `<ul>` unmounts (removing whatever inside it
 * was focused), and the browser drops focus straight to `document.body`
 * with nothing to pick it back up.
 *
 * Restores focus to the trigger in exactly that case — `document.activeElement`
 * having fallen back to `<body>` is what distinguishes "the menu itself
 * closed out from under the user's focus" from a close caused by clicking a
 * DIFFERENT focusable element (the outside-click-to-close path), where that
 * element already claimed focus intentionally and should keep it.
 */
export function useRestoreFocusOnClose(open: boolean, triggerRef: RefObject<HTMLElement | null>): void {
  const wasOpenRef = useRef(open);

  useEffect(() => {
    if (wasOpenRef.current && !open && document.activeElement === document.body) {
      triggerRef.current?.focus();
    }
    wasOpenRef.current = open;
  }, [open, triggerRef]);
}

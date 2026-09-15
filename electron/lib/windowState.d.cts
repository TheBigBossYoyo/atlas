export interface WindowBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface WindowState extends WindowBounds {
  isMaximized: boolean;
}

export interface DisplayLike {
  workArea: WindowBounds;
}

export const STATE_FILE_NAME: string;
export const MIN_VISIBLE_PX: number;

export function isSufficientlyVisible(
  bounds: WindowBounds,
  displays: ReadonlyArray<DisplayLike>,
): boolean;

export function clampBoundsToDisplays(
  bounds: WindowBounds | null,
  displays: ReadonlyArray<DisplayLike>,
): WindowBounds | null;

export function loadWindowState(storeDir: string): WindowState | null;

export function saveWindowState(storeDir: string, state: WindowState): void;

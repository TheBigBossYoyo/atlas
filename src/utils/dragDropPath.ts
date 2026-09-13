// Resolves a dropped `File` to an absolute, main-process-vouched path
// (P1.2/P1.4 — ELEC-01/SHELL-01/LOAD-02/RUN-04/SHELL-24).
//
// Electron 32+ removed `File.path`, so the path must come from
// `webUtils.getPathForFile` (exposed via preload as `getPathForFile`); the
// resolved path must then be registered with the main process's read
// allowlist (`registerDroppedPath`) before any IPC read handler will accept
// it. Extracted as a pure async function (independent of React) so the
// resolution logic is unit-testable without mounting the whole app.

export interface DragDropPathApi {
  readonly getPathForFile?: (file: File) => string;
  readonly registerDroppedPath?: (path: string) => Promise<{ readonly ok: boolean }>;
}

export async function resolveDroppedFilePath(
  api: DragDropPathApi | undefined,
  file: File | undefined | null
): Promise<string | null> {
  if (!file || !api?.getPathForFile) return null;

  const path = api.getPathForFile(file);
  if (!path) return null;

  if (api.registerDroppedPath) {
    const registration = await api.registerDroppedPath(path);
    if (!registration.ok) return null;
  }

  return path;
}

import React, { Suspense, lazy, useCallback, useMemo, useState } from 'react'
import type { LoadedFile, ViewerComponent } from '../formats/types'
import { viewerRegistry } from '../formats/registry'
import { ViewerLoading } from './ViewerLoading'
import { ViewerErrorBoundary } from './ViewerErrorBoundary'

export type ViewerRouterProps = {
  readonly file: LoadedFile
}

export function ViewerRouter({ file }: ViewerRouterProps): React.ReactElement {
  // Bumped by ViewerErrorBoundary's onReset ("Try again"). React.lazy()
  // permanently caches a rejected import promise, so a transient chunk-load
  // failure could never be retried (LOAD-09/DAT-21) — including retryCount
  // in the memo key below forces a brand-new lazy() (and thus a fresh
  // dynamic import) on every retry instead of replaying the dead promise.
  const [retryCount, setRetryCount] = useState(0)

  // LOAD-21 evaluated (wave-3 shell-polish follow-up), not implemented:
  // caching this `lazy()` wrapper per `file.format` alone (dropping
  // `file.path` from the dep array) would NOT actually avoid the
  // unmount/remount cost on a same-format file switch, because
  // `ViewerErrorBoundary` below is keyed on `file.path` (LOAD-08/RUN-09) —
  // a React `key` change on an ancestor already forces a full remount of
  // everything under it, independent of whether `LazyComponent`'s own
  // identity also changes. Removing/relaxing that key to fix the remount
  // would reintroduce LOAD-08/RUN-09's bug (a crash on one file permanently
  // showing "Viewer crashed" for every later file) unless
  // `ViewerErrorBoundary` were restructured to clear its own error state
  // from a lifecycle hook instead of from being unmounted — out of this
  // file's scope. Separately, at least one viewer (`DocxViewer`) seeds
  // several of its own `useState` calls (`documentModel`, `savePath`,
  // `lastSavedDocument`, `commentsPaneOpen`, `trackChangesEnabled`) directly
  // from `bundle`/`file.path` with no companion effect to resync them on a
  // later prop change — safe only because it currently always remounts on
  // file change; kept mounted across a same-format switch, it would keep
  // rendering/editing the PREVIOUS file's content. Revisit only alongside
  // that ViewerErrorBoundary rework and a full per-viewer state-rederivation
  // audit — not a scope this polish wave took on.
  const LazyComponent = useMemo(
    () =>
      lazy(
        (): Promise<{ default: ViewerComponent }> =>
          viewerRegistry[file.format]().then((Component) => ({ default: Component }))
      ),
    // file.format is included (LOAD-19) so a format change without a path
    // change still picks a fresh viewer instead of a stale cached lazy
    // component, and retryCount (LOAD-09/DAT-21) busts the lazy() cache on
    // "Try again".
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [file.path, file.format, retryCount]
  )

  const handleReset = useCallback(() => {
    setRetryCount((count) => count + 1)
  }, [])

  return (
    // Keyed by file.path (LOAD-08/RUN-09) so a crash on one file fully
    // remounts the boundary — clearing its cached error — the moment a
    // different file is opened, instead of permanently showing "Viewer
    // crashed" for every subsequently opened valid file.
    <ViewerErrorBoundary key={file.path} onReset={handleReset}>
      <Suspense fallback={<ViewerLoading format={file.format} />}>
        <div data-viewer={file.format}>
          <LazyComponent file={file} />
        </div>
      </Suspense>
    </ViewerErrorBoundary>
  )
}

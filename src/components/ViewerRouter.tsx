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

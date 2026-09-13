import React, { Suspense, lazy, useMemo } from 'react'
import type { LoadedFile, ViewerComponent } from '../formats/types'
import { viewerRegistry } from '../formats/registry'
import { ViewerLoading } from './ViewerLoading'
import { ViewerErrorBoundary } from './ViewerErrorBoundary'

export type ViewerRouterProps = {
  readonly file: LoadedFile
}

export function ViewerRouter({ file }: ViewerRouterProps): React.ReactElement {
  const LazyComponent = useMemo(
    () =>
      lazy(
        (): Promise<{ default: ViewerComponent }> =>
          viewerRegistry[file.format]().then((Component) => ({ default: Component }))
      ),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [file.path]
  )

  return (
    <ViewerErrorBoundary>
      <Suspense fallback={<ViewerLoading format={file.format} />}>
        <div data-viewer={file.format}>
          <LazyComponent file={file} />
        </div>
      </Suspense>
    </ViewerErrorBoundary>
  )
}

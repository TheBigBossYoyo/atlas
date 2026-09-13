/**
 * Atlas — DOCX media context (Wave E.4)
 *
 * Provides the renderer with a resolver that maps an image relationshipId
 * to a usable URL (typically a blob: URL).  When no provider is mounted,
 * resolution returns null and the renderer falls back to a placeholder.
 */

import { createContext, useContext } from 'react'

export interface MediaResolver {
  readonly resolve: (relationshipId: string) => string | null
}

const NoopResolver: MediaResolver = {
  resolve: () => null,
}

export const MediaContext = createContext<MediaResolver>(NoopResolver)

export function useMediaResolver(): MediaResolver {
  return useContext(MediaContext)
}

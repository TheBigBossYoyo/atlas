import { memo, useEffect, useMemo } from 'react'

import { MarkdownRenderer } from '../components/MarkdownRenderer'
import type { ViewerProps } from '../formats/types'
import { useToc } from '../hooks/useToc'
import {
  useSetNavItems,
  useSetViewerStats,
} from './shared/useViewerContext'

function MarkdownViewerBase({ file }: ViewerProps) {
  const markdown =
    file.kind === 'text'
      ? file.content
      : new TextDecoder('utf-8').decode(file.content)

  const tocItems = useToc(markdown)

  const words = useMemo(
    () => markdown.trim().split(/\s+/).filter(Boolean).length,
    [markdown],
  )

  const navItems = useMemo(
    () =>
      tocItems.map(item => ({
        id: item.id,
        label: item.text,
        level: item.level,
        onSelect: () =>
          document
            .getElementById(item.id)
            ?.scrollIntoView({ behavior: 'smooth', block: 'start' }),
      })),
    [tocItems],
  )

  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()

  useEffect(() => {
    setNavItems(navItems)
  }, [setNavItems, navItems])

  useEffect(() => {
    setStats({ kind: 'markdown', words, headings: tocItems.length })
  }, [setStats, words, tocItems.length])

  return (
    <div className="markdown-viewer">
      <MarkdownRenderer markdown={markdown} />
    </div>
  )
}

export const MarkdownViewer = memo(MarkdownViewerBase)

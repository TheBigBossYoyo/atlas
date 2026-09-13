import { memo, useEffect, useState, useMemo } from 'react'

import type { ViewerProps, NavItem } from '../formats/types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'
import { getShikiThemeForAppTheme } from './shared/shikiTheme'
import { useTheme } from '../hooks/useTheme'
import { getLangForExt } from './extToLang'

function CodeViewerBase({ file }: ViewerProps) {
  const content = file.kind === 'text' ? file.content : ''
  const linesCount = useMemo(() => content.split('\n').length, [content])
  
  const { theme: appTheme } = useTheme()
  const shikiTheme = getShikiThemeForAppTheme(appTheme)
  
  const extMatch = file.path.match(/\.[^.]+$/)
  const ext = extMatch ? extMatch[0] : ''
  const lang = getLangForExt(ext) || 'text'
  
  const [html, setHtml] = useState<string>('')

  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()

  const navItems = useMemo(() => {
    const lines = content.split('\n')
    const items: NavItem[] = []
    
    let regexes: RegExp[] = []
    if (['typescript', 'tsx', 'javascript', 'jsx'].includes(lang)) {
      regexes = [
        /(?:export\s+)?(?:async\s+)?function\s+(\w+)/,
        /(?:export\s+)?(?:const|let)\s+(\w+)\s*=\s*(?:async\s+)?\(/,
        /class\s+(\w+)/
      ]
    } else if (lang === 'python') {
      regexes = [/^(?:def|class)\s+(\w+)/]
    } else if (lang === 'rust') {
      regexes = [/^(?:fn|struct|enum)\s+(\w+)/]
    }

    if (regexes.length > 0) {
      lines.forEach((line, idx) => {
        for (const regex of regexes) {
          const match = line.match(regex)
          if (match && match[1]) {
            items.push({
              id: `line-${idx}`,
              label: match[1],
              onSelect: () => {
                const els = document.querySelectorAll('.code-viewer__pre .line')
                if (els[idx]) {
                  els[idx].scrollIntoView({ behavior: 'smooth', block: 'start' })
                }
              }
            })
            break
          }
        }
      })
    }
    
    return items
  }, [content, lang])

  useEffect(() => {
    setNavItems(navItems)
  }, [setNavItems, navItems])

  useEffect(() => {
    setStats({ kind: 'code', language: lang, lines: linesCount })
  }, [setStats, lang, linesCount])

  useEffect(() => {
    let cancelled = false

    const load = async () => {
      try {
        const { getSingletonHighlighter } = await import('shiki')
        const langs = lang === 'text' ? [] : [lang]
        
        // We cast langs to any to bypass strict literal checks if needed, but Shiki handles strings well.
        const highlighter = await getSingletonHighlighter({
          themes: [shikiTheme],
          langs: langs
        })
        
        if (cancelled) return
        
        const loadedLangs = highlighter.getLoadedLanguages()
        const finalLang = loadedLangs.includes(lang) ? lang : 'text'

        const htmlStr = highlighter.codeToHtml(content, {
          lang: finalLang,
          theme: shikiTheme
        })
        
        if (!cancelled) {
          setHtml(htmlStr)
        }
      } catch (err) {
        console.error('Failed to highlight code:', err)
        if (!cancelled) {
          setHtml(`<pre><code>${content.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</code></pre>`)
        }
      }
    }
    
    load()

    return () => {
      cancelled = true
    }
  }, [content, lang, shikiTheme])

  return (
    <div className="code-viewer">
      {html ? (
        <div 
          className="code-viewer__pre"
          dangerouslySetInnerHTML={{ __html: html }}
        />
      ) : (
        <pre className="code-viewer__pre"><code>{content}</code></pre>
      )}
    </div>
  )
}

export const CodeViewer = memo(CodeViewerBase)

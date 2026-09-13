import { memo, useCallback, useEffect, useMemo, useState } from 'react'

import type { ViewerProps } from '../formats/types'
import { SlideDeck } from './shared/SlideDeck'
import type { SlideData } from './shared/SlideDeck.types'
import { useSetNavItems, useSetViewerStats } from './shared/useViewerContext'

type CancelSignal = {
  cancelled: boolean
}

type ZipEntry = {
  async(type: 'string'): Promise<string>
  async(type: 'base64'): Promise<string>
}

type ZipArchive = {
  file(path: string): ZipEntry | null
}

const DEFAULT_SLIDE_WIDTH = 1280
const DEFAULT_SLIDE_HEIGHT = 720
const CENTIMETERS_TO_PIXELS = 37.7952755906

function OdpViewerBase({ file }: ViewerProps) {
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()
  const [slides, setSlides] = useState<ReadonlyArray<SlideData>>([])
  const [activeIndex, setActiveIndex] = useState(0)
  const [error, setError] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)

  const handleSelectSlide = useCallback((index: number) => {
    setActiveIndex(currentIndex => (currentIndex === index ? currentIndex : index))
  }, [])

  const navItems = useMemo(
    () =>
      slides.map((slide, index) => ({
        id: slide.id,
        label: `Slide ${index + 1}`,
        onSelect: () => {
          handleSelectSlide(index)
        },
      })),
    [handleSelectSlide, slides],
  )

  useEffect(() => {
    setNavItems(navItems)
  }, [navItems, setNavItems])

  useEffect(() => {
    if (slides.length === 0) {
      setStats(null)
      return
    }

    setStats({ kind: 'slides', slide: activeIndex + 1, slideCount: slides.length })
  }, [activeIndex, setStats, slides.length])

  useEffect(() => {
    if (slides.length === 0) {
      return
    }

    setActiveIndex(currentIndex => Math.min(currentIndex, slides.length - 1))
  }, [slides.length])

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (slides.length === 0) {
        return
      }

      const target = event.target
      if (
        target instanceof HTMLElement
        && (target.isContentEditable
          || target.tagName === 'INPUT'
          || target.tagName === 'TEXTAREA'
          || target.tagName === 'SELECT')
      ) {
        return
      }

      if (event.key === 'PageDown') {
        event.preventDefault()
        setActiveIndex(currentIndex => Math.min(currentIndex + 1, slides.length - 1))
      }

      if (event.key === 'PageUp') {
        event.preventDefault()
        setActiveIndex(currentIndex => Math.max(currentIndex - 1, 0))
      }
    }

    window.addEventListener('keydown', handleKeyDown)

    return () => {
      window.removeEventListener('keydown', handleKeyDown)
    }
  }, [slides.length])

  useEffect(() => {
    setNavItems([])
    setStats(null)
    setSlides([])
    setActiveIndex(0)
    setError(null)

    if (file.kind !== 'binary') {
      setIsLoading(false)
      return
    }

    const signal: CancelSignal = { cancelled: false }

    setIsLoading(true)

    void (async () => {
      try {
        const { default: JSZip } = await import('jszip')

        if (signal.cancelled) {
          return
        }

        const zip = await JSZip.loadAsync(file.content)

        if (signal.cancelled) {
          return
        }

        const nextSlides = await parseOdpSlides(zip as ZipArchive, signal)

        if (signal.cancelled) {
          return
        }

        setSlides(nextSlides)
        setActiveIndex(0)
        setError(nextSlides.length > 0 ? null : 'No slides found in ODP.')
        setIsLoading(false)
      } catch (err: unknown) {
        if (signal.cancelled) {
          return
        }

        setError(err instanceof Error ? err.message : String(err))
        setIsLoading(false)
      }
    })()

    return () => {
      signal.cancelled = true
    }
  }, [file, setNavItems, setStats])

  if (file.kind !== 'binary') {
    return (
      <div className="odp-viewer odp-viewer--error">
        OdpViewer received a text file; expected binary.
      </div>
    )
  }

  if (error !== null) {
    return (
      <div className="odp-viewer odp-viewer--error">
        Failed to render ODP: {error}
      </div>
    )
  }

  if (isLoading) {
    return <div className="odp-viewer">Loading ODP slides…</div>
  }

  return (
    <div className="odp-viewer" style={{ width: '100%', height: '100%' }}>
      <SlideDeck
        slides={slides}
        activeIndex={activeIndex}
        onSelect={handleSelectSlide}
      />
    </div>
  )
}

function parseXml(xml: string): XMLDocument {
  const document = new DOMParser().parseFromString(xml, 'application/xml')
  const parserError = document.getElementsByTagName('parsererror')[0]

  if (parserError) {
    throw new Error(parserError.textContent?.trim() || 'Invalid XML document.')
  }

  return document
}

function getElementsByLocalName(root: XMLDocument | Element, localName: string): Element[] {
  return Array.from(root.getElementsByTagNameNS('*', localName))
}

function getFirstByLocalName(root: XMLDocument | Element, localName: string): Element | null {
  return getElementsByLocalName(root, localName)[0] ?? null
}

function getAttributeValue(element: Element, names: ReadonlyArray<string>): string | null {
  for (const name of names) {
    const value = element.getAttribute(name)
    if (value !== null) {
      return value
    }
  }

  return null
}

function normalizeZipPath(path: string): string {
  const parts: string[] = []

  for (const segment of path.replace(/\\/g, '/').split('/')) {
    if (!segment || segment === '.') {
      continue
    }

    if (segment === '..') {
      parts.pop()
      continue
    }

    parts.push(segment)
  }

  return parts.join('/')
}

function resolvePartPath(sourcePartPath: string, targetPath: string): string {
  if (targetPath.startsWith('/')) {
    return normalizeZipPath(targetPath.slice(1))
  }

  const sourceParts = sourcePartPath.split('/')
  sourceParts.pop()

  return normalizeZipPath([...sourceParts, targetPath].join('/'))
}

function getRequiredZipEntry(zip: ZipArchive, path: string): ZipEntry {
  const entry = zip.file(path)

  if (!entry) {
    throw new Error(`Missing archive entry: ${path}`)
  }

  return entry
}

function getMimeTypeFromPath(path: string): string {
  const extension = path.split('.').pop()?.toLowerCase()

  switch (extension) {
    case 'png':
      return 'image/png'
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg'
    case 'gif':
      return 'image/gif'
    case 'bmp':
      return 'image/bmp'
    case 'svg':
      return 'image/svg+xml'
    case 'webp':
      return 'image/webp'
    case 'tif':
    case 'tiff':
      return 'image/tiff'
    default:
      return 'application/octet-stream'
  }
}

function parseLengthToPixels(value: string | null): number | undefined {
  if (!value) {
    return undefined
  }

  const match = value.trim().match(/^(-?\d*\.?\d+)(cm|mm|in|pt|pc|px)?$/i)
  if (!match) {
    return undefined
  }

  const amount = Number(match[1])
  if (!Number.isFinite(amount)) {
    return undefined
  }

  const unit = (match[2] ?? 'px').toLowerCase()

  switch (unit) {
    case 'cm':
      return amount * CENTIMETERS_TO_PIXELS
    case 'mm':
      return amount * (CENTIMETERS_TO_PIXELS / 10)
    case 'in':
      return amount * 96
    case 'pt':
      return amount * (96 / 72)
    case 'pc':
      return amount * 16
    case 'px':
      return amount
    default:
      return undefined
  }
}

function getFrameBox(frame: Element): {
  x?: number
  y?: number
  w?: number
  h?: number
} {
  return {
    x: parseLengthToPixels(getAttributeValue(frame, ['svg:x', 'x'])),
    y: parseLengthToPixels(getAttributeValue(frame, ['svg:y', 'y'])),
    w: parseLengthToPixels(getAttributeValue(frame, ['svg:width', 'width'])),
    h: parseLengthToPixels(getAttributeValue(frame, ['svg:height', 'height'])),
  }
}

async function toDataUrl(zip: ZipArchive, path: string, signal: CancelSignal): Promise<string | null> {
  if (signal.cancelled) {
    return null
  }

  const entry = zip.file(path)
  if (!entry) {
    return null
  }

  const base64 = await entry.async('base64')

  if (signal.cancelled) {
    return null
  }

  return `data:${getMimeTypeFromPath(path)};base64,${base64}`
}

async function parseOdpSlides(
  zip: ZipArchive,
  signal: CancelSignal,
): Promise<ReadonlyArray<SlideData>> {
  const contentXml = await getRequiredZipEntry(zip, 'content.xml').async('string')

  if (signal.cancelled) {
    return []
  }

  const contentDocument = parseXml(contentXml)
  const pageLayoutProperties = getFirstByLocalName(contentDocument, 'page-layout-properties')
  const width = parseLengthToPixels(
    getAttributeValue(pageLayoutProperties ?? contentDocument.documentElement, ['fo:page-width', 'page-width']),
  ) ?? DEFAULT_SLIDE_WIDTH
  const height = parseLengthToPixels(
    getAttributeValue(pageLayoutProperties ?? contentDocument.documentElement, ['fo:page-height', 'page-height']),
  ) ?? DEFAULT_SLIDE_HEIGHT

  const slides: SlideData[] = []
  const pages = getElementsByLocalName(contentDocument, 'page')

  for (const [index, page] of pages.entries()) {
    if (signal.cancelled) {
      return slides
    }

    const processedParagraphs = new Set<Element>()
    const texts: Array<SlideData['texts'][number]> = []
    const images: Array<SlideData['images'][number]> = []

    for (const frame of getElementsByLocalName(page, 'frame')) {
      const paragraphs = getElementsByLocalName(frame, 'p')
      const frameText = paragraphs
        .map(paragraph => {
          processedParagraphs.add(paragraph)
          const text = (paragraph.textContent ?? '').trim()
          return text.length > 0 ? text : ''
        })
        .filter(Boolean)
        .join('\n')

      const box = getFrameBox(frame)

      if (frameText) {
        texts.push({
          text: frameText,
          x: box.x,
          y: box.y,
          w: box.w,
          h: box.h,
        })
      }

      const image = getFirstByLocalName(frame, 'image')
      const href = image ? getAttributeValue(image, ['xlink:href', 'href']) : null

      if (!href) {
        continue
      }

      const imagePath = resolvePartPath('content.xml', href)
      const source = await toDataUrl(zip, imagePath, signal)

      if (!source) {
        continue
      }

      images.push({
        src: source,
        x: box.x ?? 0,
        y: box.y ?? 0,
        w: box.w ?? width,
        h: box.h ?? height,
      })
    }

    for (const paragraph of getElementsByLocalName(page, 'p')) {
      if (processedParagraphs.has(paragraph)) {
        continue
      }

      const text = (paragraph.textContent ?? '').trim()
      if (!text) {
        continue
      }

      texts.push({ text })
    }

    slides.push({
      id: page.getAttribute('draw:name') ?? `slide-${index + 1}`,
      index,
      texts,
      images,
      width,
      height,
    })
  }

  return slides
}

export const OdpViewer = memo(OdpViewerBase)

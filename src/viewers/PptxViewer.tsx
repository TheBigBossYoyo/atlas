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

const EMU_PER_PIXEL = 9525
const DEFAULT_SLIDE_WIDTH = 1280
const DEFAULT_SLIDE_HEIGHT = 720

function PptxViewerBase({ file }: ViewerProps) {
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

        const nextSlides = await parsePptxSlides(zip as ZipArchive, signal)

        if (signal.cancelled) {
          return
        }

        setSlides(nextSlides)
        setActiveIndex(0)
        setError(nextSlides.length > 0 ? null : 'No slides found in PPTX.')
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
      <div className="pptx-viewer pptx-viewer--error">
        PptxViewer received a text file; expected binary.
      </div>
    )
  }

  if (error !== null) {
    return (
      <div className="pptx-viewer pptx-viewer--error">
        Failed to render PPTX: {error}
      </div>
    )
  }

  if (isLoading) {
    return <div className="pptx-viewer">Loading PPTX slides…</div>
  }

  return (
    <div className="pptx-viewer" style={{ width: '100%', height: '100%' }}>
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
    case 'wmf':
      return 'image/wmf'
    case 'emf':
      return 'image/emf'
    default:
      return 'application/octet-stream'
  }
}

function parseNumber(value: string | null): number | undefined {
  if (value === null) {
    return undefined
  }

  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

function emuToPx(value: string | null): number | undefined {
  const parsed = parseNumber(value)
  return parsed === undefined ? undefined : parsed / EMU_PER_PIXEL
}

function getTransform(root: Element): {
  x?: number
  y?: number
  w?: number
  h?: number
} {
  const transform = getFirstByLocalName(root, 'xfrm')
  const offset = transform ? getFirstByLocalName(transform, 'off') : null
  const extent = transform ? getFirstByLocalName(transform, 'ext') : null

  return {
    x: emuToPx(getAttributeValue(offset ?? root, ['x'])),
    y: emuToPx(getAttributeValue(offset ?? root, ['y'])),
    w: emuToPx(getAttributeValue(extent ?? root, ['cx', 'w'])),
    h: emuToPx(getAttributeValue(extent ?? root, ['cy', 'h'])),
  }
}

function getShapeText(shape: Element): string {
  const paragraphs = getElementsByLocalName(shape, 'p')
    .map(paragraph => {
      const text = getElementsByLocalName(paragraph, 't')
        .map(node => node.textContent ?? '')
        .join('')

      return text.trim().length > 0 ? text : ''
    })
    .filter(Boolean)

  return paragraphs.join('\n').trim()
}

async function parseRelationshipTargets(
  zip: ZipArchive,
  relationshipsPath: string,
  sourcePartPath: string,
  signal: CancelSignal,
): Promise<Map<string, string>> {
  if (signal.cancelled) {
    return new Map<string, string>()
  }

  const entry = zip.file(relationshipsPath)
  if (!entry) {
    return new Map<string, string>()
  }

  const xml = await entry.async('string')

  if (signal.cancelled) {
    return new Map<string, string>()
  }

  const document = parseXml(xml)
  const relationships = new Map<string, string>()

  for (const relationship of getElementsByLocalName(document, 'Relationship')) {
    const id = relationship.getAttribute('Id')
    const target = relationship.getAttribute('Target')
    const targetMode = relationship.getAttribute('TargetMode')

    if (!id || !target || targetMode === 'External') {
      continue
    }

    relationships.set(id, resolvePartPath(sourcePartPath, target))
  }

  return relationships
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

async function parsePptxSlides(
  zip: ZipArchive,
  signal: CancelSignal,
): Promise<ReadonlyArray<SlideData>> {
  const presentationXml = await getRequiredZipEntry(zip, 'ppt/presentation.xml').async('string')

  if (signal.cancelled) {
    return []
  }

  const presentationDocument = parseXml(presentationXml)
  const presentationRelationships = await parseRelationshipTargets(
    zip,
    'ppt/_rels/presentation.xml.rels',
    'ppt/presentation.xml',
    signal,
  )

  if (signal.cancelled) {
    return []
  }

  const slideSize = getFirstByLocalName(presentationDocument, 'sldSz')
  const width = emuToPx(getAttributeValue(slideSize ?? presentationDocument.documentElement, ['cx']))
    ?? DEFAULT_SLIDE_WIDTH
  const height = emuToPx(getAttributeValue(slideSize ?? presentationDocument.documentElement, ['cy']))
    ?? DEFAULT_SLIDE_HEIGHT

  const slidePaths = getElementsByLocalName(presentationDocument, 'sldId')
    .map(slideId => getAttributeValue(slideId, ['r:id', 'id']))
    .flatMap(relationshipId => {
      if (!relationshipId) {
        return []
      }

      const slidePath = presentationRelationships.get(relationshipId)
      return slidePath ? [slidePath] : []
    })

  const slides: SlideData[] = []

  for (const [index, slidePath] of slidePaths.entries()) {
    if (signal.cancelled) {
      return slides
    }

    const slideXml = await getRequiredZipEntry(zip, slidePath).async('string')
    const slideDocument = parseXml(slideXml)
    const slideFileName = slidePath.split('/').pop() ?? `slide${index + 1}.xml`
    const slideDirectory = slidePath.includes('/')
      ? slidePath.slice(0, slidePath.lastIndexOf('/'))
      : ''
    const relationshipTargets = await parseRelationshipTargets(
      zip,
      `${slideDirectory}/_rels/${slideFileName}.rels`,
      slidePath,
      signal,
    )
    const fallbackTargets = relationshipTargets.size > 0
      ? relationshipTargets
      : await parseRelationshipTargets(zip, `ppt/_rels/${slideFileName}.rels`, slidePath, signal)

    const processedTextNodes = new Set<Element>()
    const texts: Array<SlideData['texts'][number]> = []

    for (const shape of getElementsByLocalName(slideDocument, 'sp')) {
      const shapeTextNodes = getElementsByLocalName(shape, 't')
      if (shapeTextNodes.length === 0) {
        continue
      }

      for (const textNode of shapeTextNodes) {
        processedTextNodes.add(textNode)
      }

      const text = getShapeText(shape)
      if (!text) {
        continue
      }

      const transform = getTransform(shape)
      texts.push({
        text,
        x: transform.x,
        y: transform.y,
        w: transform.w,
        h: transform.h,
      })
    }

    for (const textNode of getElementsByLocalName(slideDocument, 't')) {
      if (processedTextNodes.has(textNode)) {
        continue
      }

      const text = (textNode.textContent ?? '').trim()
      if (!text) {
        continue
      }

      texts.push({ text })
    }

    const images: Array<SlideData['images'][number]> = []

    for (const picture of getElementsByLocalName(slideDocument, 'pic')) {
      const blip = getFirstByLocalName(picture, 'blip')
      const relationshipId = blip ? getAttributeValue(blip, ['r:embed', 'embed']) : null
      const imagePath = relationshipId ? fallbackTargets.get(relationshipId) : undefined

      if (!imagePath) {
        continue
      }

      const source = await toDataUrl(zip, imagePath, signal)
      if (!source) {
        continue
      }

      const transform = getTransform(picture)
      images.push({
        src: source,
        x: transform.x ?? 0,
        y: transform.y ?? 0,
        w: transform.w ?? width,
        h: transform.h ?? height,
      })
    }

    slides.push({
      id: `slide-${index + 1}`,
      index,
      texts,
      images,
      width,
      height,
    })
  }

  return slides
}

export const PptxViewer = memo(PptxViewerBase)

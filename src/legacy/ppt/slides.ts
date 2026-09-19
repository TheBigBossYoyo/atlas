/**
 * .ppt (PowerPoint 97-2003) slide extraction orchestrator (wave-4
 * legacy-office).
 *
 * Produces the app's shared `SlideData[]` model (`viewers/shared/
 * SlideDeck.types.ts`) so the existing `SlideDeck` component renders a
 * legacy deck exactly like it renders a PPTX/ODP one — just with each
 * slide's shapes simplified down to stacked text boxes.
 *
 * What this covers: per-slide title/body/other text runs, split into
 * paragraphs on `\r` (PPT's paragraph separator) with `\v` (0x0B) kept as an
 * embedded soft line break, laid out as simple full-width bands stacked
 * top-to-bottom.
 *
 * What this does NOT cover (all out of scope — see the file-level
 * limitation this wave accepts): real shape positions/sizes (this binary
 * format's actual EMU-equivalent layout is a full OfficeArt shape-tree
 * resolution well beyond a text-only preview, so nothing here is trusted —
 * every shape uses this module's own fixed stacked layout instead), images,
 * charts/SmartArt/tables/backgrounds, per-run character formatting (bold,
 * color, font, ...), speaker notes, and hidden-slide flags. Slide order is
 * the *physical* order `Slide` containers appear in the stream, which is
 * usually but not guaranteedly the presentation's visual order — resolving
 * the true order requires walking the separate slide-persistence/user-edit
 * structures, out of scope here.
 */
import type { SlideData, SlideParagraph, SlideShape, SlideTextBox } from '../../viewers/shared/SlideDeck.types'
import { BinaryReader } from '../binaryReader'
import { LegacyFormatError } from '../errors'
import { findEntry, readCfb } from '../cfb'
import { walkPptRecords, type PptRecordHeader } from './records'
import { collectSlideTextGroups, type PptTextGroup } from './text'

const RT_SLIDE = 1006

/**
 * Real decks stay far below this. Each empty slide container costs only 8
 * bytes, so without a cap a few MB of crafted records would make the viewer
 * build and render hundreds of thousands of slides.
 */
export const MAX_LEGACY_SLIDES = 5_000

// No real slide geometry survives this text-only extraction (see module
// header) — these mirror the PPTX/ODP parsers' own 16:9 fallback default
// (`viewers/slides/pptx/parser.ts`'s `DEFAULT_SLIDE_WIDTH`/`_HEIGHT`) so a
// legacy deck renders at the same size `SlideDeck` already expects when it
// has nothing better to go on.
const SLIDE_WIDTH = 1280
const SLIDE_HEIGHT = 720
const MARGIN = 48
const TITLE_LINE_HEIGHT = 40
const BODY_LINE_HEIGHT = 24
const GROUP_GAP = 16

// [MS-PPT] 2.4.14 TextHeaderAtom.textType — Title and CenterTitle.
const TITLE_TEXT_TYPES: ReadonlySet<number> = new Set([0, 6])

function paragraphsFromGroupText(text: string): ReadonlyArray<SlideParagraph> {
  return text
    .split('\r')
    .map((line) => line.replace(/\v/g, '\n'))
    .map((line) => ({ runs: [{ text: line }], level: 0 }))
}

function isBlankGroup(group: PptTextGroup): boolean {
  return group.text.replace(/[\r\v]/g, '').trim().length === 0
}

function isTitleGroup(group: PptTextGroup): boolean {
  return group.textType !== undefined && TITLE_TEXT_TYPES.has(group.textType)
}

function firstNonEmptyLine(text: string): string | undefined {
  for (const line of text.split(/[\r\v]/)) {
    const trimmed = line.trim()
    if (trimmed.length > 0) {
      return trimmed
    }
  }
  return undefined
}

/** Simple stacked layout: each non-blank text group becomes a full-width band, top to bottom — see module header on why no real positions are used. */
function buildShapes(groups: ReadonlyArray<PptTextGroup>): ReadonlyArray<SlideShape> {
  const shapes: SlideShape[] = []
  let cursorY = MARGIN

  groups.forEach((group, index) => {
    if (isBlankGroup(group)) {
      return
    }

    const paragraphs = paragraphsFromGroupText(group.text)
    const lineHeight = isTitleGroup(group) ? TITLE_LINE_HEIGHT : BODY_LINE_HEIGHT
    const height = Math.max(lineHeight, paragraphs.length * lineHeight)

    const shape: SlideTextBox = {
      kind: 'text',
      id: `text-${index}`,
      transform: { x: MARGIN, y: cursorY, w: SLIDE_WIDTH - MARGIN * 2, h: height },
      paragraphs,
      text: paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n'),
    }
    shapes.push(shape)
    cursorY += height + GROUP_GAP
  })

  return shapes
}

function findSlideContainers(reader: BinaryReader): ReadonlyArray<PptRecordHeader> {
  const slides: PptRecordHeader[] = []
  walkPptRecords(reader, 0, reader.length, (header) => {
    if (header.type === RT_SLIDE) {
      if (slides.length >= MAX_LEGACY_SLIDES) {
        throw new LegacyFormatError(`This presentation has more than ${MAX_LEGACY_SLIDES.toLocaleString('en-US')} slides, which is more than Atlas can show.`)
      }
      slides.push(header)
    }
  })
  return slides
}

function buildSlideData(reader: BinaryReader, header: PptRecordHeader, index: number): SlideData {
  try {
    const groups = collectSlideTextGroups(reader, header.dataStart, header.dataEnd)
    const titleGroup = groups.find((group) => isTitleGroup(group) && !isBlankGroup(group))

    return {
      id: `slide-${index + 1}`,
      index,
      title: titleGroup ? firstNonEmptyLine(titleGroup.text) : undefined,
      width: SLIDE_WIDTH,
      height: SLIDE_HEIGHT,
      shapes: buildShapes(groups),
    }
  } catch (err: unknown) {
    // One corrupted slide becomes a placeholder, not a discarded deck —
    // mirrors the ODP/PPTX parsers' own per-slide resilience (S9).
    return {
      id: `slide-${index + 1}`,
      index,
      width: SLIDE_WIDTH,
      height: SLIDE_HEIGHT,
      shapes: [],
      error: err instanceof Error ? err.message : 'This slide failed to load.',
    }
  }
}

/** Extracts every slide's text from a legacy PowerPoint 97-2003 (.ppt) file into the shared `SlideData[]` model. */
export function extractLegacyPptSlides(bytes: Uint8Array): ReadonlyArray<SlideData> {
  const cfb = readCfb(bytes)
  const entry = findEntry(cfb, 'PowerPoint Document')

  if (!entry || entry.type !== 'stream' || !entry.content) {
    throw new LegacyFormatError(
      'This file doesn\'t look like a valid PowerPoint 97-2003 presentation: missing the "PowerPoint Document" stream.',
    )
  }

  const reader = new BinaryReader(entry.content)
  return findSlideContainers(reader).map((header, index) => buildSlideData(reader, header, index))
}

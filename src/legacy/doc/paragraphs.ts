/**
 * Splits a raw decoded .doc story (still containing Word's in-band special
 * characters) into paragraphs, stripping field codes and the handful of
 * special characters called out for this wave: 0x07 (table cell/row mark),
 * 0x0D (paragraph mark), 0x0B (line break), 0x0C (page break), and
 * 0x13/0x14/0x15 (field begin/separator/end).
 *
 * Field codes are represented in-band as `0x13 <instructions> 0x14 <cached
 * result> 0x15` (the `0x14`/result half is omitted for a field with no
 * cached result). This keeps only the visible result text and drops the
 * instruction text (e.g. ` PAGE \* MERGEFORMAT `) — the same thing Word
 * itself shows when "field codes" display is off. Nesting is tracked only
 * by depth: a nested field's own instructions *and* result are both dropped
 * along with its parent's instructions (loses a nested field's result, e.g.
 * a REF inside a TOC entry) — an accepted best-effort simplification for a
 * read-only text preview, not a full field-evaluation engine.
 */
const CHAR_PARAGRAPH_MARK = 0x0d
const CHAR_CELL_MARK = 0x07
const CHAR_LINE_BREAK = 0x0b
const CHAR_PAGE_BREAK = 0x0c
const CHAR_FIELD_BEGIN = 0x13
const CHAR_FIELD_SEPARATOR = 0x14
const CHAR_FIELD_END = 0x15
const CHAR_TAB = 0x09

/**
 * Builds paragraph strings from slice-based runs rather than per-character
 * `+=` concatenation, so this stays linear on a large document instead of
 * quadratic (this wave's stated performance requirement for a 10MB .doc).
 */
export function splitIntoParagraphs(rawText: string): ReadonlyArray<string> {
  const paragraphs: string[] = []
  let currentParts: string[] = []
  let fieldDepth = 0
  let inFieldResult = false
  // Start index of the pending literal run, or -1 while nothing is pending
  // (either just flushed, or currently inside hidden field-instruction text).
  let segmentStart = -1

  const isVisible = (): boolean => fieldDepth === 0 || (fieldDepth === 1 && inFieldResult)

  const closeSegment = (endExclusive: number): void => {
    if (segmentStart !== -1 && endExclusive > segmentStart) {
      currentParts.push(rawText.slice(segmentStart, endExclusive))
    }
    segmentStart = -1
  }

  for (let i = 0; i < rawText.length; i += 1) {
    const code = rawText.charCodeAt(i)

    if (code === CHAR_FIELD_BEGIN) {
      closeSegment(i)
      fieldDepth += 1
      continue
    }
    if (code === CHAR_FIELD_SEPARATOR) {
      closeSegment(i)
      if (fieldDepth === 1) {
        inFieldResult = true
      }
      continue
    }
    if (code === CHAR_FIELD_END) {
      closeSegment(i)
      if (fieldDepth === 1) {
        inFieldResult = false
      }
      fieldDepth = Math.max(0, fieldDepth - 1)
      continue
    }

    if (!isVisible()) {
      continue // inside field instructions (or a fully nested field) — not visible text.
    }

    if (code === CHAR_PARAGRAPH_MARK || code === CHAR_CELL_MARK) {
      closeSegment(i)
      paragraphs.push(currentParts.join(''))
      currentParts = []
      continue
    }
    if (code === CHAR_LINE_BREAK || code === CHAR_PAGE_BREAK) {
      closeSegment(i)
      currentParts.push('\n')
      continue
    }
    if (code < 0x20 && code !== CHAR_TAB) {
      closeSegment(i) // strip other control characters (bookmark/annotation reference marks, etc.).
      continue
    }

    if (segmentStart === -1) {
      segmentStart = i
    }
  }

  if (isVisible()) {
    closeSegment(rawText.length)
  }
  if (currentParts.length > 0) {
    paragraphs.push(currentParts.join(''))
  }

  return paragraphs
}

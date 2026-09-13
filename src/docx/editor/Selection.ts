export type Position = Readonly<{
  paragraphPath: ReadonlyArray<number>
  runIndex: number
  charOffset: number
}>

export type Range = Readonly<{
  anchor: Position
  focus: Position
}>

function compareParagraphPaths(
  a: ReadonlyArray<number>,
  b: ReadonlyArray<number>,
): -1 | 0 | 1 {
  const sharedLength = Math.min(a.length, b.length)

  for (let index = 0; index < sharedLength; index += 1) {
    if (a[index] < b[index]) {
      return -1
    }

    if (a[index] > b[index]) {
      return 1
    }
  }

  if (a.length < b.length) {
    return -1
  }

  if (a.length > b.length) {
    return 1
  }

  return 0
}

export function comparePositions(a: Position, b: Position): -1 | 0 | 1 {
  const paragraphComparison = compareParagraphPaths(a.paragraphPath, b.paragraphPath)
  if (paragraphComparison !== 0) {
    return paragraphComparison
  }

  if (a.runIndex < b.runIndex) {
    return -1
  }

  if (a.runIndex > b.runIndex) {
    return 1
  }

  if (a.charOffset < b.charOffset) {
    return -1
  }

  if (a.charOffset > b.charOffset) {
    return 1
  }

  return 0
}

export function positionEquals(a: Position, b: Position): boolean {
  return comparePositions(a, b) === 0
}

export function isRangeCollapsed(range: Range): boolean {
  return positionEquals(range.anchor, range.focus)
}

export function normalizeRange(
  range: Range,
): Readonly<{
  start: Position
  end: Position
}> {
  if (comparePositions(range.anchor, range.focus) <= 0) {
    return {
      start: range.anchor,
      end: range.focus,
    }
  }

  return {
    start: range.focus,
    end: range.anchor,
  }
}

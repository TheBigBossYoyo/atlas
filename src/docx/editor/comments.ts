/**
 * Atlas — DOCX Comments helpers (Wave E.2)
 *
 * Helpers for the CommentsPane UI:
 *   - extractCommentText: pulls plain text from the parser's `UnknownNode`
 *     payload (currently stored as JSON.stringify of the raw `<w:comment>`
 *     element).  Walks all `w:t` text nodes inside the JSON tree.
 *   - findCommentAnchors: walks a Document's sections and returns, for each
 *     comment id, the paragraphPath of the FIRST anchor (range-start or
 *     reference) so the UI can scroll to it on click.
 */

import type {
  Block,
  Comment as CommentNode,
  Document as DocxDocument,
  HyperlinkChild,
  ParagraphChild,
  Run,
  RunChild,
} from '../model/document'

// ---------------------------------------------------------------------------
// Comment body text extraction
// ---------------------------------------------------------------------------

interface RawXmlNode {
  readonly [key: string]: unknown
}

function isPlainObject(value: unknown): value is RawXmlNode {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function collectTextFromRawXml(node: unknown, out: string[]): void {
  if (typeof node === 'string') {
    out.push(node)
    return
  }

  if (typeof node === 'number' || typeof node === 'boolean') {
    out.push(String(node))
    return
  }

  if (Array.isArray(node)) {
    for (const child of node) {
      collectTextFromRawXml(child, out)
    }
    return
  }

  if (!isPlainObject(node)) {
    return
  }

  for (const [key, value] of Object.entries(node)) {
    // Skip XML attributes (prefixed with @_) and the synthetic #text wrapper
    // is fine to include via recursion below.
    if (key.startsWith('@_')) {
      continue
    }

    // Suppress deleted-text and instruction-text noise; keep w:t and #text.
    if (key === 'w:delText' || key === 'w:instrText') {
      continue
    }

    collectTextFromRawXml(value, out)
  }
}

/**
 * Extract a flat plain-text rendition of a parsed Comment node.
 *
 * Wave A.5 stores the comment body as a single `UnknownNode` whose `xml`
 * field is `JSON.stringify(rawCommentObject)`.  We re-parse that JSON,
 * walk all `w:t` (and #text) leaves, join with a single space between
 * paragraphs, and trim whitespace.
 */
export function extractCommentText(comment: CommentNode): string {
  const parts: string[] = []
  const structuredBody = (comment as CommentNode & { readonly body?: ReadonlyArray<Block> }).body
  const legacyBlocks = (comment as CommentNode & { readonly blocks?: ReadonlyArray<Block> }).blocks
  const blocks = Array.isArray(structuredBody)
    ? structuredBody
    : Array.isArray(legacyBlocks)
      ? legacyBlocks
      : []

  for (const block of blocks) {
    if (block.kind === 'paragraph') {
      parts.push(extractParagraphText(block.children))
      continue
    }

    if (block.kind === 'unknown') {
      const text = extractRawCommentText(block.xml)
      if (text.length > 0) {
        parts.push(text)
      }
    }
  }

  return parts
    .map(part => part.replace(/\s+/g, ' ').trim())
    .filter(part => part.length > 0)
    .join(' ')
}

function extractRawCommentText(rawJson: string): string {
  try {
    const parsed = JSON.parse(rawJson) as unknown
    const paragraphs = getRawParagraphNodes(parsed)
    if (paragraphs.length === 0) {
      const textParts: string[] = []
      collectTextFromRawXml(parsed, textParts)
      return textParts.join(' ')
    }

    return paragraphs
      .map((paragraph) => {
        const textParts: string[] = []
        collectTextFromRawXml(paragraph, textParts)
        return textParts.join('')
      })
      .join(' ')
  } catch {
    return ''
  }
}

function getRawParagraphNodes(node: unknown): ReadonlyArray<unknown> {
  if (!isPlainObject(node)) {
    return []
  }

  const directParagraphs = node['w:p']
  if (Array.isArray(directParagraphs)) {
    return directParagraphs
  }

  if (directParagraphs !== undefined) {
    return [directParagraphs]
  }

  const commentNode = node['w:comment']
  if (commentNode !== undefined) {
    return getRawParagraphNodes(commentNode)
  }

  return []
}

function extractParagraphText(children: ReadonlyArray<ParagraphChild>): string {
  return children.map(extractInlineText).join('')
}

function extractInlineText(child: ParagraphChild): string {
  if (child.kind === 'run') {
    return extractRunText(child)
  }

  if (child.kind === 'hyperlink') {
    return child.children.map(extractHyperlinkChildText).join('')
  }

  if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
    return getRevisionChildren(child).map(extractInlineText).join('')
  }

  return ''
}

function extractHyperlinkChildText(child: HyperlinkChild): string {
  if (child.kind === 'run') {
    return extractRunText(child)
  }

  return ''
}

function extractRunText(run: Run): string {
  return run.children.map(extractRunChildText).join('')
}

function extractRunChildText(child: RunChild): string {
  if (child.kind === 'text') {
    return child.value
  }

  if (child.kind === 'tab') {
    return ' '
  }

  if (child.kind === 'break') {
    return '\n'
  }

  return ''
}

// ---------------------------------------------------------------------------
// Anchor discovery
// ---------------------------------------------------------------------------

export interface CommentAnchor {
  readonly id: string
  readonly paragraphPath: ReadonlyArray<number>
}

function paragraphContainsCommentAnchor(
  children: ReadonlyArray<ParagraphChild>,
): ReadonlyArray<{ id: string; priority: number }> {
  const found: Array<{ id: string; priority: number }> = []

  for (const child of children) {
    if (child.kind === 'comment-range' && child.boundary === 'start') {
      found.push({ id: child.id, priority: 0 })
      continue
    }

    if (child.kind === 'comment-reference') {
      found.push({ id: child.id, priority: 1 })
      continue
    }

    if (child.kind === 'run') {
      for (const grandchild of child.children) {
        if (grandchild.kind === 'comment-reference') {
          found.push({ id: grandchild.id, priority: 1 })
        }
      }
      continue
    }

    if (child.kind === 'hyperlink') {
      for (const grandchild of child.children) {
        if (grandchild.kind === 'comment-range' && grandchild.boundary === 'start') {
          found.push({ id: grandchild.id, priority: 0 })
        } else if (grandchild.kind === 'comment-reference') {
          found.push({ id: grandchild.id, priority: 1 })
        } else if (grandchild.kind === 'run') {
          for (const greatGrandchild of grandchild.children) {
            if (greatGrandchild.kind === 'comment-reference') {
              found.push({ id: greatGrandchild.id, priority: 1 })
            }
          }
        }
      }
      continue
    }

    if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
      found.push(...paragraphContainsCommentAnchor(getRevisionChildren(child)))
    }
  }

  return found
}

function getRevisionChildren(
  revision: Extract<ParagraphChild, { kind: 'ins-revision' | 'del-revision' }>,
): ReadonlyArray<ParagraphChild> {
  return revision.children as ReadonlyArray<ParagraphChild>
}

function visitBlocks(
  blocks: ReadonlyArray<Block>,
  paragraphIndexRef: { current: number },
  hits: Map<string, { paragraphIndex: number; priority: number }>,
): void {
  for (const block of blocks) {
    if (block.kind === 'paragraph') {
      const matches = paragraphContainsCommentAnchor(block.children)
      for (const match of matches) {
        const existing = hits.get(match.id)
        if (existing === undefined || match.priority < existing.priority) {
          hits.set(match.id, {
            paragraphIndex: paragraphIndexRef.current,
            priority: match.priority,
          })
        }
      }
      paragraphIndexRef.current += 1
      continue
    }

    if (block.kind === 'table') {
      for (const row of block.rows) {
        if (row.kind === 'table-row') {
          for (const cell of row.cells) {
            if (cell.kind === 'table-cell') {
              visitBlocks(cell.blocks, paragraphIndexRef, hits)
            }
          }
        }
      }
    }
  }
}

/**
 * Build a map from commentId → flattened paragraph index of its first anchor.
 * Paragraph indexing matches the order produced by PageView (sequential
 * across sections, descending into tables).
 */
export function findCommentAnchors(
  document: DocxDocument,
): ReadonlyMap<string, number> {
  const hits = new Map<string, { paragraphIndex: number; priority: number }>()
  const paragraphIndexRef = { current: 0 }

  for (const section of document.sections) {
    visitBlocks(section.blocks, paragraphIndexRef, hits)
  }

  const result = new Map<string, number>()
  for (const [id, hit] of hits) {
    result.set(id, hit.paragraphIndex)
  }
  return result
}

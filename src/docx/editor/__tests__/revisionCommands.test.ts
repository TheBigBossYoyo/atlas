import { describe, expect, it } from 'vitest'

import {
  acceptAllRevisions,
  acceptRevision,
  applyCommand,
  rejectAllRevisions,
  rejectRevision,
} from '../index'
import type {
  Document,
  Paragraph,
  ParagraphChild,
  Run,
  Section,
} from '../../model'

describe('revision command helpers', () => {
  it('accepts and rejects individual revisions by id', () => {
    const document = createDocument([
      createParagraphWithRevision('ins-revision', 'ins-1', 'kept'),
      createParagraphWithRevision('del-revision', 'del-1', 'restored'),
    ])

    const accepted = applyCommand(document, acceptRevision('ins-1'))
    const rejected = applyCommand(document, rejectRevision('del-1'))

    expect(paragraphTexts(accepted.document)).toEqual([['kept'], []])
    expect(paragraphTexts(rejected.document)).toEqual([[], ['restored']])
  })

  it('accepts and rejects all revisions with helper commands', () => {
    const document = createDocument([
      createParagraphWithRevision('ins-revision', 'ins-1', 'alpha'),
      createParagraphWithRevision('del-revision', 'del-1', 'beta'),
    ])

    const accepted = applyCommand(document, acceptAllRevisions())
    const rejected = applyCommand(document, rejectAllRevisions())

    expect(paragraphTexts(accepted.document)).toEqual([['alpha'], []])
    expect(paragraphTexts(rejected.document)).toEqual([[], ['beta']])
  })
})

function createDocument(paragraphs: ReadonlyArray<Paragraph>): Document {
  const section: Section = {
    kind: 'section',
    props: {},
    blocks: Object.freeze([...paragraphs]),
  }

  return {
    kind: 'document',
    sections: Object.freeze([section]),
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

function createParagraphWithRevision(
  kind: 'ins-revision' | 'del-revision',
  id: string,
  text: string,
): Paragraph {
  const run: Run = {
    kind: 'run',
    children: [{ kind: 'text', value: text }],
  }
  const revisionChild: ParagraphChild = {
    kind,
    id,
    children: Object.freeze([run]) as ReadonlyArray<ParagraphChild> & ReadonlyArray<Run>,
  }

  return {
    kind: 'paragraph',
    children: [revisionChild],
  }
}

function paragraphTexts(document: Document): string[][] {
  return document.sections[0].blocks.map((block) => {
    if (block.kind !== 'paragraph') {
      return []
    }

    return block.children
      .filter((child): child is Run => child.kind === 'run')
      .map((child) =>
        child.children.map((runChild) => runChild.kind === 'text' ? runChild.value : '').join(''),
      )
      .filter((text) => text.length > 0)
  })
}

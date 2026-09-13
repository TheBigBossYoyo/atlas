import { XMLBuilder } from 'fast-xml-parser'

import type { Comment } from '../model'
import type { OrderedXmlNode } from './partWriterSupport'
import { XML_DECLARATION, WORD_NAMESPACE, buildBlockNodes } from './partWriterSupport'

const WORD_2012_NAMESPACE = 'http://schemas.microsoft.com/office/word/2012/wordml'

const orderedXmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

export function writeCommentsXml(comments: ReadonlyArray<Comment>): string {
  const needsWord2012Namespace = comments.some((comment) => comment.parentId !== undefined)
  const root: OrderedXmlNode = {
    'w:comments': comments.map(buildCommentNode),
    ':@': {
      '@_xmlns:w': WORD_NAMESPACE,
      ...(needsWord2012Namespace ? { '@_xmlns:w15': WORD_2012_NAMESPACE } : {}),
    },
  }

  return `${XML_DECLARATION}${orderedXmlBuilder.build([root])}`
}

function buildCommentNode(comment: Comment): OrderedXmlNode {
  return {
    'w:comment': [...buildBlockNodes(comment.body)],
    ':@': {
      '@_w:id': comment.id,
      ...(comment.author !== undefined ? { '@_w:author': comment.author } : {}),
      ...(comment.date !== undefined ? { '@_w:date': comment.date } : {}),
      ...(comment.initials !== undefined ? { '@_w:initials': comment.initials } : {}),
      ...(comment.parentId !== undefined ? { '@_w15:parentId': comment.parentId } : {}),
    },
  }
}

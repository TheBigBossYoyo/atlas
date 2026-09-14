import { XMLBuilder } from 'fast-xml-parser'

import type { Comment } from '../model'
import type { OrderedXmlNode } from './partWriterSupport'
import { XML_DECLARATION, buildBlockNodes } from './partWriterSupport'
import { buildNamespaceDeclarationAttributes, STANDARD_NAMESPACE_URIS } from './documentWriter'

const orderedXmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  format: false,
  processEntities: true,
  suppressEmptyNode: true,
})

/**
 * D19 / DXS-08: declares the full standard namespace set (matching
 * `documentWriter.ts`'s document root and `partWriterSupport.ts`'s other
 * standalone parts) rather than just `xmlns:w` (+ a conditional `xmlns:w15`
 * for `w15:parentId`) — a comment body can contain a table (wave 1
 * follow-up) or, via `buildBlockNodes`'s reuse of `documentWriter.ts`'s own
 * builders, a drawing/hyperlink/revision, any of which needs a namespace
 * prefix beyond `w`.
 */
export function writeCommentsXml(comments: ReadonlyArray<Comment>): string {
  const root: OrderedXmlNode = {
    'w:comments': comments.map(buildCommentNode),
    ':@': buildNamespaceDeclarationAttributes(STANDARD_NAMESPACE_URIS),
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

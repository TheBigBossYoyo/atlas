export type { Command } from './commandTypes'
export * from './commandTypes'

export { applyCommand } from './commands'
export { findParagraph, findEnclosingTable } from './commands'
export type { EnclosingTable } from './commands'
export { History } from './History'

export {
  handleBeforeInput,
  handleKeyDown,
  moveCursorLeft,
  moveCursorRight,
  moveCursorToLineStart,
  moveCursorToLineEnd,
  moveCursorToDocStart,
  moveCursorToDocEnd,
  extendOrCollapse,
} from './Input'
export type { InputContext, InputResult } from './Input'

export { findAll, findNext, findPrev, buildReplaceCommands, buildReplaceAllCommands } from './Find'
export type { FindMatch, FindOptions } from './Find'
export { FindReplace } from './FindReplace'
export type { CompositionState } from './Composition'
export { useComposition } from './useComposition'
export { toolbarToCommand } from './toolbarAdapter'
export { CommentsPane } from './CommentsPane'
export type { CommentsPaneProps } from './CommentsPane'
export { extractCommentText, findCommentAnchors } from './comments'
export type { CommentAnchor } from './comments'
export { SpellCheckMenu } from './SpellCheckMenu'
export type { SpellCheckMenuProps } from './SpellCheckMenu'
export { useSpellCheck } from './useSpellCheck'
export type { SpellCheckState, UseSpellCheckResult } from './useSpellCheck'
export { insertImageIntoBundle, decodeImageNaturalSizePt } from './insertImage'
export type { InsertImageInput, InsertImageResult, ImageMimeType } from './insertImage'
export { insertHyperlinkIntoBundle } from './insertHyperlink'
export type { InsertHyperlinkResult } from './insertHyperlink'
export { ensureListNumbering } from './insertList'
export type { ListKind } from './insertList'
export { buildSpellCheckReplacement } from './spellCheckReplace'
export type { SpellCheckReplacement } from './spellCheckReplace'
export { buildPasteCommands, htmlToParagraphs, textToParagraphs } from './htmlPaste'
export type { ParsedParagraph, ParsedRun, PasteCommandsResult } from './htmlPaste'
export { friendlyDocxErrorMessage } from './friendlyDocxError'
export { htmlToPasteBlocks } from './pasteBlocks'
export type { PasteBlock, PasteParagraph, PasteTable } from './pasteBlocks'
export { buildRichPasteCommands, bundleContextFor } from './pasteRich'
export type { RichPasteBundleContext, RichPasteBundlePatch, RichPasteResult } from './pasteRich'

// D29 — header/footer editing.
export { listHeaderFooterParts, setHeaderFooterText, blocksToText, type HeaderFooterPart, type HeaderFooterKind } from './headerFooter'

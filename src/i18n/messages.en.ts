import type { MessageCatalogue } from './types';

/**
 * English catalogue — the source of truth for every key. `messages.fr.ts`
 * must define exactly the same key set (enforced by
 * `src/i18n/__tests__/catalogueParity.test.ts`).
 *
 * Keys are grouped by the component/module that owns them, in roughly the
 * order the shell renders them, so a reviewer can find a string by opening
 * the matching component and searching this file's comments.
 */
export const messages = {
  // Shared across several components.
  'common.unsavedChanges': 'Unsaved changes',
  'common.dismiss': 'Dismiss',

  // WelcomeScreen.tsx
  'welcome.subtitle': 'A universal document viewer for Word, Excel, PowerPoint, PDF, code, markdown, and more.',
  'welcome.openFile': 'Open File',
  'welcome.loadSample': 'Load Sample Document',
  'welcome.dragDropHint': 'or drag & drop any document anywhere',
  'welcome.recentFiles': 'Recent files',
  'welcome.removeRecentAria': 'Remove {name} from recent',
  'welcome.removeRecentTitle': 'Remove from recent',
  'welcome.relative.justNow': 'just now',
  'welcome.relative.minutesAgo': { one: '{count}m ago', other: '{count}m ago' },
  'welcome.relative.hoursAgo': { one: '{count}h ago', other: '{count}h ago' },
  'welcome.relative.daysAgo': { one: '{count}d ago', other: '{count}d ago' },
  'welcome.feature.office': 'Word, Excel, PowerPoint, PDF',
  'welcome.feature.markdown': 'Markdown + code highlighting',
  'welcome.feature.math': 'Math + Mermaid diagrams',
  'welcome.feature.themes': '5 beautiful themes',
  'welcome.feature.search': 'In-document search',
  'welcome.feature.export': 'PDF export, plus DOCX/HTML for Markdown',

  // DropZone.tsx
  'dropzone.title': 'Drop your document',
  'dropzone.detail': 'Word, Excel, PowerPoint, PDF, OpenDocument, Markdown, CSV, RTF, and code files',

  // ViewerErrorBoundary.tsx
  'viewerError.title': 'Viewer crashed',
  'viewerError.tryAgain': 'Try again',

  // ViewerLoading.tsx
  'viewerLoading.loading': 'Loading {format}…',
  'viewerLoading.documentFallback': 'document',

  // Sidebar.tsx
  'sidebar.title': 'Table of Contents',
  'sidebar.empty': 'No outline available',

  // RawEditor.tsx — I18N-1. The header label doubles as the textarea's
  // aria-label so the accessible name survives once the placeholder
  // disappears (WCAG 3.3.2).
  'rawEditor.title': 'Markdown Source',
  'rawEditor.placeholder': 'Type or paste markdown here...',

  // Toolbar.tsx
  'toolbar.showSidebar': 'Show sidebar',
  'toolbar.hideSidebar': 'Hide sidebar',
  'toolbar.toggleSidebarAria': 'Toggle sidebar',
  'toolbar.open': 'Open',
  'toolbar.openTitle': 'Open file (Ctrl+O)',
  'toolbar.save': 'Save',
  'toolbar.saveTitle': 'Save (Ctrl+S)',
  'toolbar.closeFileAria': 'Close file',
  'toolbar.closeFileTitle': 'Close file (Ctrl+W)',
  'toolbar.viewMode.preview': 'Preview',
  'toolbar.viewMode.split': 'Split',
  'toolbar.viewMode.editor': 'Editor',
  'toolbar.fontSizeGroupAria': 'Font size',
  'toolbar.decreaseFontAria': 'Decrease font size',
  'toolbar.decreaseFontTitle': 'Decrease font size (Ctrl+-)',
  'toolbar.resetFontTitle': 'Reset font size (Ctrl+0)',
  'toolbar.increaseFontAria': 'Increase font size',
  'toolbar.increaseFontTitle': 'Increase font size (Ctrl+=)',
  'toolbar.searchAria': 'Search',
  'toolbar.searchTitle': 'Search (Ctrl+F)',
  'toolbar.shortcutsAria': 'Keyboard shortcuts',
  'toolbar.shortcutsTitle': 'Keyboard shortcuts (Ctrl+/)',

  // NewDocumentMenu.tsx
  'newDocument.trigger': 'New',
  'newDocument.triggerTitle': 'New document (Ctrl+N)',
  'newDocument.menuLabel': 'New document',
  'newDocument.markdown': 'Markdown (.md)',
  'newDocument.docx': 'Word document (.docx)',
  'newDocument.xlsx': 'Excel workbook (.xlsx)',
  'newDocument.ods': 'OpenDocument Spreadsheet (.ods)',
  'newDocument.pptx': 'PowerPoint presentation (.pptx)',
  'newDocument.odp': 'OpenDocument Presentation (.odp)',

  // ExportMenu.tsx
  'exportMenu.triggerAria': 'Export',
  'exportMenu.triggerTitle': 'Export',
  'exportMenu.menuLabel': 'Export options',
  'exportMenu.html': 'HTML',
  'exportMenu.pdf': 'PDF',
  'exportMenu.docx': 'DOCX',
  'exportMenu.markdown': 'Markdown',
  'exportMenu.saveCopy': 'Save a copy',
  'exportMenu.exportToPdf': 'Export to PDF',
  'exportMenu.exportToCsv': 'Export to CSV',
  'exportMenu.exportToHtml': 'Export to HTML',

  // ThemeMenu.tsx
  'theme.triggerAria': 'Theme',
  'theme.triggerTitle': 'Theme',
  'theme.menuLabel': 'Theme options',

  // LanguageMenu.tsx
  'language.triggerAria': 'Language',
  'language.triggerTitle': 'Language',
  'language.menuLabel': 'Language options',
  'language.english': 'English',
  'language.french': 'French',
  'language.system': 'System',

  // TabBar.tsx
  'tabBar.label': 'Open documents',
  'tabBar.closeAria': 'Close {name}',
  'tabBar.closeTitle': 'Close {name}',

  // ShortcutsModal.tsx
  'shortcuts.title': 'Keyboard Shortcuts',
  'shortcuts.closeAria': 'Close',
  'shortcuts.sectionFile': 'File',
  'shortcuts.newDocumentMenu': 'New document menu',
  'shortcuts.openFile': 'Open file',
  'shortcuts.save': 'Save',
  'shortcuts.saveAs': 'Save As',
  'shortcuts.closeFile': 'Close file',
  'shortcuts.nextPrevTab': 'Next / previous tab',
  'shortcuts.reopenClosed': 'Reopen last closed document',
  'shortcuts.printExport': 'Print / export',
  'shortcuts.exportMenu': 'Export menu',
  'shortcuts.sectionView': 'View',
  'shortcuts.preview': 'Preview',
  'shortcuts.split': 'Split',
  'shortcuts.editor': 'Editor',
  'shortcuts.toggleSidebar': 'Toggle sidebar',
  'shortcuts.cycleTheme': 'Cycle theme',
  'shortcuts.sectionEdit': 'Edit',
  'shortcuts.increaseFont': 'Increase font',
  'shortcuts.decreaseFont': 'Decrease font',
  'shortcuts.resetFont': 'Reset font',
  'shortcuts.sectionSearchHelp': 'Search & Help',
  'shortcuts.find': 'Find (Markdown, Text, Code, RTF, ODT)',
  'shortcuts.nextPrev': 'Next/prev',
  'shortcuts.closeDialog': 'Close',
  'shortcuts.toggleDialog': 'Toggle this dialog',
  'shortcuts.note': "While editing a DOCX, its own editor shortcuts (bold/italic/underline, alignment, find/replace, line spacing, undo/redo, save, print) take priority over the shortcuts above.",

  // UnsavedChangesDialog.tsx
  'unsavedDialog.title': 'Unsaved changes',
  'unsavedDialog.body': 'This document has unsaved changes. Save them before continuing, discard them, or cancel.',
  'unsavedDialog.cancel': 'Cancel',
  'unsavedDialog.discard': 'Discard',
  'unsavedDialog.save': 'Save',
  'unsavedDialog.saving': 'Saving…',

  // DraftRecoveryBanner.tsx
  'draftRecovery.foundNamed': 'Unsaved changes to "{fileName}" from {savedAt} were found.',
  'draftRecovery.foundUnnamed': 'An unsaved draft from {savedAt} was found.',
  'draftRecovery.discard': 'Discard',
  'draftRecovery.restore': 'Restore',

  // FileStatusBanner.tsx
  'fileStatus.opening': 'Opening file…',
  'fileStatus.dismissErrorAria': 'Dismiss error',
  'fileStatus.dismissErrorTitle': 'Dismiss',

  // StatusBar.tsx
  'statusBar.untitled': 'Untitled',
  'statusBar.unsavedChangesTitle': 'Unsaved changes',
  'statusBar.words': { one: '{count} word', other: '{count} words' },
  'statusBar.headings': { one: '{count} heading', other: '{count} headings' },
  'statusBar.markdownStats': '{words} · {headings}',
  'statusBar.rows': { one: '{count} row', other: '{count} rows' },
  'statusBar.cols': { one: '{count} col', other: '{count} cols' },
  'statusBar.spreadsheetStats': 'Sheet: {sheet} · {rows} × {cols}',
  'statusBar.pdfStats': 'Page {page} / {pageCount}',
  'statusBar.slidesStats': 'Slide {slide} / {slideCount}',
  'statusBar.lines': { one: '{count} line', other: '{count} lines' },
  'statusBar.codeStats': '{language} · {lines}',
  'statusBar.chars': { one: '{count} char', other: '{count} chars' },
  'statusBar.textStats': '{lines} · {chars}',
  'statusBar.pages': { one: '{count} page', other: '{count} pages' },
  'statusBar.documentStats': '{words} · {pages}',

  // Toast.tsx
  'toast.viewportAria': 'Notifications',
  'toast.dismissAria': 'Dismiss notification',

  // SearchOverlay.tsx
  'search.placeholder': 'Search in document...',
  'search.noResults': 'No results',
  'search.resultCount': '{current} of {total}',
  'search.previousAria': 'Previous match',
  'search.previousTitle': 'Previous (Shift+Enter)',
  'search.nextAria': 'Next match',
  'search.nextTitle': 'Next (Enter)',
  'search.closeAria': 'Close search',
  'search.closeTitle': 'Close (Esc)',

  // EmptyFileNotice.tsx
  // LargeMarkdownNotice.tsx
  'largeMarkdown.title': "This document is too large to preview quickly",
  'largeMarkdown.detail': "Rendering these {size} MB of Markdown can take a while. The editor is open and fully usable; you can render the preview anyway if you need it.",
  'largeMarkdown.renderAnyway': "Render the preview anyway",

  // MarkdownRenderer.tsx / useMarkdownHastTree.ts
  'markdownWorker.parsing': 'Rendering large document…',
  'markdownWorker.error': "This document couldn't be rendered: {error}",

  'emptyFile.title': 'This file is empty',
  'emptyFile.detail': "{fileName} has no content yet, and Atlas doesn't have a blank template for this format. Open it in another app to create content first.",

  // LegacyFormatBanner.tsx
  'legacyFormat.notice': '{formatLabel} — legacy format, shown as read-only text. Re-save it as {modernExtension} in Word, PowerPoint, or a compatible app for full fidelity editing.',

  // CodeViewer.tsx
  'codeViewer.toolbarAria': 'Code editor',
  'codeViewer.stop': 'Stop',
  'codeViewer.run': 'Run',
  'codeViewer.stopTitle': 'Stop the running program',
  'codeViewer.runTitle': 'Save and run this file',
  'codeViewer.findReplaceAria': 'Find and replace',
  'codeViewer.findReplaceTitle': 'Find and replace (Ctrl+F)',
  'codeViewer.goToLineAria': 'Go to line',
  'codeViewer.goToLineTitle': 'Go to line (Ctrl+G)',
  'codeViewer.wordWrapAria': 'Word wrap',
  'codeViewer.saveAria': 'Save',
  'codeViewer.saveTitle': 'Save (Ctrl+S)',
  'codeViewer.saveCancelled': 'Save was cancelled or unavailable.',
  'codeViewer.outputRegionAria': 'Program output',
  'codeViewer.running': 'Running…',
  'codeViewer.output': 'Output',
  'codeViewer.clearOutputAria': 'Clear output',
  'codeViewer.closeOutputAria': 'Close output',

  // useCodeRun.ts
  'codeRun.stopped': 'Stopped.',
  'codeRun.timedOut': 'Stopped after the 60 second time limit.',
  'codeRun.finishedOk': 'Finished (exit code 0).',
  'codeRun.exitedWithCode': 'Exited with code {code}.',
  'codeRun.couldNotStart': 'The program could not be started.',

  // PdfToolbar.tsx
  'pdfToolbar.toggleThumbnails': 'Toggle page thumbnails',
  'pdfToolbar.previousPage': 'Previous Page (Up / PageUp)',
  'pdfToolbar.nextPage': 'Next Page (Down / PageDown)',
  'pdfToolbar.pageNumberAria': 'Page number',
  'pdfToolbar.find': 'Find (Ctrl+F)',
  'pdfToolbar.rotate': 'Rotate page',
  'pdfToolbar.print': 'Print (Ctrl+P)',
  'pdfToolbar.zoomOut': 'Zoom Out (Ctrl+-)',
  'pdfToolbar.zoomIn': 'Zoom In (Ctrl++)',
  'pdfToolbar.fitWidth': 'Fit Width',
  'pdfToolbar.fitPage': 'Fit Page',

  // SpreadsheetViewer.tsx
  'spreadsheet.format.xlsx': 'Excel Workbook (.xlsx)',
  'spreadsheet.format.xlsm': 'Excel Macro-Enabled Workbook (.xlsm)',
  'spreadsheet.format.xlsb': 'Excel Binary Workbook (.xlsb)',
  'spreadsheet.format.xls': 'Excel 97-2003 Workbook (.xls)',
  'spreadsheet.format.ods': 'OpenDocument Spreadsheet (.ods)',
  'spreadsheet.format.fods': 'Flat OpenDocument Spreadsheet (.fods)',

  // CsvViewer.tsx
  'csv.format.csv': 'CSV (comma-separated)',
  'csv.format.tsv': 'TSV (tab-separated)',
  'spreadsheet.hideHiddenSheets': 'Hide hidden sheets',
  'spreadsheet.showHiddenSheets': 'Show hidden sheets',
  'spreadsheet.searchRowsPlaceholder': 'Search rows...',
  'spreadsheet.rowsCount': { one: '{count} row', other: '{count} rows' },
  'spreadsheet.colsCount': { one: '{count} column', other: '{count} columns' },
  'spreadsheet.dimensions': '{rows} × {cols}',
  'spreadsheet.noSearchResults': 'No rows match your search.',
  'spreadsheet.emptySheet': 'This sheet is empty.',
  'spreadsheet.renameSheetAria': 'Rename sheet {name}',
  'spreadsheet.renameSheetTitle': 'Double-click to rename',
  'spreadsheet.deleteSheetAria': 'Delete sheet {name}',
  'spreadsheet.deleteSheetTitle': 'Delete sheet {name}',
  'spreadsheet.addSheetAria': 'Add sheet',
  'spreadsheet.addSheetTitle': 'Add sheet',
  // A11Y pass 3 — accessible name for the canvas-rendered grid (glide-data-grid
  // has no DOM cells of its own for a screen reader to name) and the live
  // region announcing the currently-selected cell; shared by SpreadsheetViewer
  // and CsvViewer via useGridCellAnnouncement.ts.
  'spreadsheet.gridAria': '{sheet} sheet grid, {dimensions}',
  'csv.gridAria': 'Data grid, {dimensions}',
  'grid.cellAnnouncement': '{ref}: {value}',
  'grid.cellAnnouncementEmpty': '{ref}, blank',

  // SpreadsheetEditToolbar.tsx
  'spreadsheetToolbar.undoAria': 'Undo',
  'spreadsheetToolbar.undoTitle': 'Undo (Ctrl+Z)',
  'spreadsheetToolbar.redoAria': 'Redo',
  'spreadsheetToolbar.redoTitle': 'Redo (Ctrl+Y)',
  'spreadsheetToolbar.insertRowAboveAria': 'Insert row above',
  'spreadsheetToolbar.insertRowAboveTitle': 'Insert row above the selected cell',
  'spreadsheetToolbar.deleteRowAria': 'Delete row',
  'spreadsheetToolbar.deleteRowTitle': 'Delete the selected row',
  'spreadsheetToolbar.insertColumnLeftAria': 'Insert column left',
  'spreadsheetToolbar.insertColumnLeftTitle': 'Insert column left of the selected cell',
  'spreadsheetToolbar.deleteColumnAria': 'Delete column',
  'spreadsheetToolbar.deleteColumnTitle': 'Delete the selected column',
  'spreadsheetToolbar.pasteAria': 'Paste',
  'spreadsheetToolbar.pasteTitle': 'Paste from clipboard',
  'spreadsheetToolbar.saveAria': 'Save',
  'spreadsheetToolbar.saveTitle': 'Save (Ctrl+S)',
  'spreadsheetToolbar.saveAsFormatAria': 'Save As format',
  'spreadsheetToolbar.saveAsAria': 'Save As',
  'spreadsheetToolbar.saveAsLabel': 'Save As…',

  // electron/lib/atomicWrite.cjs (classifyWriteError) — translated here, not
  // hard-coded in main; main only ever returns the stable `errorCode` below.
  'errors.write.permissionDenied': "Permission denied — you don't have access to save to this location.",
  'errors.write.diskFull': 'Not enough disk space to save this file.',
  'errors.write.isDirectory': 'That location is a folder, not a file — choose a different name.',
  'errors.write.destinationMissing': 'The destination folder no longer exists — choose a different location.',
  'errors.write.readOnly': 'That location is read-only — choose a different location.',
  'errors.write.nameTooLong': 'That file name or path is too long — choose a shorter one.',
  'errors.write.fileLocked': 'This file appears to be open in another program — close it there first and try again.',
  'errors.write.unknownNewDocument': 'Could not create the new document.',
  'errors.write.saveFailed': 'Failed to save the file. Please try again.',

  // src/hooks/useFileHandler.ts
  'errors.browserModeUnavailable': 'This feature requires the Atlas desktop app — file access is unavailable in a plain browser tab.',
  'errors.fileNotFound': 'This file could not be found — it may have been moved, renamed, or deleted.',

  // src/utils/friendlyLibraryError.ts
  'errors.library.corruptZip': 'the file could not be read as a valid Office document (it may be corrupted or not a real Office file)',
  'errors.library.unsupportedFormat': 'the file is not in a format this app recognizes',
  'errors.library.tooLarge': 'the document is too large to process',
  'errors.library.corruptPdf': 'the PDF file appears to be corrupted or malformed',
  'errors.library.passwordProtected': 'the file is password-protected and cannot be opened',

  // src/docx/editor/toolbar/Toolbar.tsx
  'docx.toolbar.rootAria': 'Document formatting toolbar',
  'docx.toolbar.tabsAria': 'Toolbar tabs',
  'docx.toolbar.tab.home': 'Home',
  'docx.toolbar.tab.insert': 'Insert',
  'docx.toolbar.tab.layout': 'Layout',
  'docx.toolbar.tab.review': 'Review',
  'docx.toolbar.notYetSupported': 'Not yet supported',
  'docx.toolbar.undo': 'Undo',
  'docx.toolbar.redo': 'Redo',
  'docx.toolbar.fontSizeAria': 'Font size',
  'docx.toolbar.sizePlaceholder': 'Size',
  'docx.toolbar.bold': 'Bold',
  'docx.toolbar.italic': 'Italic',
  'docx.toolbar.underline': 'Underline',
  'docx.toolbar.strikethrough': 'Strikethrough',
  'docx.toolbar.subscript': 'Subscript',
  'docx.toolbar.superscript': 'Superscript',
  'docx.toolbar.fontColor': 'Font Color',
  'docx.toolbar.highlightColor': 'Highlight Color',
  'docx.toolbar.colorSwatchAria': 'Color {hex}',
  'docx.toolbar.colorGridAria': 'Color swatches',
  'docx.toolbar.customColorLabel': 'Custom color (hex)',
  'docx.toolbar.ok': 'Ok',
  'docx.toolbar.alignLeft': 'Align left',
  'docx.toolbar.alignCenter': 'Align center',
  'docx.toolbar.alignRight': 'Align right',
  'docx.toolbar.justify': 'Justify',
  'docx.toolbar.bulletList': 'Bullet List',
  'docx.toolbar.numberedList': 'Numbered List',
  'docx.toolbar.decreaseIndent': 'Decrease Indent',
  'docx.toolbar.increaseIndent': 'Increase Indent',
  'docx.toolbar.findAndReplace': 'Find and Replace',
  'docx.toolbar.pageBreak': 'Page Break',
  'docx.toolbar.insertTable': 'Insert Table',
  'docx.toolbar.insertTableDialogAria': 'Insert table',
  'docx.toolbar.tableSizeAria': 'Table size',
  'docx.toolbar.tableCellAria': '{rows} by {cols} table',
  'docx.toolbar.tableSizeLabel': '{rows}x{cols} Table',
  'docx.toolbar.image': 'Image',
  'docx.toolbar.hyperlink': 'Hyperlink',
  'docx.toolbar.header': 'Header',
  'docx.toolbar.footer': 'Footer',
  'docx.toolbar.comment': 'Comment',
  'docx.toolbar.margins': 'Margins',
  'docx.toolbar.orientation': 'Orientation',
  'docx.toolbar.pageSize': 'Size',
  'docx.toolbar.columns': 'Columns',
  'docx.toolbar.toggleSpellCheck': 'Toggle spell check',
  'docx.toolbar.toggleTrackChanges': 'Toggle track changes',
  'docx.toolbar.accept': 'Accept',
  'docx.toolbar.reject': 'Reject',
  'docx.toolbar.acceptAll': 'Accept All',
  'docx.toolbar.rejectAll': 'Reject All',
  'docx.toolbar.comments': 'Comments',
  'docx.toolbar.editTable': 'Edit Table',
  'docx.toolbar.editTablePlaceCursor': 'Place the cursor inside a table to edit it',
  'docx.toolbar.editTableDialogAria': 'Edit table',
  'docx.toolbar.tableProperties': 'Table Properties…',

  // src/docx/editor/toolbar/FontPicker.tsx
  'docx.fontPicker.inputAria': 'Font',
  'docx.fontPicker.placeholder': 'Font',
  'docx.fontPicker.listAria': 'Fonts',

  // src/docx/editor/toolbar/tableEditActions.ts (rendered by TableEditMenuItems.tsx)
  'docx.tableEdit.insertRowAbove': 'Insert Row Above',
  'docx.tableEdit.insertRowBelow': 'Insert Row Below',
  'docx.tableEdit.insertColumnLeft': 'Insert Column Left',
  'docx.tableEdit.insertColumnRight': 'Insert Column Right',
  'docx.tableEdit.mergeRight': 'Merge Right',
  'docx.tableEdit.splitCell': 'Split Cell',
  'docx.tableEdit.deleteRow': 'Delete Row',
  'docx.tableEdit.deleteColumn': 'Delete Column',
  'docx.tableEdit.deleteTable': 'Delete Table',

  // src/docx/editor/toolbar/TablePropertiesDialog.tsx
  'docx.tableProperties.ariaLabel': 'Table properties',
  'docx.tableProperties.widthLabel': 'Width (in)',
  'docx.tableProperties.widthAria': 'Table width in inches',
  'docx.tableProperties.alignmentLegend': 'Table alignment',
  'docx.tableProperties.alignLeft': 'Left',
  'docx.tableProperties.alignCenter': 'Center',
  'docx.tableProperties.alignRight': 'Right',
  'docx.tableProperties.showBorders': 'Show borders',
  'docx.tableProperties.cancel': 'Cancel',
  'docx.tableProperties.apply': 'Apply',

  // src/docx/editor/FindReplace.tsx
  'docx.findReplace.dialogAria': 'Find and Replace',
  'docx.findReplace.findPlaceholder': 'Find…',
  'docx.findReplace.findAria': 'Find',
  'docx.findReplace.close': 'Close',
  'docx.findReplace.replacePlaceholder': 'Replace…',
  'docx.findReplace.replaceFieldAria': 'Replace',
  'docx.findReplace.caseSensitive': 'Case sensitive',
  'docx.findReplace.wholeWord': 'Whole word',
  'docx.findReplace.regex': 'Regex',
  'docx.findReplace.findPreviousAria': 'Find previous',
  'docx.findReplace.prev': 'Prev',
  'docx.findReplace.findNextAria': 'Find next',
  'docx.findReplace.next': 'Next',
  'docx.findReplace.replace': 'Replace',
  'docx.findReplace.replaceAllAria': 'Replace all',
  'docx.findReplace.replaceAllLabel': 'Replace All',
  'docx.findReplace.noMatches': 'No matches',
  'docx.findReplace.matchCount': { one: '{count} match', other: '{count} matches' },

  // src/docx/editor/CommentsPane.tsx
  'docx.comments.aria': 'Comments',
  'docx.comments.addComment': 'Add comment',
  'docx.comments.count': { one: '{count} comment', other: '{count} comments' },
  'docx.comments.empty': 'No comments yet.',
  'docx.comments.unanchoredThread': 'Unanchored thread',
  'docx.comments.anchorParagraph': 'Anchor paragraph {n}',
  'docx.comments.emptyCommentBody': '(empty comment)',
  'docx.comments.reply': 'Reply',
  'docx.comments.resolve': 'Resolve',
  'docx.comments.delete': 'Delete',
  'docx.comments.unknownAuthor': 'Unknown author',

  // src/docx/editor/SpellCheckMenu.tsx
  'docx.spellCheck.suggestionsAria': 'Spell check suggestions for "{word}"',
  'docx.spellCheck.noSuggestions': 'No suggestions',
  'docx.spellCheck.addToDictionary': 'Add to dictionary',

  // src/viewers/DocxViewer.tsx — header/footer panel
  'docx.headerFooter.title': 'Header and footer',
  'docx.headerFooter.editTitle': 'Edit header and footer',
  'docx.headerFooter.close': 'Close header and footer',
  'docx.headerFooter.header': 'Header',
  'docx.headerFooter.footer': 'Footer',
  'docx.headerFooter.partWithType': '{part} ({type} page)',
  'docx.headerFooter.pageTypeFirst': 'first',
  'docx.headerFooter.pageTypeEven': 'even',
  'docx.headerFooter.lineNumberPrefix': 'Line {n}: ',
  'docx.headerFooter.lineLabel': '{base} line {n}',
  'docx.headerFooter.editableTextAria': '{label} — editable text',
  'docx.headerFooter.removeLine': 'Remove this line',
  'docx.headerFooter.removeLabel': 'Remove {label}',
  'docx.headerFooter.notPlainText': 'Not plain text — left exactly as it is.',
  'docx.headerFooter.nonTextParts': 'Non-text parts (images, fields, links, …) are left exactly as they are.',
  'docx.headerFooter.addLine': '+ Add line',
  'docx.headerFooter.placeholderPageNumber': 'Page number',
  'docx.headerFooter.placeholderTotalPages': 'Total pages',
  'docx.headerFooter.placeholderDate': 'Date',
  'docx.headerFooter.placeholderField': 'Field',
  'docx.headerFooter.placeholderComment': 'Comment',
  'docx.headerFooter.placeholderFootnote': 'Footnote',
  'docx.headerFooter.placeholderEndnote': 'Endnote',
  'docx.headerFooter.placeholderOtherContent': 'Other content',
  'docx.headerFooter.placeholderTable': 'Table',
  'docx.headerFooter.placeholderUnknownContent': 'Unknown content',
  'docx.headerFooter.placeholderEmptyParagraph': 'Empty paragraph',
  'docx.headerFooter.placeholderTrackedChange': 'Tracked change',
  'docx.headerFooter.placeholderLink': 'Link',
  'docx.headerFooter.placeholderLinkPrefix': 'Link',

  // src/viewers/DocxViewer.tsx — editor shell (toolbar trailing, zoom, save, errors)
  'docx.viewer.pagesCount': 'Pages: {count}',
  'docx.viewer.zoomGroupAria': 'Zoom',
  'docx.viewer.zoomOut': 'Zoom out',
  'docx.viewer.zoomReset': 'Reset zoom to 100%',
  'docx.viewer.zoomResetTitle': 'Reset zoom',
  'docx.viewer.zoomIn': 'Zoom in',
  'docx.viewer.updateFields': 'Update fields',
  'docx.viewer.updateFieldsTitle': 'Update fields — recalculate DATE/TIME/AUTHOR/TITLE/REF/PAGEREF/SEQ/PAGE/NUMPAGES',
  'docx.viewer.updateToc': 'Update table of contents',
  'docx.viewer.printDocument': 'Print document',
  'docx.viewer.printTitle': 'Print (Ctrl+P)',
  'docx.viewer.saveAs': 'Save As',
  'docx.viewer.saveAsTitle': 'Save as…',
  'docx.viewer.save': 'Save',
  'docx.viewer.saveTitle': 'Save (Ctrl+S)',
  'docx.viewer.editorAria': 'Document editor',
  'docx.viewer.tableEditingAria': 'Table editing',
  'docx.viewer.dismissError': 'Dismiss error',
  'docx.viewer.dismissMessage': 'Dismiss message',
  'docx.viewer.layoutFailedTitle': 'Failed to lay out document',
  'docx.viewer.layingOut': 'Laying out document…',
  'docx.viewer.blocksProgress': '{completed} / {total} blocks',
  'docx.viewer.saveAsWordFilter': 'Word Documents',
  'docx.viewer.unexpectedTextFile': 'Unexpected text file routed to DocxViewer.',
  'docx.viewer.imagePickerUnavailable': 'Image picker is unavailable in this environment.',
  'docx.viewer.placeCursorBeforeImage': 'Place the cursor in the document before inserting an image.',
  'docx.viewer.selectBeforeHyperlink': 'Select text or place the cursor before inserting a hyperlink.',
  'docx.viewer.enterUrlPrompt': 'Enter a URL',
  'docx.viewer.selectParagraphBeforeList': 'Select a paragraph before toggling a list.',
  'docx.viewer.selectBeforeComment': 'Select text before adding a comment.',
  'docx.viewer.addCommentPrompt': 'Add comment',
  'docx.viewer.replyPrompt': 'Reply',
  'docx.viewer.saveCancelled': 'Save was cancelled or unavailable.',
  'docx.viewer.noFieldsToUpdate': 'No fields needed updating.',
  'docx.viewer.fieldsUpdated': { one: 'Updated {count} field.', other: 'Updated {count} fields.' },
  'docx.viewer.noTocFound': 'No table of contents found to update.',
  'docx.viewer.tocUpdated': 'Table of contents updated.',
  // Round-trip fidelity audit, DXS round 2 follow-up — plain-language,
  // non-technical copy for `detectLossySaveWarnings`' findings (see
  // `docx/fidelity/lossySaveWarnings.ts` and `buildFidelitySaveWarningMessage`
  // in DocxViewer.tsx). Deliberately names the feature the user recognizes,
  // never the underlying XML element.
  'docx.viewer.fidelityWarningIntro': "This document was saved, but Atlas couldn't fully preserve everything in it:",
  'docx.viewer.fidelityWarningContentControls': 'Its content controls (such as dropdowns or date fields) kept their text, but not the controls themselves.',
  'docx.viewer.fidelityWarningShapeFallback': "A shape or text box's fallback drawing was not preserved.",
  'docx.viewer.fidelityWarningOther': 'Some other formatting or content was not fully preserved.',

  // src/docx/render/PageView.tsx
  'docx.render.resizeColumn': 'Resize column {n}',

  // src/docx/editor/friendlyDocxError.ts
  'docx.error.notOpenedDetail': 'This file was not opened: {detail}',
  'docx.error.notSavedDetail': 'This file was not saved: {detail}',
  'docx.error.invalidWordFile': "This file doesn't appear to be a valid Word document (.docx). It may be corrupted, password-protected, or a different file type entirely. ({detail})",
  'docx.error.saveCorruptedData': 'The document could not be saved because its file data appears corrupted. Try again, or save a copy under a new name. ({detail})',
  'docx.error.openXmlCorrupted': "This document's internal structure appears to be corrupted and could not be read. It may have been damaged by another application. ({detail})",
  'docx.error.saveXmlInvalid': 'The document could not be saved because Atlas generated XML that failed its own validity check. Your changes were not written to disk — please try again or report this issue. ({detail})',
  'docx.error.saveVerifyFailed': 'Atlas could not verify the saved file was valid, so nothing was written to disk. Please try again or report this issue. ({detail})',
  'docx.error.openParseFailed': 'This Word document could not be opened. ({detail})',
  'docx.error.saveParseFailed': 'This Word document could not be saved. ({detail})',
  'docx.error.openGenericFailed': "Couldn't open this Word document. ({detail})",
  'docx.error.saveGenericFailed': "Couldn't save this Word document. Your changes have not been written to disk. ({detail})",

  // src/viewers/shared/SlideDeck.tsx
  'slides.deck.slideTitle': 'Slide {n}',
  'slides.deck.thumbnailRailAria': 'Slide thumbnails',
  'slides.deck.zoomOut': 'Zoom out',
  'slides.deck.zoomIn': 'Zoom in',
  'slides.deck.zoomReset': 'Reset zoom',
  'slides.deck.speakerNotes': 'Speaker notes',
  'slides.deck.presentFullscreen': 'Present (fullscreen)',
  'slides.deck.presenterView': 'Presenter view',
  'slides.deck.noSlidesAvailable': 'No slides available.',
  'slides.deck.exitPresentation': 'Exit presentation (Esc)',
  'slides.deck.closeNotes': 'Close notes',
  'slides.deck.notesPlaceholder': 'Click to add notes',

  // src/viewers/shared/SlideEditToolbar.tsx
  'slides.editToolbar.editPresentationAria': 'Edit presentation',
  'slides.editToolbar.undo': 'Undo',
  'slides.editToolbar.redo': 'Redo',
  'slides.editToolbar.newSlide': 'New slide',
  'slides.editToolbar.duplicateSlide': 'Duplicate slide',
  'slides.editToolbar.deleteSlide': 'Delete slide',
  'slides.editToolbar.moveSlideUp': 'Move slide up',
  'slides.editToolbar.moveSlideDown': 'Move slide down',
  'slides.editToolbar.insertTextBox': 'Insert text box',
  'slides.editToolbar.saveAs': 'Save As',
  'slides.editToolbar.save': 'Save',
  'slides.editToolbar.saveTitle': 'Save (Ctrl+S)',

  // src/viewers/shared/PresenterView.tsx
  'slides.presenter.ariaLabel': 'Presenter view',
  'slides.presenter.next': 'Next',
  'slides.presenter.endOfPresentation': 'End of presentation',
  'slides.presenter.notes': 'Notes',
  'slides.presenter.noNotesForSlide': 'No notes for this slide.',
  'slides.presenter.elapsedTimeAria': 'Elapsed time',
  'slides.presenter.previousSlideAria': 'Previous slide',
  'slides.presenter.slideCounter': 'Slide {current} / {total}',
  'slides.presenter.nextSlideAria': 'Next slide',
  'slides.presenter.exitPresenterViewAria': 'Exit presenter view',

  // src/viewers/shared/SlideEditCanvas.tsx
  'slides.editCanvas.editTextAria': 'Edit slide text',
  'slides.editCanvas.surfaceAria': 'Slide editing surface',

  // src/viewers/pdf/PdfFindBar.tsx
  'pdf.findBar.searching': 'Searching…',
  'pdf.findBar.placeholder': 'Find in document…',
  'pdf.findBar.inputAria': 'Find in PDF',
  'pdf.findBar.indexingProgress': 'Indexing {indexed}/{total}…',
  'pdf.findBar.previousMatchTitle': 'Previous match (Shift+Enter)',
  'pdf.findBar.nextMatchTitle': 'Next match (Enter)',
  'pdf.findBar.closeAria': 'Close find',

  // src/viewers/pdf/PdfPasswordDialog.tsx
  'pdf.password.dialogAria': 'Password required',
  'pdf.password.title': 'This PDF is password protected',
  'pdf.password.placeholder': 'Enter password',
  'pdf.password.inputAria': 'PDF password',
  'pdf.password.incorrectError': 'Incorrect password. Try again.',
  'pdf.password.unlock': 'Unlock',

  // src/viewers/pdf/PdfThumbnailRail.tsx
  'pdf.thumbnails.goToPage': 'Go to page {n}',
  'pdf.thumbnails.railAria': 'Page thumbnails',

  // src/formats/legacyOffice.ts (consumed by UnknownViewer.tsx, and reused
  // by LegacyDocViewer.tsx/LegacyPptViewer.tsx for the same format labels)
  'legacyOffice.message': "This is a {label}. Atlas doesn't support the legacy binary Office formats yet — re-save it as {modernExtension} in Word, Excel, PowerPoint, or a compatible app, then reopen it here.",
  'legacyOffice.label.doc': 'Word 97-2003 document (.doc)',
  'legacyOffice.label.xls': 'Excel 97-2003 workbook (.xls)',
  'legacyOffice.label.ppt': 'PowerPoint 97-2003 presentation (.ppt)',
  'legacyOffice.label.unknown': 'legacy Microsoft Office document',
  'legacyOffice.modernExtensionUnknown': 'its modern XML format (.docx / .xlsx / .pptx)',

  // src/viewers/LegacyDocViewer.tsx
  'legacy.doc.unexpectedTextFile': 'LegacyDocViewer received a text file; expected binary.',
  'legacy.doc.readError': "Couldn't read this Word 97-2003 document: {error}",
  'legacy.doc.reading': 'Reading legacy Word document…',
  'legacy.doc.emptyDocument': 'This document has no readable text.',

  // src/viewers/LegacyPptViewer.tsx
  'legacy.ppt.noSlidesFound': 'No slides found in this presentation.',
  'legacy.ppt.unexpectedTextFile': 'LegacyPptViewer received a text file; expected binary.',
  'legacy.ppt.readError': "Couldn't read this PowerPoint 97-2003 presentation: {error}",
  'legacy.ppt.reading': 'Reading legacy PowerPoint presentation…',

  // src/viewers/UnknownViewer.tsx
  'unknown.viewer.revealRequiresDesktop': 'Reveal in folder requires the Atlas desktop app.',
  'unknown.viewer.revealFailed': 'Could not reveal this file — it may have moved or been deleted.',
  'unknown.viewer.back': 'Back',
  'unknown.viewer.truncatedNotice': 'Showing the first {size} only',
  'unknown.viewer.genericExplanation': "Atlas doesn't recognize this file's format, so there's nothing to preview here.",
  'unknown.viewer.openAsText': 'Open as text',
  'unknown.viewer.revealInFolder': 'Reveal in folder',

  // src/hooks/useFileHandler.ts
  'fileHandler.formatLabel.docx': 'Word (.docx)',
  'fileHandler.formatLabel.xlsx': 'Excel (.xlsx)',
  'fileHandler.formatLabel.pptx': 'PowerPoint (.pptx)',
  'fileHandler.formatLabel.pdf': 'PDF (.pdf)',
  'fileHandler.formatLabel.odt': 'OpenDocument Text (.odt)',
  'fileHandler.formatLabel.ods': 'OpenDocument Spreadsheet (.ods)',
  'fileHandler.formatLabel.odp': 'OpenDocument Presentation (.odp)',
  'fileHandler.formatLabel.rtf': 'Rich Text (.rtf)',
  'fileHandler.formatLabel.doc': 'Word 97-2003 (.doc)',
  'fileHandler.formatLabel.ppt': 'PowerPoint 97-2003 (.ppt)',
  'fileHandler.extensionMismatchConfirm': '"{name}" doesn\'t look like a valid {extFormat} file — its contents look like {magicFormat} instead. Open it anyway?',
} as const satisfies MessageCatalogue;

export type MessageKey = keyof typeof messages;

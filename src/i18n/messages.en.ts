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

  // Sidebar.tsx
  'sidebar.title': 'Table of Contents',
  'sidebar.empty': 'No outline available',

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
} as const satisfies MessageCatalogue;

export type MessageKey = keyof typeof messages;

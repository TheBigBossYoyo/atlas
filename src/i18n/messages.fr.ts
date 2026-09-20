import type { MessageValue } from './types';
import type { MessageKey } from './messages.en';

/**
 * French catalogue. Must define exactly the same key set as
 * `messages.en.ts` (enforced by `src/i18n/__tests__/catalogueParity.test.ts`).
 *
 * Typography: French usage puts a narrow no-break space before a colon,
 * semicolon, exclamation or question mark, and a regular no-break space
 * against French guillemets. Written as JS unicode escapes in the string
 * values below (codepoint 202F for the narrow no-break space, codepoint 00A0
 * for the regular one) rather than as literal invisible whitespace, so a
 * reviewer can see it in the source. This comment spells the codepoints out
 * in words rather than writing the escape sequence itself, since the escape
 * sequence would be decoded to the literal (lint-flagged) character the
 * moment this file is written.
 */
export const messages = {
  'common.unsavedChanges': 'Modifications non enregistrées',
  'common.dismiss': 'Ignorer',

  // WelcomeScreen.tsx
  'welcome.subtitle': 'Une visionneuse de documents universelle pour Word, Excel, PowerPoint, PDF, le code, Markdown et plus encore.',
  'welcome.openFile': 'Ouvrir un fichier',
  'welcome.loadSample': 'Charger un document d’exemple',
  'welcome.dragDropHint': 'ou glissez-déposez un document n’importe où',
  'welcome.recentFiles': 'Fichiers récents',
  'welcome.removeRecentAria': 'Retirer {name} des fichiers récents',
  'welcome.removeRecentTitle': 'Retirer des fichiers récents',
  'welcome.relative.justNow': 'à l’instant',
  'welcome.relative.minutesAgo': { one: 'il y a {count} min', other: 'il y a {count} min' },
  'welcome.relative.hoursAgo': { one: 'il y a {count} h', other: 'il y a {count} h' },
  'welcome.relative.daysAgo': { one: 'il y a {count} j', other: 'il y a {count} j' },
  'welcome.feature.office': 'Word, Excel, PowerPoint, PDF',
  'welcome.feature.markdown': 'Markdown et coloration syntaxique du code',
  'welcome.feature.math': 'Formules mathématiques et diagrammes Mermaid',
  'welcome.feature.themes': '5 thèmes magnifiques',
  'welcome.feature.search': 'Recherche dans le document',
  'welcome.feature.export': 'Export PDF, plus DOCX/HTML pour Markdown',

  // DropZone.tsx
  'dropzone.title': 'Déposez votre document',
  'dropzone.detail': 'Word, Excel, PowerPoint, PDF, OpenDocument, Markdown, CSV, RTF et fichiers de code',

  // Sidebar.tsx
  'sidebar.title': 'Table des matières',
  'sidebar.empty': 'Aucun plan disponible',

  // Toolbar.tsx
  'toolbar.showSidebar': 'Afficher la barre latérale',
  'toolbar.hideSidebar': 'Masquer la barre latérale',
  'toolbar.toggleSidebarAria': 'Afficher/masquer la barre latérale',
  'toolbar.open': 'Ouvrir',
  'toolbar.openTitle': 'Ouvrir un fichier (Ctrl+O)',
  'toolbar.save': 'Enregistrer',
  'toolbar.saveTitle': 'Enregistrer (Ctrl+S)',
  'toolbar.closeFileAria': 'Fermer le fichier',
  'toolbar.closeFileTitle': 'Fermer le fichier (Ctrl+W)',
  'toolbar.viewMode.preview': 'Aperçu',
  'toolbar.viewMode.split': 'Partagé',
  'toolbar.viewMode.editor': 'Éditeur',
  'toolbar.fontSizeGroupAria': 'Taille de police',
  'toolbar.decreaseFontAria': 'Réduire la taille de police',
  'toolbar.decreaseFontTitle': 'Réduire la taille de police (Ctrl+-)',
  'toolbar.resetFontTitle': 'Réinitialiser la taille de police (Ctrl+0)',
  'toolbar.increaseFontAria': 'Augmenter la taille de police',
  'toolbar.increaseFontTitle': 'Augmenter la taille de police (Ctrl+=)',
  'toolbar.searchAria': 'Rechercher',
  'toolbar.searchTitle': 'Rechercher (Ctrl+F)',
  'toolbar.shortcutsAria': 'Raccourcis clavier',
  'toolbar.shortcutsTitle': 'Raccourcis clavier (Ctrl+/)',

  // NewDocumentMenu.tsx
  'newDocument.trigger': 'Nouveau',
  'newDocument.triggerTitle': 'Nouveau document (Ctrl+N)',
  'newDocument.menuLabel': 'Nouveau document',
  'newDocument.markdown': 'Markdown (.md)',
  'newDocument.docx': 'Document Word (.docx)',
  'newDocument.xlsx': 'Classeur Excel (.xlsx)',
  'newDocument.ods': 'Classeur OpenDocument (.ods)',
  'newDocument.pptx': 'Présentation PowerPoint (.pptx)',
  'newDocument.odp': 'Présentation OpenDocument (.odp)',

  // ExportMenu.tsx
  'exportMenu.triggerAria': 'Exporter',
  'exportMenu.triggerTitle': 'Exporter',
  'exportMenu.menuLabel': 'Options d’export',
  'exportMenu.html': 'HTML',
  'exportMenu.pdf': 'PDF',
  'exportMenu.docx': 'DOCX',
  'exportMenu.markdown': 'Markdown',
  'exportMenu.saveCopy': 'Enregistrer une copie',
  'exportMenu.exportToPdf': 'Exporter en PDF',
  'exportMenu.exportToCsv': 'Exporter en CSV',
  'exportMenu.exportToHtml': 'Exporter en HTML',

  // ThemeMenu.tsx
  'theme.triggerAria': 'Thème',
  'theme.triggerTitle': 'Thème',
  'theme.menuLabel': 'Options de thème',

  // LanguageMenu.tsx
  'language.triggerAria': 'Langue',
  'language.triggerTitle': 'Langue',
  'language.menuLabel': 'Options de langue',
  'language.english': 'Anglais',
  'language.french': 'Français',
  'language.system': 'Système',

  // TabBar.tsx
  'tabBar.label': 'Documents ouverts',
  'tabBar.closeAria': 'Fermer {name}',
  'tabBar.closeTitle': 'Fermer {name}',

  // ShortcutsModal.tsx
  'shortcuts.title': 'Raccourcis clavier',
  'shortcuts.closeAria': 'Fermer',
  'shortcuts.sectionFile': 'Fichier',
  'shortcuts.newDocumentMenu': 'Menu Nouveau document',
  'shortcuts.openFile': 'Ouvrir un fichier',
  'shortcuts.save': 'Enregistrer',
  'shortcuts.saveAs': 'Enregistrer sous',
  'shortcuts.closeFile': 'Fermer le fichier',
  'shortcuts.reopenClosed': 'Rouvrir le dernier document fermé',
  'shortcuts.printExport': 'Imprimer / exporter',
  'shortcuts.exportMenu': 'Menu Export',
  'shortcuts.sectionView': 'Affichage',
  'shortcuts.preview': 'Aperçu',
  'shortcuts.split': 'Partagé',
  'shortcuts.editor': 'Éditeur',
  'shortcuts.toggleSidebar': 'Afficher/masquer la barre latérale',
  'shortcuts.cycleTheme': 'Changer de thème',
  'shortcuts.sectionEdit': 'Édition',
  'shortcuts.increaseFont': 'Augmenter la police',
  'shortcuts.decreaseFont': 'Réduire la police',
  'shortcuts.resetFont': 'Réinitialiser la police',
  'shortcuts.sectionSearchHelp': 'Recherche et aide',
  'shortcuts.find': 'Rechercher (Markdown, texte, code, RTF, ODT)',
  'shortcuts.nextPrev': 'Suivant/précédent',
  'shortcuts.closeDialog': 'Fermer',
  'shortcuts.toggleDialog': 'Afficher/masquer cette boîte de dialogue',
  'shortcuts.note': 'Lors de l’édition d’un document DOCX, ses propres raccourcis d’édition (gras/italique/souligné, alignement, rechercher/remplacer, interligne, annuler/rétablir, enregistrer, imprimer) sont prioritaires sur les raccourcis ci-dessus.',

  // UnsavedChangesDialog.tsx
  'unsavedDialog.title': 'Modifications non enregistrées',
  'unsavedDialog.body': 'Ce document contient des modifications non enregistrées. Enregistrez-les avant de continuer, ignorez-les ou annulez.',
  'unsavedDialog.cancel': 'Annuler',
  'unsavedDialog.discard': 'Ignorer',
  'unsavedDialog.save': 'Enregistrer',
  'unsavedDialog.saving': 'Enregistrement…',

  // DraftRecoveryBanner.tsx
  'draftRecovery.foundNamed': 'Des modifications non enregistrées de « {fileName} » datant du {savedAt} ont été trouvées.',
  'draftRecovery.foundUnnamed': 'Un brouillon non enregistré datant du {savedAt} a été trouvé.',
  'draftRecovery.discard': 'Ignorer',
  'draftRecovery.restore': 'Restaurer',

  // FileStatusBanner.tsx
  'fileStatus.opening': 'Ouverture du fichier…',
  'fileStatus.dismissErrorAria': 'Ignorer l’erreur',
  'fileStatus.dismissErrorTitle': 'Ignorer',

  // StatusBar.tsx
  'statusBar.untitled': 'Sans titre',
  'statusBar.unsavedChangesTitle': 'Modifications non enregistrées',
  'statusBar.words': { one: '{count} mot', other: '{count} mots' },
  'statusBar.headings': { one: '{count} titre', other: '{count} titres' },
  'statusBar.markdownStats': '{words} · {headings}',
  'statusBar.rows': { one: '{count} ligne', other: '{count} lignes' },
  'statusBar.cols': { one: '{count} col.', other: '{count} col.' },
  'statusBar.spreadsheetStats': 'Feuille : {sheet} · {rows} × {cols}',
  'statusBar.pdfStats': 'Page {page} / {pageCount}',
  'statusBar.slidesStats': 'Diapositive {slide} / {slideCount}',
  'statusBar.lines': { one: '{count} ligne', other: '{count} lignes' },
  'statusBar.codeStats': '{language} · {lines}',
  'statusBar.chars': { one: '{count} caractère', other: '{count} caractères' },
  'statusBar.textStats': '{lines} · {chars}',
  'statusBar.pages': { one: '{count} page', other: '{count} pages' },
  'statusBar.documentStats': '{words} · {pages}',

  // Toast.tsx
  'toast.viewportAria': 'Notifications',
  'toast.dismissAria': 'Ignorer la notification',

  // SearchOverlay.tsx
  'search.placeholder': 'Rechercher dans le document…',
  'search.noResults': 'Aucun résultat',
  'search.resultCount': '{current} sur {total}',
  'search.previousAria': 'Occurrence précédente',
  'search.previousTitle': 'Précédent (Maj+Entrée)',
  'search.nextAria': 'Occurrence suivante',
  'search.nextTitle': 'Suivant (Entrée)',
  'search.closeAria': 'Fermer la recherche',
  'search.closeTitle': 'Fermer (Échap)',

  // EmptyFileNotice.tsx
  'emptyFile.title': 'Ce fichier est vide',
  'emptyFile.detail': '{fileName} n’a pas encore de contenu, et Atlas ne dispose pas de modèle vierge pour ce format. Ouvrez-le dans une autre application pour y ajouter du contenu.',

  // LegacyFormatBanner.tsx
  'legacyFormat.notice': '{formatLabel} — format hérité, affiché en lecture seule sous forme de texte. Réenregistrez-le au format {modernExtension} dans Word, PowerPoint ou une application compatible pour une édition fidèle.',

  // CodeViewer.tsx
  'codeViewer.toolbarAria': 'Éditeur de code',
  'codeViewer.stop': 'Arrêter',
  'codeViewer.run': 'Exécuter',
  'codeViewer.stopTitle': 'Arrêter le programme en cours d’exécution',
  'codeViewer.runTitle': 'Enregistrer et exécuter ce fichier',
  'codeViewer.findReplaceAria': 'Rechercher et remplacer',
  'codeViewer.findReplaceTitle': 'Rechercher et remplacer (Ctrl+F)',
  'codeViewer.goToLineAria': 'Atteindre la ligne',
  'codeViewer.goToLineTitle': 'Atteindre la ligne (Ctrl+G)',
  'codeViewer.wordWrapAria': 'Retour à la ligne automatique',
  'codeViewer.saveAria': 'Enregistrer',
  'codeViewer.saveTitle': 'Enregistrer (Ctrl+S)',
  'codeViewer.saveCancelled': 'L’enregistrement a été annulé ou n’est pas disponible.',
  'codeViewer.outputRegionAria': 'Sortie du programme',
  'codeViewer.running': 'Exécution…',
  'codeViewer.output': 'Sortie',
  'codeViewer.clearOutputAria': 'Effacer la sortie',
  'codeViewer.closeOutputAria': 'Fermer la sortie',

  // useCodeRun.ts
  'codeRun.stopped': 'Arrêté.',
  'codeRun.timedOut': 'Arrêté après la limite de 60 secondes.',
  'codeRun.finishedOk': 'Terminé (code de sortie 0).',
  'codeRun.exitedWithCode': 'Terminé avec le code {code}.',
  'codeRun.couldNotStart': 'Le programme n’a pas pu être démarré.',

  // PdfToolbar.tsx
  'pdfToolbar.toggleThumbnails': 'Afficher/masquer les vignettes de page',
  'pdfToolbar.previousPage': 'Page précédente (Haut / Page préc.)',
  'pdfToolbar.nextPage': 'Page suivante (Bas / Page suiv.)',
  'pdfToolbar.pageNumberAria': 'Numéro de page',
  'pdfToolbar.find': 'Rechercher (Ctrl+F)',
  'pdfToolbar.rotate': 'Faire pivoter la page',
  'pdfToolbar.print': 'Imprimer (Ctrl+P)',
  'pdfToolbar.zoomOut': 'Zoom arrière (Ctrl+-)',
  'pdfToolbar.zoomIn': 'Zoom avant (Ctrl++)',
  'pdfToolbar.fitWidth': 'Ajuster à la largeur',
  'pdfToolbar.fitPage': 'Ajuster à la page',

  // SpreadsheetViewer.tsx
  'spreadsheet.format.xlsx': 'Classeur Excel (.xlsx)',
  'spreadsheet.format.xlsm': 'Classeur Excel (.xlsm) avec macros',
  'spreadsheet.format.xlsb': 'Classeur Excel binaire (.xlsb)',
  'spreadsheet.format.xls': 'Classeur Excel 97-2003 (.xls)',
  'spreadsheet.format.ods': 'Classeur OpenDocument (.ods)',
  'spreadsheet.format.fods': 'Classeur OpenDocument plat (.fods)',

  // CsvViewer.tsx
  'csv.format.csv': 'CSV (séparateur virgule)',
  'csv.format.tsv': 'TSV (séparateur tabulation)',
  'spreadsheet.hideHiddenSheets': 'Masquer les feuilles masquées',
  'spreadsheet.showHiddenSheets': 'Afficher les feuilles masquées',
  'spreadsheet.searchRowsPlaceholder': 'Rechercher dans les lignes…',
  'spreadsheet.rowsCount': { one: '{count} ligne', other: '{count} lignes' },
  'spreadsheet.colsCount': { one: '{count} colonne', other: '{count} colonnes' },
  'spreadsheet.dimensions': '{rows} × {cols}',
  'spreadsheet.noSearchResults': 'Aucune ligne ne correspond à votre recherche.',
  'spreadsheet.emptySheet': 'Cette feuille est vide.',
  'spreadsheet.renameSheetAria': 'Renommer la feuille {name}',
  'spreadsheet.renameSheetTitle': 'Double-cliquez pour renommer',
  'spreadsheet.deleteSheetAria': 'Supprimer la feuille {name}',
  'spreadsheet.deleteSheetTitle': 'Supprimer la feuille {name}',
  'spreadsheet.addSheetAria': 'Ajouter une feuille',
  'spreadsheet.addSheetTitle': 'Ajouter une feuille',

  // SpreadsheetEditToolbar.tsx
  'spreadsheetToolbar.undoAria': 'Annuler',
  'spreadsheetToolbar.undoTitle': 'Annuler (Ctrl+Z)',
  'spreadsheetToolbar.redoAria': 'Rétablir',
  'spreadsheetToolbar.redoTitle': 'Rétablir (Ctrl+Y)',
  'spreadsheetToolbar.insertRowAboveAria': 'Insérer une ligne au-dessus',
  'spreadsheetToolbar.insertRowAboveTitle': 'Insérer une ligne au-dessus de la cellule sélectionnée',
  'spreadsheetToolbar.deleteRowAria': 'Supprimer la ligne',
  'spreadsheetToolbar.deleteRowTitle': 'Supprimer la ligne sélectionnée',
  'spreadsheetToolbar.insertColumnLeftAria': 'Insérer une colonne à gauche',
  'spreadsheetToolbar.insertColumnLeftTitle': 'Insérer une colonne à gauche de la cellule sélectionnée',
  'spreadsheetToolbar.deleteColumnAria': 'Supprimer la colonne',
  'spreadsheetToolbar.deleteColumnTitle': 'Supprimer la colonne sélectionnée',
  'spreadsheetToolbar.pasteAria': 'Coller',
  'spreadsheetToolbar.pasteTitle': 'Coller depuis le presse-papiers',
  'spreadsheetToolbar.saveAria': 'Enregistrer',
  'spreadsheetToolbar.saveTitle': 'Enregistrer (Ctrl+S)',
  'spreadsheetToolbar.saveAsFormatAria': 'Format Enregistrer sous',
  'spreadsheetToolbar.saveAsAria': 'Enregistrer sous',
  'spreadsheetToolbar.saveAsLabel': 'Enregistrer sous…',

  // electron/lib/atomicWrite.cjs (classifyWriteError)
  'errors.write.permissionDenied': 'Accès refusé — vous n’avez pas l’autorisation d’enregistrer à cet emplacement.',
  'errors.write.diskFull': 'Espace disque insuffisant pour enregistrer ce fichier.',
  'errors.write.isDirectory': 'Cet emplacement est un dossier, pas un fichier — choisissez un autre nom.',
  'errors.write.destinationMissing': 'Le dossier de destination n’existe plus — choisissez un autre emplacement.',
  'errors.write.readOnly': 'Cet emplacement est en lecture seule — choisissez un autre emplacement.',
  'errors.write.nameTooLong': 'Ce nom de fichier ou ce chemin est trop long — choisissez-en un plus court.',
  'errors.write.fileLocked': 'Ce fichier semble ouvert dans un autre programme — fermez-le puis réessayez.',
  'errors.write.unknownNewDocument': 'Impossible de créer le nouveau document.',
  'errors.write.saveFailed': 'Échec de l’enregistrement du fichier. Veuillez réessayer.',

  // src/hooks/useFileHandler.ts
  'errors.browserModeUnavailable': 'Cette fonctionnalité nécessite l’application de bureau Atlas — l’accès aux fichiers n’est pas disponible dans un simple onglet de navigateur.',
  'errors.fileNotFound': 'Ce fichier est introuvable — il a peut-être été déplacé, renommé ou supprimé.',

  // src/utils/friendlyLibraryError.ts
  'errors.library.corruptZip': 'le fichier n’a pas pu être lu comme un document Office valide (il est peut-être corrompu ou n’est pas un véritable fichier Office)',
  'errors.library.unsupportedFormat': 'le fichier n’est pas dans un format reconnu par cette application',
  'errors.library.tooLarge': 'le document est trop volumineux pour être traité',
  'errors.library.corruptPdf': 'le fichier PDF semble corrompu ou mal formé',
  'errors.library.passwordProtected': 'le fichier est protégé par mot de passe et ne peut pas être ouvert',
} as const satisfies Record<MessageKey, MessageValue>;

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

  // ViewerErrorBoundary.tsx
  'viewerError.title': 'La visionneuse a planté',
  'viewerError.tryAgain': 'Réessayer',

  // ViewerLoading.tsx
  'viewerLoading.loading': 'Chargement de {format}…',
  'viewerLoading.documentFallback': 'document',

  // Sidebar.tsx
  'sidebar.title': 'Table des matières',
  'sidebar.empty': 'Aucun plan disponible',

  // RawEditor.tsx
  'rawEditor.title': 'Source Markdown',
  'rawEditor.placeholder': 'Saisissez ou collez du Markdown ici…',

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
  'toolbar.viewMode.split': 'Fractionné',
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
  'shortcuts.nextPrevTab': 'Onglet suivant / précédent',
  'shortcuts.reopenClosed': 'Rouvrir le dernier document fermé',
  'shortcuts.printExport': 'Imprimer / exporter',
  'shortcuts.exportMenu': 'Menu Export',
  'shortcuts.sectionView': 'Affichage',
  'shortcuts.preview': 'Aperçu',
  'shortcuts.split': 'Fractionné',
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
  // LargeMarkdownNotice.tsx
  'largeMarkdown.title': "Ce document est trop volumineux pour un aperçu rapide",
  'largeMarkdown.detail': "Afficher ces {size} Mo de Markdown peut prendre un certain temps. L’éditeur est ouvert et pleinement utilisable ; vous pouvez tout de même demander l’aperçu.",
  'largeMarkdown.renderAnyway': "Afficher l’aperçu quand même",

  // MarkdownRenderer.tsx / useMarkdownHastTree.ts
  'markdownWorker.parsing': 'Affichage du document volumineux…',
  'markdownWorker.error': "Ce document n’a pas pu être affiché : {error}",

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
  'spreadsheet.gridAria': 'Grille de la feuille {sheet}, {dimensions}',
  'spreadsheet.saveFallbackWarning':
    "Ce fichier a été enregistré, mais Atlas n'a pas pu réutiliser sa mise en page d'origine ; certains styles, graphiques, filtres ou tableaux peuvent ne pas avoir été préservés.",
  'csv.gridAria': 'Grille de données, {dimensions}',
  'grid.cellAnnouncement': '{ref} : {value}',
  'grid.cellAnnouncementEmpty': '{ref}, vide',

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

  // src/docx/editor/toolbar/Toolbar.tsx
  'docx.toolbar.rootAria': 'Barre d’outils de mise en forme du document',
  'docx.toolbar.tabsAria': 'Onglets de la barre d’outils',
  'docx.toolbar.tab.home': 'Accueil',
  'docx.toolbar.tab.insert': 'Insertion',
  'docx.toolbar.tab.layout': 'Mise en page',
  'docx.toolbar.tab.review': 'Révision',
  'docx.toolbar.notYetSupported': 'Pas encore pris en charge',
  'docx.toolbar.undo': 'Annuler',
  'docx.toolbar.redo': 'Rétablir',
  'docx.toolbar.fontSizeAria': 'Taille de police',
  'docx.toolbar.sizePlaceholder': 'Taille',
  'docx.toolbar.bold': 'Gras',
  'docx.toolbar.italic': 'Italique',
  'docx.toolbar.underline': 'Souligné',
  'docx.toolbar.strikethrough': 'Barré',
  'docx.toolbar.subscript': 'Indice',
  'docx.toolbar.superscript': 'Exposant',
  'docx.toolbar.fontColor': 'Couleur de police',
  'docx.toolbar.highlightColor': 'Couleur de surbrillance',
  'docx.toolbar.colorSwatchAria': 'Couleur {hex}',
  'docx.toolbar.colorGridAria': 'Nuancier de couleurs',
  'docx.toolbar.customColorLabel': 'Couleur personnalisée (hex)',
  'docx.toolbar.ok': 'OK',
  'docx.toolbar.alignLeft': 'Aligner à gauche',
  'docx.toolbar.alignCenter': 'Centrer',
  'docx.toolbar.alignRight': 'Aligner à droite',
  'docx.toolbar.justify': 'Justifier',
  'docx.toolbar.bulletList': 'Puces',
  'docx.toolbar.numberedList': 'Numérotation',
  'docx.toolbar.decreaseIndent': 'Réduire le retrait',
  'docx.toolbar.increaseIndent': 'Augmenter le retrait',
  'docx.toolbar.findAndReplace': 'Rechercher et remplacer',
  'docx.toolbar.pageBreak': 'Saut de page',
  'docx.toolbar.insertTable': 'Insérer un tableau',
  'docx.toolbar.insertTableDialogAria': 'Insérer un tableau',
  'docx.toolbar.tableSizeAria': 'Taille du tableau',
  'docx.toolbar.tableCellAria': 'Tableau {rows} × {cols}',
  'docx.toolbar.tableSizeLabel': 'Tableau {rows} × {cols}',
  'docx.toolbar.image': 'Image',
  'docx.toolbar.hyperlink': 'Lien hypertexte',
  'docx.toolbar.header': 'En-tête',
  'docx.toolbar.footer': 'Pied de page',
  'docx.toolbar.comment': 'Commentaire',
  'docx.toolbar.margins': 'Marges',
  'docx.toolbar.orientation': 'Orientation',
  'docx.toolbar.pageSize': 'Taille',
  'docx.toolbar.columns': 'Colonnes',
  'docx.toolbar.toggleSpellCheck': 'Activer/désactiver la vérification orthographique',
  'docx.toolbar.toggleTrackChanges': 'Activer/désactiver le suivi des modifications',
  'docx.toolbar.accept': 'Accepter',
  'docx.toolbar.reject': 'Rejeter',
  'docx.toolbar.acceptAll': 'Tout accepter',
  'docx.toolbar.rejectAll': 'Tout rejeter',
  'docx.toolbar.comments': 'Commentaires',
  'docx.toolbar.editTable': 'Modifier le tableau',
  'docx.toolbar.editTablePlaceCursor': 'Placez le curseur dans un tableau pour le modifier',
  'docx.toolbar.editTableDialogAria': 'Modifier le tableau',
  'docx.toolbar.tableProperties': 'Propriétés du tableau…',

  // src/docx/editor/toolbar/FontPicker.tsx
  'docx.fontPicker.inputAria': 'Police',
  'docx.fontPicker.placeholder': 'Police',
  'docx.fontPicker.listAria': 'Polices',

  // src/docx/editor/toolbar/tableEditActions.ts (rendered by TableEditMenuItems.tsx)
  'docx.tableEdit.insertRowAbove': 'Insérer une ligne au-dessus',
  'docx.tableEdit.insertRowBelow': 'Insérer une ligne en dessous',
  'docx.tableEdit.insertColumnLeft': 'Insérer une colonne à gauche',
  'docx.tableEdit.insertColumnRight': 'Insérer une colonne à droite',
  'docx.tableEdit.mergeRight': 'Fusionner à droite',
  'docx.tableEdit.splitCell': 'Fractionner la cellule',
  'docx.tableEdit.deleteRow': 'Supprimer la ligne',
  'docx.tableEdit.deleteColumn': 'Supprimer la colonne',
  'docx.tableEdit.deleteTable': 'Supprimer le tableau',

  // src/docx/editor/toolbar/TablePropertiesDialog.tsx
  'docx.tableProperties.ariaLabel': 'Propriétés du tableau',
  'docx.tableProperties.widthLabel': 'Largeur (po)',
  'docx.tableProperties.widthAria': 'Largeur du tableau en pouces',
  'docx.tableProperties.alignmentLegend': 'Alignement du tableau',
  'docx.tableProperties.alignLeft': 'Gauche',
  'docx.tableProperties.alignCenter': 'Centre',
  'docx.tableProperties.alignRight': 'Droite',
  'docx.tableProperties.showBorders': 'Afficher les bordures',
  'docx.tableProperties.cancel': 'Annuler',
  'docx.tableProperties.apply': 'Appliquer',

  // src/docx/editor/FindReplace.tsx
  'docx.findReplace.dialogAria': 'Rechercher et remplacer',
  'docx.findReplace.findPlaceholder': 'Rechercher…',
  'docx.findReplace.findAria': 'Rechercher',
  'docx.findReplace.close': 'Fermer',
  'docx.findReplace.replacePlaceholder': 'Remplacer…',
  'docx.findReplace.replaceFieldAria': 'Remplacer',
  'docx.findReplace.caseSensitive': 'Respecter la casse',
  'docx.findReplace.wholeWord': 'Mot entier',
  'docx.findReplace.regex': 'Expression régulière',
  'docx.findReplace.findPreviousAria': 'Occurrence précédente',
  'docx.findReplace.prev': 'Préc.',
  'docx.findReplace.findNextAria': 'Occurrence suivante',
  'docx.findReplace.next': 'Suiv.',
  'docx.findReplace.replace': 'Remplacer',
  'docx.findReplace.replaceAllAria': 'Tout remplacer',
  'docx.findReplace.replaceAllLabel': 'Tout remplacer',
  'docx.findReplace.noMatches': 'Aucun résultat',
  'docx.findReplace.matchCount': { one: '{count} correspondance', other: '{count} correspondances' },

  // src/docx/editor/CommentsPane.tsx
  'docx.comments.aria': 'Commentaires',
  'docx.comments.addComment': 'Ajouter un commentaire',
  'docx.comments.count': { one: '{count} commentaire', other: '{count} commentaires' },
  'docx.comments.empty': 'Aucun commentaire pour l’instant.',
  'docx.comments.unanchoredThread': 'Fil sans ancrage',
  'docx.comments.anchorParagraph': 'Ancré au paragraphe {n}',
  'docx.comments.emptyCommentBody': '(commentaire vide)',
  'docx.comments.reply': 'Répondre',
  'docx.comments.resolve': 'Résoudre',
  'docx.comments.delete': 'Supprimer',
  'docx.comments.unknownAuthor': 'Auteur inconnu',

  // src/docx/editor/SpellCheckMenu.tsx
  'docx.spellCheck.suggestionsAria': 'Suggestions orthographiques pour « {word} »',
  'docx.spellCheck.noSuggestions': 'Aucune suggestion',
  'docx.spellCheck.addToDictionary': 'Ajouter au dictionnaire',

  // src/viewers/DocxViewer.tsx — header/footer panel
  'docx.headerFooter.title': 'En-tête et pied de page',
  'docx.headerFooter.editTitle': 'Modifier l’en-tête et le pied de page',
  'docx.headerFooter.close': 'Fermer l’en-tête et le pied de page',
  'docx.headerFooter.header': 'En-tête',
  'docx.headerFooter.footer': 'Pied de page',
  'docx.headerFooter.partWithType': '{part} ({type})',
  'docx.headerFooter.pageTypeFirst': 'première page',
  'docx.headerFooter.pageTypeEven': 'page paire',
  'docx.headerFooter.lineNumberPrefix': 'Ligne {n} : ',
  'docx.headerFooter.lineLabel': '{base}, ligne {n}',
  'docx.headerFooter.editableTextAria': '{label} — texte modifiable',
  'docx.headerFooter.removeLine': 'Supprimer cette ligne',
  'docx.headerFooter.removeLabel': 'Supprimer {label}',
  'docx.headerFooter.notPlainText': 'Ce n’est pas du texte brut — laissé tel quel.',
  'docx.headerFooter.nonTextParts': 'Les éléments non textuels (images, champs, liens, etc.) sont laissés tels quels.',
  'docx.headerFooter.addLine': '+ Ajouter une ligne',
  'docx.headerFooter.placeholderPageNumber': 'Numéro de page',
  'docx.headerFooter.placeholderTotalPages': 'Nombre total de pages',
  'docx.headerFooter.placeholderDate': 'Date',
  'docx.headerFooter.placeholderField': 'Champ',
  'docx.headerFooter.placeholderComment': 'Commentaire',
  'docx.headerFooter.placeholderFootnote': 'Note de bas de page',
  'docx.headerFooter.placeholderEndnote': 'Note de fin',
  'docx.headerFooter.placeholderOtherContent': 'Autre contenu',
  'docx.headerFooter.placeholderTable': 'Tableau',
  'docx.headerFooter.placeholderUnknownContent': 'Contenu inconnu',
  'docx.headerFooter.placeholderEmptyParagraph': 'Paragraphe vide',
  'docx.headerFooter.placeholderTrackedChange': 'Modification suivie',
  'docx.headerFooter.placeholderLink': 'Lien',
  'docx.headerFooter.placeholderLinkPrefix': 'Lien',

  // src/viewers/DocxViewer.tsx — editor shell (toolbar trailing, zoom, save, errors)
  'docx.viewer.pagesCount': 'Pages : {count}',
  'docx.viewer.zoomGroupAria': 'Zoom',
  'docx.viewer.zoomOut': 'Zoom arrière',
  'docx.viewer.zoomReset': 'Réinitialiser le zoom à 100 %',
  'docx.viewer.zoomResetTitle': 'Réinitialiser le zoom',
  'docx.viewer.zoomIn': 'Zoom avant',
  'docx.viewer.updateFields': 'Mettre à jour les champs',
  'docx.viewer.updateFieldsTitle': 'Mettre à jour les champs — recalculer DATE/TIME/AUTHOR/TITLE/REF/PAGEREF/SEQ/PAGE/NUMPAGES',
  'docx.viewer.updateToc': 'Mettre à jour la table des matières',
  'docx.viewer.printDocument': 'Imprimer le document',
  'docx.viewer.printTitle': 'Imprimer (Ctrl+P)',
  'docx.viewer.saveAs': 'Enregistrer sous',
  'docx.viewer.saveAsTitle': 'Enregistrer sous…',
  'docx.viewer.save': 'Enregistrer',
  'docx.viewer.saveTitle': 'Enregistrer (Ctrl+S)',
  'docx.viewer.editorAria': 'Éditeur de document',
  'docx.viewer.tableEditingAria': 'Modification du tableau',
  'docx.viewer.dismissError': 'Ignorer l’erreur',
  'docx.viewer.dismissMessage': 'Ignorer le message',
  'docx.viewer.layoutFailedTitle': 'Échec de la mise en page du document',
  'docx.viewer.layingOut': 'Mise en page du document…',
  'docx.viewer.blocksProgress': '{completed} / {total} blocs',
  'docx.viewer.saveAsWordFilter': 'Documents Word',
  'docx.viewer.unexpectedTextFile': 'Fichier texte inattendu redirigé vers DocxViewer.',
  'docx.viewer.imagePickerUnavailable': 'Le sélecteur d’images n’est pas disponible dans cet environnement.',
  'docx.viewer.placeCursorBeforeImage': 'Placez le curseur dans le document avant d’insérer une image.',
  'docx.viewer.selectBeforeHyperlink': 'Sélectionnez du texte ou placez le curseur avant d’insérer un lien hypertexte.',
  'docx.viewer.enterUrlPrompt': 'Saisissez une URL',
  'docx.viewer.selectParagraphBeforeList': 'Sélectionnez un paragraphe avant d’activer une liste.',
  'docx.viewer.selectBeforeComment': 'Sélectionnez du texte avant d’ajouter un commentaire.',
  'docx.viewer.addCommentPrompt': 'Ajouter un commentaire',
  'docx.viewer.replyPrompt': 'Répondre',
  'docx.viewer.saveCancelled': 'L’enregistrement a été annulé ou n’est pas disponible.',
  'docx.viewer.noFieldsToUpdate': 'Aucun champ n’avait besoin d’être mis à jour.',
  'docx.viewer.fieldsUpdated': { one: '{count} champ mis à jour.', other: '{count} champs mis à jour.' },
  'docx.viewer.noTocFound': 'Aucune table des matières trouvée à mettre à jour.',
  'docx.viewer.tocUpdated': 'Table des matières mise à jour.',
  // Audit de fidélité à l’enregistrement (round 2, DXS) — texte en langage
  // clair, non technique, pour les constats de `detectLossySaveWarnings`
  // (voir `docx/fidelity/lossySaveWarnings.ts` et
  // `buildFidelitySaveWarningMessage` dans DocxViewer.tsx). Nomme
  // volontairement la fonctionnalité que l’utilisateur reconnaît, jamais
  // l’élément XML sous-jacent.
  'docx.viewer.fidelityWarningIntro': 'Ce document a été enregistré, mais Atlas n’a pas pu tout préserver :',
  'docx.viewer.fidelityWarningContentControls': 'Ses contrôles de contenu (comme des listes déroulantes ou des champs de date) ont conservé leur texte, mais pas les contrôles eux-mêmes.',
  'docx.viewer.fidelityWarningShapeFallback': 'Le dessin de secours d’une forme ou d’une zone de texte n’a pas été préservé.',
  'docx.viewer.fidelityWarningOther': 'Certains autres éléments de mise en forme ou de contenu n’ont pas été entièrement préservés.',

  // src/docx/render/PageView.tsx
  'docx.render.resizeColumn': 'Redimensionner la colonne {n}',

  // src/docx/editor/friendlyDocxError.ts
  'docx.error.notOpenedDetail': 'Ce fichier n’a pas été ouvert : {detail}',
  'docx.error.notSavedDetail': 'Ce fichier n’a pas été enregistré : {detail}',
  'docx.error.invalidWordFile': 'Ce fichier ne semble pas être un document Word valide (.docx). Il est peut-être corrompu, protégé par mot de passe, ou d’un tout autre type. ({detail})',
  'docx.error.saveCorruptedData': 'Le document n’a pas pu être enregistré car ses données semblent corrompues. Réessayez, ou enregistrez une copie sous un nouveau nom. ({detail})',
  'docx.error.openXmlCorrupted': 'La structure interne de ce document semble corrompue et n’a pas pu être lue. Il a peut-être été endommagé par une autre application. ({detail})',
  'docx.error.saveXmlInvalid': 'Le document n’a pas pu être enregistré car Atlas a généré un XML qui a échoué à son propre contrôle de validité. Vos modifications n’ont pas été écrites sur le disque — veuillez réessayer ou signaler ce problème. ({detail})',
  'docx.error.saveVerifyFailed': 'Atlas n’a pas pu vérifier que le fichier enregistré était valide, donc rien n’a été écrit sur le disque. Veuillez réessayer ou signaler ce problème. ({detail})',
  'docx.error.openParseFailed': 'Ce document Word n’a pas pu être ouvert. ({detail})',
  'docx.error.saveParseFailed': 'Ce document Word n’a pas pu être enregistré. ({detail})',
  'docx.error.openGenericFailed': 'Impossible d’ouvrir ce document Word. ({detail})',
  'docx.error.saveGenericFailed': 'Impossible d’enregistrer ce document Word. Vos modifications n’ont pas été écrites sur le disque. ({detail})',

  // src/viewers/shared/SlideDeck.tsx
  'slides.deck.slideTitle': 'Diapositive {n}',
  'slides.deck.thumbnailRailAria': 'Vignettes des diapositives',
  'slides.deck.zoomOut': 'Zoom arrière',
  'slides.deck.zoomIn': 'Zoom avant',
  'slides.deck.zoomReset': 'Réinitialiser le zoom',
  'slides.deck.speakerNotes': 'Notes du présentateur',
  'slides.deck.presentFullscreen': 'Présenter (plein écran)',
  'slides.deck.presenterView': 'Mode Présentateur',
  'slides.deck.noSlidesAvailable': 'Aucune diapositive disponible.',
  'slides.deck.exitPresentation': 'Quitter la présentation (Échap)',
  'slides.deck.closeNotes': 'Fermer les notes',
  'slides.deck.notesPlaceholder': 'Cliquez pour ajouter des notes',

  // src/viewers/shared/SlideEditToolbar.tsx
  'slides.editToolbar.editPresentationAria': 'Modifier la présentation',
  'slides.editToolbar.undo': 'Annuler',
  'slides.editToolbar.redo': 'Rétablir',
  'slides.editToolbar.newSlide': 'Nouvelle diapositive',
  'slides.editToolbar.duplicateSlide': 'Dupliquer la diapositive',
  'slides.editToolbar.deleteSlide': 'Supprimer la diapositive',
  'slides.editToolbar.moveSlideUp': 'Déplacer la diapositive vers le haut',
  'slides.editToolbar.moveSlideDown': 'Déplacer la diapositive vers le bas',
  'slides.editToolbar.insertTextBox': 'Insérer une zone de texte',
  'slides.editToolbar.saveAs': 'Enregistrer sous',
  'slides.editToolbar.save': 'Enregistrer',
  'slides.editToolbar.saveTitle': 'Enregistrer (Ctrl+S)',

  // src/viewers/shared/PresenterView.tsx
  'slides.presenter.ariaLabel': 'Mode Présentateur',
  'slides.presenter.next': 'Suivante',
  'slides.presenter.endOfPresentation': 'Fin de la présentation',
  'slides.presenter.notes': 'Notes',
  'slides.presenter.noNotesForSlide': 'Aucune note pour cette diapositive.',
  'slides.presenter.elapsedTimeAria': 'Temps écoulé',
  'slides.presenter.previousSlideAria': 'Diapositive précédente',
  'slides.presenter.slideCounter': 'Diapositive {current} / {total}',
  'slides.presenter.nextSlideAria': 'Diapositive suivante',
  'slides.presenter.exitPresenterViewAria': 'Quitter le mode Présentateur',

  // src/viewers/shared/SlideEditCanvas.tsx
  'slides.editCanvas.editTextAria': 'Modifier le texte de la diapositive',
  'slides.editCanvas.surfaceAria': 'Zone de modification de la diapositive',

  // src/viewers/pdf/PdfFindBar.tsx
  'pdf.findBar.searching': 'Recherche…',
  'pdf.findBar.placeholder': 'Rechercher dans le document…',
  'pdf.findBar.inputAria': 'Rechercher dans le PDF',
  'pdf.findBar.indexingProgress': 'Indexation {indexed}/{total}…',
  'pdf.findBar.previousMatchTitle': 'Occurrence précédente (Maj+Entrée)',
  'pdf.findBar.nextMatchTitle': 'Occurrence suivante (Entrée)',
  'pdf.findBar.closeAria': 'Fermer la recherche',

  // src/viewers/pdf/PdfPasswordDialog.tsx
  'pdf.password.dialogAria': 'Mot de passe requis',
  'pdf.password.title': 'Ce PDF est protégé par mot de passe',
  'pdf.password.placeholder': 'Saisissez le mot de passe',
  'pdf.password.inputAria': 'Mot de passe du PDF',
  'pdf.password.incorrectError': 'Mot de passe incorrect. Réessayez.',
  'pdf.password.unlock': 'Déverrouiller',

  // src/viewers/pdf/PdfThumbnailRail.tsx
  'pdf.thumbnails.goToPage': 'Aller à la page {n}',
  'pdf.thumbnails.railAria': 'Vignettes des pages',

  // src/formats/legacyOffice.ts (consumed by UnknownViewer.tsx, and reused
  // by LegacyDocViewer.tsx/LegacyPptViewer.tsx for the same format labels)
  'legacyOffice.message': 'Il s’agit d’un {label}. Atlas ne prend pas encore en charge les anciens formats binaires Office — réenregistrez-le au format {modernExtension} dans Word, Excel, PowerPoint ou une application compatible, puis rouvrez-le ici.',
  'legacyOffice.label.doc': 'document Word 97-2003 (.doc)',
  'legacyOffice.label.xls': 'classeur Excel 97-2003 (.xls)',
  'legacyOffice.label.ppt': 'présentation PowerPoint 97-2003 (.ppt)',
  'legacyOffice.label.unknown': 'ancien document Microsoft Office',
  'legacyOffice.modernExtensionUnknown': 'son format XML moderne (.docx / .xlsx / .pptx)',

  // src/viewers/LegacyDocViewer.tsx
  'legacy.doc.unexpectedTextFile': 'LegacyDocViewer a reçu un fichier texte ; un fichier binaire était attendu.',
  'legacy.doc.readError': 'Impossible de lire ce document Word 97-2003 : {error}',
  'legacy.doc.reading': 'Lecture du document Word hérité…',
  'legacy.doc.emptyDocument': 'Ce document ne contient aucun texte lisible.',

  // src/viewers/LegacyPptViewer.tsx
  'legacy.ppt.noSlidesFound': 'Aucune diapositive trouvée dans cette présentation.',
  'legacy.ppt.unexpectedTextFile': 'LegacyPptViewer a reçu un fichier texte ; un fichier binaire était attendu.',
  'legacy.ppt.readError': 'Impossible de lire cette présentation PowerPoint 97-2003 : {error}',
  'legacy.ppt.reading': 'Lecture de la présentation PowerPoint héritée…',

  // src/viewers/UnknownViewer.tsx
  'unknown.viewer.revealRequiresDesktop': 'Afficher dans le dossier nécessite l’application de bureau Atlas.',
  'unknown.viewer.revealFailed': 'Impossible d’afficher ce fichier — il a peut-être été déplacé ou supprimé.',
  'unknown.viewer.back': 'Retour',
  'unknown.viewer.truncatedNotice': 'Affichage limité aux premiers {size}',
  'unknown.viewer.genericExplanation': 'Atlas ne reconnaît pas le format de ce fichier, il n’y a donc rien à prévisualiser ici.',
  'unknown.viewer.openAsText': 'Ouvrir en tant que texte',
  'unknown.viewer.revealInFolder': 'Afficher dans le dossier',

  // src/hooks/useFileHandler.ts
  'fileHandler.formatLabel.docx': 'Word (.docx)',
  'fileHandler.formatLabel.xlsx': 'Excel (.xlsx)',
  'fileHandler.formatLabel.pptx': 'PowerPoint (.pptx)',
  'fileHandler.formatLabel.pdf': 'PDF (.pdf)',
  'fileHandler.formatLabel.odt': 'Texte OpenDocument (.odt)',
  'fileHandler.formatLabel.ods': 'Classeur OpenDocument (.ods)',
  'fileHandler.formatLabel.odp': 'Présentation OpenDocument (.odp)',
  'fileHandler.formatLabel.rtf': 'Texte enrichi (.rtf)',
  'fileHandler.formatLabel.doc': 'Word 97-2003 (.doc)',
  'fileHandler.formatLabel.ppt': 'PowerPoint 97-2003 (.ppt)',
  'fileHandler.extensionMismatchConfirm': '« {name} » ne ressemble pas à un fichier {extFormat} valide — son contenu ressemble plutôt à du {magicFormat}. L’ouvrir quand même ?',
} as const satisfies Record<MessageKey, MessageValue>;

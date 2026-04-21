/**
 * Docs Merge Toolkit
 * Version: 6.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('DOCS_MERGE', {
    SHEET_NAME: SHEET_NAMES.DOCS_MERGE,
    TITLE: '📄 Docs Merge Toolkit',
    MENU_LABEL: '📄 Start Docs Merge',
    MENU_ENTRYPOINT: 'DocsMerge_openSidebar',
    MENU_ORDER: 50,
    SIDEBAR_HTML: 'DocsMerge_Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 1,
    COL_WIDTHS: [120, 200, 250],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 1,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['Generate PDF', 'Generate Doc'] },
            { header: 'Document Name', type: 'TEXT' },
            { header: 'Merged File Link', type: 'URL' }
        ]
    }
});

// Column-index aliases — kept for backward compatibility within this file.
// Metadata (title, sidebar, headers, widths) now lives in SyncEngine.getTool('DOCS_MERGE').
var DOCS_MERGE_CFG = {
  COLUMNS: {
    ACTION: 0, DOC_NAME: 1
  },
  HEADER_ROW: 1
};

/** @deprecated — Use _App_ensureSheetExists('DOCS_MERGE') instead. */
function _DocsMerge_ensureSheetExistsAndActivate() {
  return _App_ensureSheetExists('DOCS_MERGE');
}

/** Opens the Docs Merge sidebar and ensures the sheet exists. */
function DocsMerge_openSidebar() {
  return Logger.run('DOCS_MERGE', 'Open Sidebar', function () {
    _App_launchTool('DOCS_MERGE');
  });
}

function DocsMerge_getConfig() {
  return Logger.run('DOCS_MERGE', 'Get Config', function () {
    var templateUrl = _App_getProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_URL) || "";
    var folderUrl = _App_getProperty(APP_PROPS.DOCS_MERGE_FOLDER_URL) || "";

    var templateName = _App_getProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_NAME) || "";
    if (templateUrl && !templateName) {
      try {
        templateName = DriveApp.getFileById(_DocsMerge_extractIdFromUrl(templateUrl)).getName();
        _App_setProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_NAME, templateName);
      } catch (e) { }
    }

    var folderName = _App_getProperty(APP_PROPS.DOCS_MERGE_FOLDER_NAME) || "";
    if (folderUrl && !folderName) {
      try {
        folderName = DriveApp.getFolderById(_DocsMerge_extractIdFromUrl(folderUrl)).getName();
        _App_setProperty(APP_PROPS.DOCS_MERGE_FOLDER_NAME, folderName);
      } catch (e) { }
    }

    return _App_ok('Configuration loaded.', {
      templateUrl: templateUrl,
      folderUrl: folderUrl,
      templateName: templateName,
      folderName: folderName
    });
  });
}

function DocsMerge_saveConfig(config) {
  return Logger.run('DOCS_MERGE', 'Save Config', function () {
    if (config.templateUrl !== undefined) _App_setProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_URL, config.templateUrl);
    if (config.folderUrl !== undefined) _App_setProperty(APP_PROPS.DOCS_MERGE_FOLDER_URL, config.folderUrl);
    if (config.templateName !== undefined) _App_setProperty(APP_PROPS.DOCS_MERGE_TEMPLATE_NAME, config.templateName);
    if (config.folderName !== undefined) _App_setProperty(APP_PROPS.DOCS_MERGE_FOLDER_NAME, config.folderName);
    return _App_ok('Config saved.');
  });
}

function DocsMerge_searchFolders(query) {
  return Logger.run('DOCS_MERGE', 'Search Folders', function () {
    if (!query || query.length < 2) return _App_ok('Folder search skipped.', { results: [] });
    var results = [];
    try {
      // Search for folders containing the query in the name
      var folders = DriveApp.searchFolders("title contains '" + query.replace(/'/g, "\\'") + "'");
      var count = 0;
      while (folders.hasNext() && count < 10) {
        var f = folders.next();
        results.push({
          id: f.getId(),
          name: f.getName(),
          url: f.getUrl()
        });
        count++;
      }
    } catch (e) {
      // Ignore errors if query is invalid
    }
    return _App_ok('Folder search complete.', { results: results });
  });
}

function DocsMerge_searchDocs(query) {
  return Logger.run('DOCS_MERGE', 'Search Docs', function () {
    if (!query || query.length < 2) return _App_ok('Document search skipped.', { results: [] });
    var results = [];
    try {
      var files = DriveApp.searchFiles("mimeType = 'application/vnd.google-apps.document' and title contains '" + query.replace(/'/g, "\\'") + "'");
      var count = 0;
      while (files.hasNext() && count < 10) {
        var f = files.next();
        results.push({
          id: f.getId(),
          name: f.getName(),
          url: f.getUrl()
        });
        count++;
      }
    } catch (e) {
      // Ignore errors if query is invalid
    }
    return _App_ok('Document search complete.', { results: results });
  });
}


function _DocsMerge_extractIdFromUrl(url) {
  if (!url) return null;
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function DocsMerge_syncPlaceholders(templateUrl) {
  return Logger.run('DOCS_MERGE', 'Sync Placeholders', function () {
    var templateId = _DocsMerge_extractIdFromUrl(templateUrl);
    if (!templateId) throw new Error("Could not extract Template ID. Paste the full Doc URL.");

    try {
      var docText = DocumentApp.openById(templateId).getBody().getText();
      var placeholders = [];
      var regex = /\{\{([^{}]+)\}\}/g;
      var match;

      while ((match = regex.exec(docText)) !== null) {
        if (placeholders.indexOf(match[1]) === -1) placeholders.push(match[1]);
      }

      var syncResult = SheetManager.syncDynamicColumns('DOCS_MERGE', placeholders, {
        anchorHeader: 'Merged File Link',
        dynamicColWidth: 150
      });

      return _App_ok('Synced ' + placeholders.length + ' placeholders.', {
        placeholders: placeholders,
        headers: syncResult.headers
      });
    } catch (e) {
      var toolConfig = SyncEngine.getTool('DOCS_MERGE') || { TITLE: 'DOCS_MERGE' };
      Logger.error(toolConfig.TITLE, 'Sync Placeholders', e);
      return _App_fail("Sync failed: " + e.message + (e.stack ? "\nTrace:\n" + e.stack : "") + ". Ensure you have editor access to the Doc.");
    }
  });
}

function _DocsMerge_replacePlaceholders(body, headers, row) {
  for (var h = 2; h < headers.length; h++) {
    body.replaceText('{{' + headers[h] + '}}', row[h] !== undefined ? String(row[h]) : "");
  }
}

function _DocsMerge_escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// execution context - no longer a single monolith function

function DocsMerge_initExport(config) {
  return Logger.run('DOCS_MERGE', 'Init Export', function () {
    var templateUrl = config.templateUrl;
    var folderUrl = config.folderUrl;
    var mode = config.mode; // 'SINGLE' or 'INDIVIDUAL'

    var templateId = _DocsMerge_extractIdFromUrl(templateUrl);
    var folderId = _DocsMerge_extractIdFromUrl(folderUrl);

    if (!templateId || !folderId) {
      throw new Error("Could not extract IDs from URLs. Please provide full valid URLs.");
    }

    // Save config for future use
    var saveResult = DocsMerge_saveConfig({
      templateUrl: templateUrl,
      folderUrl: folderUrl,
      templateName: config.templateName,
      folderName: config.folderName
    });
    if (!saveResult.success) {
      throw new Error(saveResult.message);
    }

    var sheet = _App_assertActiveSheet(SHEET_NAMES.DOCS_MERGE);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var dataRange = sheet.getDataRange();
    var data = dataRange.getDisplayValues();
    if (data.length < 2) throw new Error("Sheet is empty.");

    try {
      var templateFile = DriveApp.getFileById(templateId);
      var targetFolder = DriveApp.getFolderById(folderId);
    } catch (e) {
      throw new Error("Permission Error: I can't access the Doc or Folder. Make sure you have 'Editor' access.");
    }

    var allRows = data.slice(1);
    var rowsToProcess = [];
    allRows.forEach(function (row, index) {
      var action = (row[DOCS_MERGE_CFG.COLUMNS.ACTION] || "").toString();
      if (action === "Generate PDF" || action === "Generate Doc") {
        rowsToProcess.push({
          row: row,
          index: index,
          action: action
        });
      }
    });

    if (rowsToProcess.length === 0) {
      throw new Error("0 rows had 'Generate PDF' or 'Generate Doc' action.");
    }

    var masterDocId = null;
    if (mode === "SINGLE") {
      try {
        var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        var masterDoc = templateFile.makeCopy('Merged_Doc_' + dateStr);
        masterDocId = masterDoc.getId();
        var masterDocOpen = DocumentApp.openById(masterDocId);
        masterDocOpen.getBody().clear();
        masterDocOpen.saveAndClose();
      } catch (e) {
        throw new Error("Error initializing master document: " + e.message);
      }
    }

    return _App_ok('Export initialized.', {
      rowsToProcess: rowsToProcess,
      masterDocId: masterDocId,
      templateId: templateId,
      folderId: folderId,
      headers: headers
    });
  });
}

function DocsMerge_processRow(item, config, masterDocId, templateId, folderId, headers, isFirst, isLast) {
  return Logger.run('DOCS_MERGE', 'Process Row', function () {
    var sheet = SpreadsheetApp.getActiveSheet();
    var templateFile = DriveApp.getFileById(templateId);
    var targetFolder = DriveApp.getFolderById(folderId);
    var mode = config.mode;
    var rowLinkColName = "Merged File Link";
    var linkColIndex = headers.indexOf(rowLinkColName) + 1; // 1-based index

    var outputFormat = item.action === "Generate PDF" ? "PDF" : "DOC";

    try {
      if (mode === "SINGLE") {
        var tempId = templateFile.makeCopy('Temp_' + item.index).getId();
        var tempDoc = DocumentApp.openById(tempId);
        var tempBody = tempDoc.getBody();

        _DocsMerge_replacePlaceholders(tempBody, headers, item.row);
        tempDoc.saveAndClose();

        var masterOpened = DocumentApp.openById(masterDocId);
        var masterBody = masterOpened.getBody();

        var tempOpened = DocumentApp.openById(tempId);
        var tempBodyOpened = tempOpened.getBody();
        var numChildren = tempBodyOpened.getNumChildren();

        for (var j = 0; j < numChildren; j++) {
          var child = tempBodyOpened.getChild(j).copy();
          var type = child.getType();

          if (isFirst && j === 0) {
            if (type === DocumentApp.ElementType.PARAGRAPH) masterBody.appendParagraph(child.asParagraph());
            else if (type === DocumentApp.ElementType.TABLE) masterBody.appendTable(child.asTable());
            else if (type === DocumentApp.ElementType.LIST_ITEM) masterBody.appendListItem(child.asListItem());
            if (masterBody.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && masterBody.getChild(0).getText() === "") {
              masterBody.removeChild(masterBody.getChild(0)); // Remove default empty paragraph safely
            }
          } else {
            if (type === DocumentApp.ElementType.PARAGRAPH) masterBody.appendParagraph(child.asParagraph());
            else if (type === DocumentApp.ElementType.TABLE) masterBody.appendTable(child.asTable());
            else if (type === DocumentApp.ElementType.LIST_ITEM) masterBody.appendListItem(child.asListItem());
          }
        }

        if (!isLast) {
          masterBody.appendPageBreak();
        }

        masterOpened.saveAndClose();
        DriveApp.getFileById(tempId).setTrashed(true);

        Logger.info(SyncEngine.getTool('DOCS_MERGE').TITLE, 'Row ' + (item.index + 2), "✅ Appended to Master");
        sheet.getRange(item.index + 2, DOCS_MERGE_CFG.COLUMNS.ACTION + 1).setValue("");

        return _App_ok("Row " + (item.index + 1) + " appended.");

      } else {
        // INDIVIDUAL
        var fileName = item.row[DOCS_MERGE_CFG.COLUMNS.DOC_NAME] || 'Document_' + (item.index + 1);
        var tempFile = templateFile.makeCopy(fileName);
        var tempDoc = DocumentApp.openById(tempFile.getId());
        _DocsMerge_replacePlaceholders(tempDoc.getBody(), headers, item.row);
        tempDoc.saveAndClose();

        var finalUrl = "";

        if (outputFormat === "PDF") {
          var pdfBlob = tempFile.getAs(MimeType.PDF);
          var newPdf = targetFolder.createFile(pdfBlob);
          finalUrl = newPdf.getUrl();
          tempFile.setTrashed(true);
        } else {
          // Move doc to target folder
          tempFile.moveTo(targetFolder);
          finalUrl = tempFile.getUrl();
        }

        Logger.info(SyncEngine.getTool('DOCS_MERGE').TITLE, 'Row ' + (item.index + 2), '✅ ' + outputFormat + ' Created');
        sheet.getRange(item.index + 2, DOCS_MERGE_CFG.COLUMNS.ACTION + 1).setValue("");

        // Insert rich text link
        var richText = SpreadsheetApp.newRichTextValue()
          .setText("View File")
          .setLinkUrl(finalUrl)
          .build();

        sheet.getRange(item.index + 2, linkColIndex).setRichTextValue(richText);

        return _App_ok("Created " + outputFormat + " for row " + (item.index + 1));
      }
    } catch (e) {
      Logger.error(SyncEngine.getTool('DOCS_MERGE').TITLE, 'Row ' + (item.index + 2), e);
      return _App_fail("Error on row " + (item.index + 1) + ": " + e.message);
    }
  });
}

function DocsMerge_finishExport(config, masterDocId, folderId, rowsProcessed) {
  return Logger.run('DOCS_MERGE', 'Finish Export', function () {
    if (config.mode === "SINGLE") {
      var targetFolder = DriveApp.getFolderById(folderId);
      var sheet = SpreadsheetApp.getActiveSheet();

      var firstAction = (rowsProcessed && rowsProcessed.length > 0) ? rowsProcessed[0].action : "Generate PDF";
      var isPdf = (firstAction === "Generate PDF");
      var formatName = isPdf ? "PDF" : "Doc";

      var masterFile = DriveApp.getFileById(masterDocId);
      var finalUrl = "";

      if (isPdf) {
        var pdfBlob = masterFile.getAs(MimeType.PDF);
        var newPdf = targetFolder.createFile(pdfBlob);
        finalUrl = newPdf.getUrl();
        masterFile.setTrashed(true);
      } else {
        masterFile.moveTo(targetFolder);
        finalUrl = masterFile.getUrl();
      }

      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var linkColIndex = headers.indexOf("Merged File Link") + 1;

      var richText = SpreadsheetApp.newRichTextValue()
        .setText("View Master File")
        .setLinkUrl(finalUrl)
        .build();

      rowsProcessed.forEach(function (item) {
        sheet.getRange(item.index + 2, linkColIndex).setRichTextValue(richText);
        Logger.info(SyncEngine.getTool('DOCS_MERGE').TITLE, 'Row ' + (item.index + 2), "✅ Merged into Single " + formatName);
      });

      return _App_ok("Successfully generated and linked Master " + formatName + ".");
    }
    return _App_ok("All individual documents processed successfully.");
  });
}

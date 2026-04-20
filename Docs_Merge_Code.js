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
    var prefs = SyncEngine.getPrefs('DOCS_MERGE');
    var templateUrl = prefs.templateUrl || "";
    var folderUrl = prefs.folderUrl || "";

    var templateName = prefs.templateName || "";
    if (templateUrl && !templateName) {
      try {
        templateName = DriveApp.getFileById(_DocsMerge_extractIdFromUrl(templateUrl)).getName();
        prefs.templateName = templateName;
        SyncEngine.setPrefs('DOCS_MERGE', prefs);
      } catch (e) { }
    }

    var folderName = prefs.folderName || "";
    if (folderUrl && !folderName) {
      try {
        folderName = DriveApp.getFolderById(_DocsMerge_extractIdFromUrl(folderUrl)).getName();
        prefs.folderName = folderName;
        SyncEngine.setPrefs('DOCS_MERGE', prefs);
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
    var prefs = SyncEngine.getPrefs('DOCS_MERGE');
    if (config.templateUrl !== undefined) prefs.templateUrl = config.templateUrl;
    if (config.folderUrl !== undefined) prefs.folderUrl = config.folderUrl;
    if (config.templateName !== undefined) prefs.templateName = config.templateName;
    if (config.folderName !== undefined) prefs.folderName = config.folderName;
    SyncEngine.setPrefs('DOCS_MERGE', prefs);
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
  headers.forEach(function(header) {
    if (!header || header === 'Action' || header === 'Merged File Link') return;
    var val = row[header];
    var valStr = (val === undefined || val === null || val === "") ? "" : String(val);
    body.replaceText('{{' + _DocsMerge_escapeRegExp(header) + '}}', valStr);
  });
}

function _DocsMerge_escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// execution context - no longer a single monolith function

function DocsMerge_runSync(config) {
  return Logger.run('DOCS_MERGE', 'Run Docs Merge', function () {
    var templateUrl = config.templateUrl;
    var folderUrl = config.folderUrl;
    var mode = config.mode || 'INDIVIDUAL'; // 'SINGLE' or 'INDIVIDUAL'

    var templateId = _DocsMerge_extractIdFromUrl(templateUrl);
    var folderId = _DocsMerge_extractIdFromUrl(folderUrl);

    if (!templateId || !folderId) {
      throw new Error("Could not extract IDs from URLs. Please provide full valid URLs.");
    }

    // Save config for future use
    DocsMerge_saveConfig({
      templateUrl: templateUrl,
      folderUrl: folderUrl,
      templateName: config.templateName,
      folderName: config.folderName
    });

    var templateFile = DriveApp.getFileById(templateId);
    var targetFolder = DriveApp.getFolderById(folderId);
    var toolCfg = SyncEngine.getTool('DOCS_MERGE');
    var headers = toolCfg.HEADERS;
    
    // 1. Initialization for SINGLE mode
    var masterDocId = null;
    var masterOpened = null;
    var masterBody = null;
    if (mode === "SINGLE") {
      var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      var masterDoc = templateFile.makeCopy('Merged_Doc_' + dateStr);
      masterDocId = masterDoc.getId();
      masterOpened = DocumentApp.openById(masterDocId);
      masterBody = masterOpened.getBody();
      masterBody.clear();
    }

    var processedRows = [];

    // 2. Execution via ExecutionService
    var stats = ExecutionService.processPendingRows('DOCS_MERGE', function(rowObj) {
        var action = String(rowObj['Action'] || '').toUpperCase();
        var outputFormat = action === "GENERATE PDF" ? "PDF" : "DOC";
        var isFirst = processedRows.length === 0;

        if (mode === "SINGLE") {
          // Process for Master Doc
          var tempId = templateFile.makeCopy('Temp_' + rowObj._rowNumber).getId();
          var tempDoc = DocumentApp.openById(tempId);
          _DocsMerge_replacePlaceholders(tempDoc.getBody(), headers, rowObj);
          tempDoc.saveAndClose();

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
                masterBody.removeChild(masterBody.getChild(0));
              }
            } else {
              if (type === DocumentApp.ElementType.PARAGRAPH) masterBody.appendParagraph(child.asParagraph());
              else if (type === DocumentApp.ElementType.TABLE) masterBody.appendTable(child.asTable());
              else if (type === DocumentApp.ElementType.LIST_ITEM) masterBody.appendListItem(child.asListItem());
            }
          }
          masterBody.appendPageBreak();
          
          tempOpened.saveAndClose();
          DriveApp.getFileById(tempId).setTrashed(true);
          
          processedRows.push(rowObj._rowNumber);
          SheetManager.patchRow('DOCS_MERGE', rowObj._rowNumber, { 'Action': '' });

        } else {
          // INDIVIDUAL
          var fileName = rowObj['Document Name'] || 'Document_' + rowObj._rowNumber;
          var tempFile = templateFile.makeCopy(fileName);
          var tempDoc = DocumentApp.openById(tempFile.getId());
          _DocsMerge_replacePlaceholders(tempDoc.getBody(), headers, rowObj);
          tempDoc.saveAndClose();

          var finalUrl = "";
          if (outputFormat === "PDF") {
            var pdfBlob = tempFile.getAs(MimeType.PDF);
            var newPdf = targetFolder.createFile(pdfBlob);
            finalUrl = newPdf.getUrl();
            tempFile.setTrashed(true);
          } else {
            tempFile.moveTo(targetFolder);
            finalUrl = tempFile.getUrl();
          }

          var richText = SpreadsheetApp.newRichTextValue()
            .setText("View File")
            .setLinkUrl(finalUrl)
            .build();

          SheetManager.patchRow('DOCS_MERGE', rowObj._rowNumber, { 
            'Action': '',
            'Merged File Link': finalUrl
          });
          
          // Apply rich text separately
          var sheet = SheetManager.getSheet('DOCS_MERGE');
          var hMap = SheetManager.getHeaderMap('DOCS_MERGE');
          sheet.getRange(rowObj._rowNumber, hMap['Merged File Link']).setRichTextValue(richText);
        }
    });

    // 3. Finalization
    if (mode === "SINGLE" && masterDocId) {
      // Remove trailing page break if any
      var lastChild = masterBody.getChild(masterBody.getNumChildren() - 1);
      if (lastChild && lastChild.getType() === DocumentApp.ElementType.PAGE_BREAK) {
        masterBody.removeChild(lastChild);
      }
      masterOpened.saveAndClose();

      var finalMasterFile = DriveApp.getFileById(masterDocId);
      var isPdf = true;
      
      var finalUrl = "";
      if (isPdf) {
        var pdfBlob = finalMasterFile.getAs(MimeType.PDF);
        var newPdf = targetFolder.createFile(pdfBlob);
        finalUrl = newPdf.getUrl();
        finalMasterFile.setTrashed(true);
      } else {
        finalMasterFile.moveTo(targetFolder);
        finalUrl = finalMasterFile.getUrl();
      }

      var richText = SpreadsheetApp.newRichTextValue()
        .setText("View Master File")
        .setLinkUrl(finalUrl)
        .build();

      var sheet = SheetManager.getSheet('DOCS_MERGE');
      var hMap = SheetManager.getHeaderMap('DOCS_MERGE');
      processedRows.forEach(function(rowNum) {
        sheet.getRange(rowNum, hMap['Merged File Link']).setRichTextValue(richText);
      });
    }

    if (stats.processed === 0 && stats.errors === 0) {
      return _App_ok("No data to merge.");
    }

    return _App_ok("Merge Complete. Success: " + stats.processed + ", Errors: " + stats.errors);
  });
}

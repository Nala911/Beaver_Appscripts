/**
 * Docs Merge
 * Version: 6.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('DOCS_MERGE', {
    SHEET_NAME: SHEET_NAMES.DOCS_MERGE,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.DOCS_MERGE,
    MENU_LABEL: SHEET_NAMES.DOCS_MERGE,
    MENU_ENTRYPOINT: 'DocsMerge_openSidebar',
    MENU_ORDER: 50,
    SIDEBAR_HTML: 'DocsMerge_Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['Generate PDF', 'Generate Doc'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Document Name', type: 'TEXT' },
            { header: 'Merged File Link', type: 'URL' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to merge your first document:</p><ol><li><strong>Prepare Doc:</strong> Create a Google Doc with <code>{{placeholder}}</code> tags.</li><li><strong>Setup URLs:</strong> Search and select your Template Doc and Destination Folder.</li><li><strong>Sync & Run:</strong> Click <strong>Pull Placeholders</strong> to add columns, then click <strong>Run</strong>.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Learn about Action, Document Name, and Merged File Link.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Action:</strong> <code>Generate PDF</code> or <code>Generate Doc</code>.</li><li><strong>Document Name:</strong> The filename for the generated document.</li><li><strong>Merged File Link:</strong> This cell will be updated with a link to the final file.</li><li><strong>Dynamic Columns:</strong> These are created automatically based on your <code>{{tags}}</code>.</li></ul>"
            },
            {
                icon: "lightbulb",
                color: "var(--warning)",
                label: "Pro Tips",
                shortDesc: "Single vs Individual mode and placeholder sync.",
                tooltipId: "help-tips",
                tooltipContent: "<p><strong>Pro Tips</strong></p><ul><li><strong>Individual Mode:</strong> Creates a separate file for every row in the folder.</li><li><strong>Single Mode:</strong> Merges all rows into one giant document (with page breaks).</li><li><strong>Permissions:</strong> Ensure you have 'Editor' access to the template and destination folder.</li></ul>"
            }
        ]
    }
});



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
      return _App_fail("Sync failed: " + e.message + ". Ensure you have editor access to the Doc.");
    }
  });
}

function _DocsMerge_replacePlaceholders(body, headers, rowObj) {
  // Headers start from Action(0), Doc Name(1), Merged Link(2). Dynamic placeholders start from index 3.
  for (var h = 3; h < headers.length; h++) {
    var key = headers[h];
    body.replaceText('{{' + key + '}}', rowObj[key] !== undefined ? String(rowObj[key]) : "");
  }
}

function _DocsMerge_escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// execution context - no longer a single monolith function

function DocsMerge_executeBatch(config, startIndex) {
  return Logger.run('DOCS_MERGE', 'Execute Batch', function () {
    var start = startIndex || 0;
    var batchSize = 10;
    var mode = config.mode || "INDIVIDUAL";

    var pendingRows = SheetManager.readPendingObjects('DOCS_MERGE', { useDisplayValues: true });

    if (pendingRows.length === 0) return _App_ok(start > 0 ? "Batch finished!" : "Nothing to do! No 'Generate PDF' or 'Generate Doc' actions pending.", { completed: true, message: start > 0 ? "Batch finished!" : "Nothing to do! No 'Generate PDF' or 'Generate Doc' actions pending." });
    if (start >= pendingRows.length) return _App_ok("Batch complete!", { completed: true, message: "Batch complete!" });

    var batchItems = pendingRows.slice(start, start + batchSize);
    var remainingPending = pendingRows.length - (start + batchItems.length);

    var templateId = _DocsMerge_extractIdFromUrl(config.templateUrl);
    var folderId = _DocsMerge_extractIdFromUrl(config.folderUrl);

    if (!templateId || !folderId) {
      throw new Error("Could not extract IDs from URLs. Please provide full valid URLs.");
    }

    if (start === 0) {
      // Save config for future use on first run
      var saveResult = DocsMerge_saveConfig({
        templateUrl: config.templateUrl,
        folderUrl: config.folderUrl,
        templateName: config.templateName,
        folderName: config.folderName
      });
      if (!saveResult.success) throw new Error(saveResult.message);
    }

    var templateFile = null;
    var targetFolder = null;
    try {
      templateFile = DriveApp.getFileById(templateId);
      targetFolder = DriveApp.getFolderById(folderId);
    } catch (e) {
      throw new Error("Permission Error: I can't access the Doc or Folder. Make sure you have 'Editor' access.");
    }

    var masterDocId = null;
    if (mode === "SINGLE") {
      if (start === 0) {
        var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        var masterDoc = templateFile.makeCopy('Merged_Doc_' + dateStr);
        masterDocId = masterDoc.getId();
        _App_setProperty(APP_PROPS.DOCS_MERGE_MASTER_DOC_ID, masterDocId);

        var masterDocOpen = DocumentApp.openById(masterDocId);
        masterDocOpen.getBody().clear();
        masterDocOpen.saveAndClose();
      } else {
        masterDocId = _App_getProperty(APP_PROPS.DOCS_MERGE_MASTER_DOC_ID);
        if (!masterDocId) throw new Error("Could not find master document ID to resume batch.");
      }
    }

    var headers = SheetManager.getHeaders('DOCS_MERGE');
    var rowLinkColName = "Merged File Link";
    var linkColIndex = headers.indexOf(rowLinkColName) + 1; // 1-based index

    var isLastBatch = (start + batchItems.length) >= pendingRows.length;

    var stats = _App_BatchProcessor('DOCS_MERGE', batchItems, function (item, index) {
      var rowUpdates = {
        action: item['Action'],
        _rowNumber: item._rowNumber,
        status: "",
        linkUrl: null
      };
      
      var isFirstInWholeProcess = (start === 0 && index === 0);
      var isLastInWholeProcess = (isLastBatch && index === (batchItems.length - 1));
      var outputFormat = item['Action'] === "Generate PDF" ? "PDF" : "DOC";

      if (mode === "SINGLE") {
          var tempId = templateFile.makeCopy('Temp_' + item._rowNumber).getId();
          var tempDoc = DocumentApp.openById(tempId);
          var tempBody = tempDoc.getBody();

          _DocsMerge_replacePlaceholders(tempBody, headers, item);
          tempDoc.saveAndClose();

          var masterOpened = DocumentApp.openById(masterDocId);
          var masterBody = masterOpened.getBody();

          var tempOpened = DocumentApp.openById(tempId);
          var tempBodyOpened = tempOpened.getBody();
          var numChildren = tempBodyOpened.getNumChildren();

          for (var j = 0; j < numChildren; j++) {
            var child = tempBodyOpened.getChild(j).copy();
            var type = child.getType();

            if (isFirstInWholeProcess && j === 0) {
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

          if (!isLastInWholeProcess) {
            masterBody.appendPageBreak();
          }

          masterOpened.saveAndClose();
          DriveApp.getFileById(tempId).setTrashed(true);

          rowUpdates.status = _App_formatStatus('SUCCESS', "Appended to Master");
          rowUpdates.action = "";
        } else {
          // INDIVIDUAL
          var fileName = item['Document Name'] || 'Document_' + item._rowNumber;
          var tempFile = templateFile.makeCopy(fileName);
          var tempDoc = DocumentApp.openById(tempFile.getId());
          _DocsMerge_replacePlaceholders(tempDoc.getBody(), headers, item);
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

          rowUpdates.status = _App_formatStatus('SUCCESS', outputFormat + ' Created');
          rowUpdates.linkUrl = finalUrl;
          rowUpdates.action = "";
        }

        return rowUpdates;

    }, {
      onBatchComplete: function (batchResults) {
        var sheet = SpreadsheetApp.getActiveSheet();
        
        batchResults.forEach(function (res) {
          if (res && res._rowNumber !== undefined && !res.isError) {
            if (res.linkUrl && linkColIndex > 0) {
              var richText = SpreadsheetApp.newRichTextValue()
                .setText("View File")
                .setLinkUrl(res.linkUrl)
                .build();
              sheet.getRange(res._rowNumber, linkColIndex).setRichTextValue(richText);
            }
          }
        });
        
        _App_batchPatchResults('DOCS_MERGE', batchResults);
      }
    });

    // Handle finish step for SINGLE mode if it's the absolute last batch
    if (mode === "SINGLE" && isLastBatch) {
      try {
        var masterFile = DriveApp.getFileById(masterDocId);
        var finalMasterUrl = "";
        
        // Find if we should output PDF based on the first pending row's original action
        var isPdf = pendingRows[0] && pendingRows[0]['Action'] === "Generate PDF";
        
        if (isPdf) {
          var masterPdfBlob = masterFile.getAs(MimeType.PDF);
          var masterPdfFile = targetFolder.createFile(masterPdfBlob);
          finalMasterUrl = masterPdfFile.getUrl();
          masterFile.setTrashed(true);
        } else {
          masterFile.moveTo(targetFolder);
          finalMasterUrl = masterFile.getUrl();
        }

        var sheet = SpreadsheetApp.getActiveSheet();
        var richTextMaster = SpreadsheetApp.newRichTextValue()
          .setText("View Master File")
          .setLinkUrl(finalMasterUrl)
          .build();
          
        if (linkColIndex > 0) {
          pendingRows.forEach(function (item) {
             sheet.getRange(item._rowNumber, linkColIndex).setRichTextValue(richTextMaster);
          });
        }
      } catch (err) {
        throw new Error("Finish Export Failed: " + err.message);
      }
    }

    return _App_ok('Processed Docs Merge batch.', {
      completed: stats.processedCount + stats.errorCount >= pendingRows.length - start,
      nextIndex: start + batchItems.length,
      remainingPending: remainingPending,
      processed: stats.processedCount
    });
  });
}

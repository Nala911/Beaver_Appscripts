/**
 * Bulk Folder Creation
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('BULK_FOLDER', {
    SHEET_NAME: SHEET_NAMES.BULK_FOLDER,
    TITLE: SHEET_NAMES.BULK_FOLDER,
    MENU_LABEL: SHEET_NAMES.BULK_FOLDER,
    MENU_ENTRYPOINT: 'BulkFolderCreation_openSidebar',
    MENU_ORDER: 80,
    SIDEBAR_HTML: 'BulkFolderCreation_Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 1,
    COL_WIDTHS: [100, 200, 200, 200],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 0,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['Create'] },
            { header: 'Level 1', type: 'TEXT' },
            { header: 'Level 2', type: 'TEXT' },
            { header: 'Level 3', type: 'TEXT' }
        ]
    }
});

// Column-index aliases — kept for backward compatibility.
// Metadata (title, sidebar, headers, widths) now lives in SyncEngine.getTool('BULK_FOLDER').
var BULKFOLDER_COL = {
  ACTION: 0
};

// --- MENU & UI HANDLERS ---

/** Opens the Bulk Folder sidebar and ensures the sheet exists. */
function BulkFolderCreation_openSidebar() {
  return Logger.run('BULK_FOLDER', 'Open Sidebar', function () {
    _App_launchTool('BULK_FOLDER');
  });
}

// --- EXPLORER LOGIC ---

function BulkFolderCreation_getDriveNavData(folderId) {
  return Logger.run('BULK_FOLDER', 'Fetch Nav Data', function () {
    try {
      var folder;
      if (!folderId || folderId === 'root') {
        folder = DriveApp.getRootFolder();
      } else {
        folder = DriveApp.getFolderById(folderId);
      }

      var currentId = folder.getId();
      var currentName = folder.getName();

      var breadcrumbs = [];
      var parent = folder;
      var depth = 0;
      while (depth < 5) {
        try {
          breadcrumbs.unshift({ id: parent.getId(), name: parent.getName() });
          var parents = parent.getParents();
          if (parents.hasNext()) {
            parent = parents.next();
          } else {
            break;
          }
        } catch (e) {
          break;
        }
        depth++;
      }

      var folders = folder.getFolders();
      var folderList = [];
      while (folders.hasNext()) {
        var f = folders.next();
        folderList.push({
          id: f.getId(),
          name: f.getName()
        });
      }

      folderList.sort(function (a, b) { return a.name.localeCompare(b.name); });

      return _App_ok('Navigation data loaded', {
        current: { id: currentId, name: currentName },
        breadcrumbs: breadcrumbs,
        children: folderList
      });

    } catch (e) {
      throw new Error("Error fetching Drive data: " + e.message);
    }
  });
}

// --- BATCH CREATION LOGIC ---

function BulkFolderCreation_getProgress() {
  return Logger.run('BULK_FOLDER', 'Get Progress', function () {
    return _App_ok('Progress', _App_getProgress('BULK_FOLDER'));
  });
}

function BulkFolderCreation_runBulkCreationSequence(targetFolderId) {
  return Logger.run('BULK_FOLDER', 'Batch Creation', function () {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) throw new Error("⚠️ System is busy. Please try again.");

    try {
      _App_resetExecutionTimer();
      
      var pendingRows = SheetManager.readPendingObjects('BULK_FOLDER');

      if (pendingRows.length === 0) {
        Logger.warn(SyncEngine.getTool('BULK_FOLDER').TITLE, 'Global', "No pending 'Create' actions found.");
        return _App_ok("No pending 'Create' actions found.");
      }

      var headers = SheetManager.getHeaders('BULK_FOLDER');
      var levelCols = headers.filter(h => h.toLowerCase().startsWith('level'));

      // --- PRE-VALIDATION START ---
      var gapErrors = [];
      var emptyErrors = [];
      for (var k = 0; k < pendingRows.length; k++) {
        var item = pendingRows[k];
        var rowNum = item._rowNumber;
        var hasEmptyLevel = false;
        var hasDataAfterEmpty = false;
        var hasAnyData = false;

        for (var c = 0; c < levelCols.length; c++) {
          var header = levelCols[c];
          var fName = String(item[header] || "").trim();

          if (fName === "") {
            hasEmptyLevel = true;
          } else {
            hasAnyData = true;
            if (hasEmptyLevel) {
              hasDataAfterEmpty = true;
              break;
            }
          }
        }

        if (!hasAnyData) {
          emptyErrors.push("Row " + rowNum);
        } else if (hasDataAfterEmpty) {
          gapErrors.push("Row " + rowNum);
        }
      }

      if (gapErrors.length > 0 || emptyErrors.length > 0) {
        var errMsgs = [];
        if (gapErrors.length > 0) {
          errMsgs.push("Missing intermediate folder names in rows: " + gapErrors.join(", ") + " (e.g., Level 1 is empty, but Level 2 has data)");
        }
        if (emptyErrors.length > 0) {
          errMsgs.push("No folder names specified in rows: " + emptyErrors.join(", "));
        }
        var fullError = "⚠️ Validation Error:\n" + errMsgs.join("\n");
        Logger.warn(SyncEngine.getTool('BULK_FOLDER').TITLE, 'Pre-Validation', fullError);
        return _App_fail(fullError + "\nPlease fix and try again.");
      }
      // --- PRE-VALIDATION END ---

      var folderCache = {};
      var stats = _App_BatchProcessor('BULK_FOLDER', pendingRows, function (item) {
        var rowNum = item._rowNumber;
        var folderNames = [];
        for (var c = 0; c < levelCols.length; c++) {
          var header = levelCols[c];
          var fName = String(item[header] || "").trim();
          if (fName) {
            folderNames.push(fName.replace(/[\\/?*]/g, "_"));
          }
        }

        if (folderNames.length === 0) {
          throw new Error("No folder names specified in Level columns.");
        }

        _BulkFolderCreation_createFolderPath(targetFolderId, folderNames, folderCache);

        Logger.success(SyncEngine.getTool('BULK_FOLDER').TITLE, 'Row ' + rowNum, '✅ Created: ' + folderNames.join('/'));
        return { _rowNumber: rowNum };

      }, {
        onBatchComplete: function (results) {
          var rowNumbers = [];
          var updates = [];
          results.forEach(function (res) {
            if (res && res._rowNumber) {
              rowNumbers.push(res._rowNumber);
              updates.push({ 'Action': '' });
            }
          });
          if (rowNumbers.length > 0) {
            SheetManager.batchPatchRows('BULK_FOLDER', rowNumbers, updates);
          }
        }
      });

      var finalMsg = "Successfully processed " + stats.processedCount + " folders.";
      if (stats.errorCount > 0) finalMsg += " (" + stats.errorCount + " errors)";
      if (stats.timeLimitReached) finalMsg = "⏳ Time limit reached. " + finalMsg;

      return _App_ok(finalMsg);

    } finally {
      lock.releaseLock();
    }
  });
}

function _BulkFolderCreation_createFolderPath(baseFolderId, folderNamesArr, folderCache) {
  var currentParentId = baseFolderId === "root" ? DriveApp.getRootFolder().getId() : baseFolderId;

  for (var i = 0; i < folderNamesArr.length; i++) {
    var fName = folderNamesArr[i];
    var cacheKey = currentParentId + "_" + fName;

    // 1. Check in-memory Cache
    if (folderCache[cacheKey]) {
      currentParentId = folderCache[cacheKey];
      continue;
    }

    // 2. Check Drive API if not in Cache
    var query = "'" + currentParentId + "' in parents and name = '" + fName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false";

    var result = _App_callWithBackoff(function () {
      return Drive.Files.list({ q: query, fields: "files(id, name)", pageSize: 1 });
    });

    if (result.files && result.files.length > 0) {
      currentParentId = result.files[0].id; // Step inside existing
    } else {
      // 3. Create it
      var resource = {
        name: fName,
        parents: [currentParentId],
        mimeType: 'application/vnd.google-apps.folder'
      };
      var newFolder = _App_callWithBackoff(function () {
        return Drive.Files.create(resource, null, { fields: 'id' });
      });
      currentParentId = newFolder.id;
    }

    // Save to Cache for subsequent rows
    folderCache[cacheKey] = currentParentId;
  }
  return currentParentId;
}



/** @deprecated — Use _App_ensureSheetExists('BULK_FOLDER') instead; it now handles setup. */
function BulkFolderCreation_setupSheet(sheet) {
  if (!sheet) {
    return _App_ensureSheetExists('BULK_FOLDER');
  }
  // For direct calls with an existing sheet (e.g., resetting), delegate setup manually
  _App_applyBodyFormatting(sheet, 0, SyncEngine.getTool('BULK_FOLDER').FORMAT_CONFIG);
  return _App_ok('Sheet has been setup successfully for Bulk Folder Creation.');
}

function _BulkFolderCreation_initializeHeaders(sheet) {
  var allHeaders = SyncEngine.getTool('BULK_FOLDER').HEADERS;
  sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders])
    .setFontWeight(SHEET_THEME.LAYOUT.HEADER_WEIGHT)
    .setBackground(SHEET_THEME.HEADER)
    .setFontColor(SHEET_THEME.TEXT)
    .setBorder(true, true, true, true, true, true, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1); // freeze system columns (Action)

  // Set widths
  if (SyncEngine.getTool('BULK_FOLDER').COL_WIDTHS) {
    SyncEngine.getTool('BULK_FOLDER').COL_WIDTHS.forEach(function (w, i) {
      if (w !== null) sheet.setColumnWidth(i + 1, w);
    });
  }
}

// Stage 1: Data validations only (body formatting handled by _App_applyBodyFormatting)
function _BulkFolderCreation_applyDataValidations(sheet) {
  var maxRows = sheet.getMaxRows();
  if (maxRows < 2) return;

  var actionRule = SpreadsheetApp.newDataValidation().requireValueInList(['Create'], true).setAllowInvalid(false).build();
  sheet.getRange(2, BULKFOLDER_COL.ACTION + 1, maxRows - 1, 1).setDataValidation(actionRule);
}

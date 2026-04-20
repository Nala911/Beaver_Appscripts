/**
 * Bulk Drive Automator
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('BULK_FOLDER', {
    SHEET_NAME: SHEET_NAMES.BULK_FOLDER,
    TITLE: '📂 Bulk Drive Automator',
    MENU_LABEL: '📂 Bulk Folder Creation',
    MENU_ENTRYPOINT: 'BulkFolder_showSidebar',
    MENU_ORDER: 80,
    SIDEBAR_HTML: 'BulkFolder_Sidebar',
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

// Global Time Limit (Google Apps Script has 6 min limit, we stop at 5.5 min)
var BULKFOLDER_START_TIME = 0;
var BULKFOLDER_MAX_EXECUTION_TIME = 330 * 1000; // 5.5 minutes

function _BulkFolder_isTimeLimitReached() {
  return (Date.now() - BULKFOLDER_START_TIME > BULKFOLDER_MAX_EXECUTION_TIME);
}

// --- MENU & UI HANDLERS ---

/** @deprecated — Use _App_ensureSheetExists('BULK_FOLDER') instead. */
function _BulkFolder_ensureSheetExistsAndActivate() {
  return _App_ensureSheetExists('BULK_FOLDER');
}

/** Opens the Bulk Folder sidebar and ensures the sheet exists. */
function BulkFolder_showSidebar() {
  return Logger.run('BULK_FOLDER', 'Open Sidebar', function () {
    _App_launchTool('BULK_FOLDER');
  });
}

/** Alias for menu compatibility. */
function BulkFolder_showSidebarbulkcreation() {
  BulkFolder_showSidebar();
}

// --- EXPLORER LOGIC ---

function BulkFolder_getDriveNavData(folderId) {
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

      return {
        current: { id: currentId, name: currentName },
        breadcrumbs: breadcrumbs,
        children: folderList
      };

    } catch (e) {
      throw new Error("Error fetching Drive data: " + e.message);
    }
  });
}

// --- BATCH CREATION LOGIC ---

function BulkFolder_getProgress() {
  return _App_getProgress('BULK_FOLDER');
}

function BulkFolder_runBulkCreationSequence(targetFolderId) {
  return Logger.run('BULK_FOLDER', 'Batch Creation', function () {
    return _App_withDocumentLock('Bulk Creation', function() {
      BULKFOLDER_START_TIME = Date.now();
      var folderCache = {};
      
      // Identify Level columns from the tool's schema
      var toolCfg = SyncEngine.getTool('BULK_FOLDER');
      var levelHeaders = toolCfg.HEADERS.filter(function(h) { 
          return String(h).toLowerCase().indexOf('level') !== -1; 
      });

      var stats = ExecutionService.processPendingRows('BULK_FOLDER', function(rowObj) {
          if (_BulkFolder_isTimeLimitReached()) throw new Error("⏳ Time limit reached.");

          var folderNames = [];
          levelHeaders.forEach(function(h) {
              var fName = String(rowObj[h] || "").trim();
              if (fName) folderNames.push(fName.replace(/[\\/?*]/g, "_"));
          });

          if (folderNames.length === 0) throw new Error("No folder names specified.");

          _BulkFolder_createFolderPath(targetFolderId, folderNames, folderCache);

          // Update success
          SheetManager.patchRow('BULK_FOLDER', rowObj._rowNumber, {
              'Action': '',
              'Log': '✅ Created: ' + folderNames.join('/')
          });
      });

      if (stats.processed === 0 && stats.errors === 0) {
        return "No pending 'Create' actions found.";
      }

      var finalMsg = "Successfully processed " + stats.processed + " folders.";
      if (stats.errors > 0) finalMsg += " (" + stats.errors + " errors)";
      
      return finalMsg;
    });
  });
}

function _BulkFolder_createFolderPath(baseFolderId, folderNamesArr, folderCache) {
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
function BulkFolder_setupSheet_bulkcreation(sheet) {
  if (!sheet) {
    return _App_ensureSheetExists('BULK_FOLDER');
  }
  // For direct calls with an existing sheet (e.g., resetting), delegate setup manually
  _App_applyBodyFormatting(sheet, 0, SyncEngine.getTool('BULK_FOLDER').FORMAT_CONFIG);
  return "Sheet has been setup successfully for Bulk Folder Creation.";
}

function _BulkFolder_initializeHeaders(sheet) {
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
function _BulkFolder_applyDataValidations(sheet) {
  var maxRows = sheet.getMaxRows();
  if (maxRows < 2) return;

  var actionRule = SpreadsheetApp.newDataValidation().requireValueInList(['Create'], true).setAllowInvalid(false).build();
  sheet.getRange(2, BULKFOLDER_COL.ACTION + 1, maxRows - 1, 1).setDataValidation(actionRule);
}

/**
 * Bulk Folder Creation
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('BULK_FOLDER', {
    SHEET_NAME: SHEET_NAMES.BULK_FOLDER,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.BULK_FOLDER,
    MENU_LABEL: SHEET_NAMES.BULK_FOLDER,
    MENU_ENTRYPOINT: 'BulkFolderCreation_openSidebar',
    MENU_ORDER: 80,
    SIDEBAR_HTML: 'tools/BulkFolderCreation_Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Level 1', type: 'TEXT' },
            { header: 'Level 2', type: 'TEXT' },
            { header: 'Level 3', type: 'TEXT' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to create folders in bulk:</p><ol><li><strong>Select Destination:</strong> Pick the parent folder in the sidebar explorer.</li><li><strong>Set Action:</strong> Set the Action column to <code>CREATE</code>.</li><li><strong>Enter Levels:</strong> Provide folder names in Level 1, Level 2, Level 3 to define nested structures.</li><li><strong>Run:</strong> Click <strong>Run Bulk Creation</strong> in the sidebar.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Action, Status, and Levels.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Action:</strong> Set to <code>CREATE</code> to create folders on Drive.</li><li><strong>Level 1/2/3:</strong> Folders are created nested. e.g. <code>Level 1/Level 2/Level 3</code>.</li><li><strong>Duplicate Check:</strong> If a folder structure already exists, the tool will step inside rather than duplicating it.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getDriveNavData: function(folderId) {
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
        },
        createFolders: function(targetFolderId) {
            return _BulkFolderCreation_runBulkCreationSequence(targetFolderId);
        }
    }
});

// Column-index aliases — kept for backward compatibility.
// Metadata (title, sidebar, headers, widths) now lives in SyncEngine.getTool('BULK_FOLDER').
var BULKFOLDER_COL = {
  ACTION: 0,
  STATUS: 1
};

// --- MENU & UI HANDLERS ---

/** Opens the Bulk Folder sidebar and ensures the sheet exists. */
function BulkFolderCreation_openSidebar() {
  return Logger.run('BULK_FOLDER', 'Open Sidebar', function () {
    _App_launchTool('BULK_FOLDER');
  });
}

// --- EXPLORER LOGIC ---

function _BulkFolderCreation_runBulkCreationSequence(targetFolderId) {
  return Logger.run('BULK_FOLDER', 'Batch Creation', function () {
    return _App_withDocumentLock('BULK_FOLDER_CREATION', function () {
      _App_resetExecutionTimer();
      
      var pendingRows = SheetManager.readPendingObjects('BULK_FOLDER');

      if (pendingRows.length === 0) {
        Logger.warn(SyncEngine.getTool('BULK_FOLDER').TITLE, 'Global', "No pending 'CREATE' actions found.");
        return _App_ok("No pending 'CREATE' actions found.");
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

        return { _rowNumber: rowNum, status: _App_formatStatus('SUCCESS', 'Created: ' + folderNames.join('/')) };

      }, {
        onBatchComplete: function (results) {
          _App_batchPatchResults('BULK_FOLDER', results);
        }
      });

      var finalMsg = "Successfully processed " + stats.processedCount + " folders.";
      if (stats.errorCount > 0) finalMsg += " (" + stats.errorCount + " errors)";
      if (stats.timeLimitReached) finalMsg = "⏳ Time limit reached. " + finalMsg;

      return _App_ok(finalMsg);
    });
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


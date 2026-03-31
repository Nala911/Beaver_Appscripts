/**
 * DRIVE SYNC MANAGER
 * Server-side Logic — Version 5.0 (Plugin Architecture — registers with BeaverEngine)
 */

BeaverEngine.registerTool('DRIVE_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Drive API', test: function() { return typeof Drive !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.DRIVE_SYNC,
    TITLE: '💾 Drive Sync Manager',
    MENU_LABEL: '💾 Google Drive',
    MENU_ENTRYPOINT: 'Drive_showSidebar',
    MENU_ORDER: 90,
    SIDEBAR_HTML: 'DriveFileDetailsSidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [90, 300, 300, 60, 110, 250, 250, 80, 200, 70, 250, 150, 150, 150, 150, 300],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 7,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['Create', 'Update', 'Delete', 'Clone'] },
            { header: 'Item Name', type: 'TEXT' },
            { header: 'Description', type: 'TEXT' },
            { header: 'Starred', type: 'CHECKBOX' },
            { header: 'Type', type: 'DROPDOWN', options: ['Folder', 'Google Doc', 'Google Sheet', 'Google Slide', 'Google Form', 'PDF', 'Image', 'Video', 'Audio', 'Zip', 'Text', 'Code', 'File'] },
            { header: 'Editors', type: 'TEXT' },
            { header: 'Viewers', type: 'TEXT' },
            { header: 'Is Public?', type: 'CHECKBOX' },
            { header: 'Folder Path', type: 'TEXT', italic: true },
            { header: 'Size', type: 'TEXT' },
            { header: 'Owner', type: 'TEXT' },
            { header: 'Mime Type', type: 'TEXT', italic: true },
            { header: 'Last Modified', type: 'TEXT', italic: true },
            { header: 'Item ID', type: 'ID' },
            { header: 'Parent ID', type: 'ID' },
            { header: 'URL', type: 'URL' }
        ]
    }
});

/* ==========================================================================
   CONFIGURATION
   ========================================================================== */

// Column-index aliases — kept for backward-compat; metadata now in BeaverEngine.getTool('DRIVE_SYNC').
var DRIVE_SYNC_COL = {
  ACTION: 0, NAME: 1, DESC: 2, STARRED: 3, TYPE: 4,
  EDITORS: 5, VIEWERS: 6, IS_PUBLIC: 7, PATH: 8,
  SIZE: 9, OWNER: 10, MIME: 11, MODIFIED: 12, ITEM_ID: 13, PARENT_ID: 14, URL: 15
};

// Global Time Limit (Google Apps Script has 6 min limit, we stop at 5.5 min)
var DRIVE_SYNC_START_TIME = 0;
var DRIVE_SYNC_MAX_EXECUTION_TIME = 330 * 1000; // 5.5 minutes

function _DriveSync_checkTimeLimit() {
  if (Date.now() - DRIVE_SYNC_START_TIME > DRIVE_SYNC_MAX_EXECUTION_TIME) {
    throw new Error("⏳ Time limit approaching. Operation paused safely.");
  }
}

// --- SIDEBAR & SHEET SETUP ---

/** @deprecated — Use _App_ensureSheetExists('DRIVE_SYNC') instead. */
function _DriveSync_ensureSheetExistsAndActivate() {
  return _App_ensureSheetExists('DRIVE_SYNC');
}

/** Opens the Drive Sync sidebar and ensures the sheet exists. */
function Drive_showSidebar() {
  return Logger.run('DRIVE_SYNC', 'Open Sidebar', function () {
    _App_launchTool('DRIVE_SYNC');
  });
}




/* ==========================================================================
   CORE LOGIC
   ========================================================================== */

function Drive_getFolderContent(folderId) {
  return Logger.run('DRIVE_SYNC', 'Get Folder Content', function () {
    try {
      var parentId = folderId || "root";
      var query = "'" + parentId + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
      var folders = [];
      var pageToken = null;

      var currentFolder = null;
      try {
        currentFolder = Drive.Files.get(parentId, { fields: "id, name", supportsAllDrives: true });
      } catch (e) {
        try {
          var drv = Drive.Drives.get(parentId);
          currentFolder = { id: drv.id, name: drv.name };
        } catch (e2) {
          currentFolder = { id: parentId, name: parentId === "root" ? "Root" : "Unknown" };
        }
      }

      do {
        var result = Drive.Files.list({
          q: query,
          fields: "nextPageToken, files(id, name)",
          orderBy: "name",
          pageToken: pageToken,
          includeItemsFromAllDrives: true,
          supportsAllDrives: true
        });
        if (result.files) folders = folders.concat(result.files);
        pageToken = result.nextPageToken;
      } while (pageToken);

      return {
        current: { id: currentFolder.id, name: currentFolder.name },
        children: folders,
        success: true
      };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });
}

function Drive_getDrivesList() {
  return Logger.run('DRIVE_SYNC', 'Get Drives List', function () {
    try {
      var drives = [];
      var pageToken = null;
      do {
        var result = Drive.Drives.list({
          pageToken: pageToken,
          fields: "nextPageToken, drives(id, name)"
        });
        if (result.drives) drives = drives.concat(result.drives);
        pageToken = result.nextPageToken;
      } while (pageToken);

      drives.sort(function (a, b) {
        return a.name.localeCompare(b.name);
      });

      return { success: true, drives: drives };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });
}

function Drive_getFolderHierarchy() {
  return Logger.run('DRIVE_SYNC', 'Get Folder Hierarchy', function () {
    try {
      var query = "mimeType = 'application/vnd.google-apps.folder' and trashed = false";
      var folders = [];
      var pageToken = null;

      // 1. Fetch all folders
      do {
        var result = _App_callWithBackoff(function () {
          return Drive.Files.list({
            q: query,
            fields: "nextPageToken, files(id, name, parents)",
            orderBy: "name",
            pageToken: pageToken,
            includeItemsFromAllDrives: true,
            supportsAllDrives: true,
            pageSize: 1000
          });
        });
        if (result.files) folders = folders.concat(result.files);
        pageToken = result.nextPageToken;
      } while (pageToken);

      // 2. Fetch all shared drives
      var drives = [];
      var drivePageToken = null;
      do {
        var dResult = Drive.Drives.list({
          pageToken: drivePageToken,
          fields: "nextPageToken, drives(id, name)"
        });
        if (dResult.drives) drives = drives.concat(dResult.drives);
        drivePageToken = dResult.nextPageToken;
      } while (drivePageToken);

      drives.sort(function (a, b) {
        return a.name.localeCompare(b.name);
      });

      var topology = {
        rootDrives: drives,
        dict: {},
        myDriveId: "root"
      };

      // 3. Resolve the actual ID of "My Drive"
      try {
        var actualRoot = Drive.Files.get("root", { fields: "id", supportsAllDrives: true });
        if (actualRoot && actualRoot.id) {
          topology.myDriveId = actualRoot.id;
        }
      } catch (e) { }

      // 4. Group folders by parentId
      for (var i = 0; i < folders.length; i++) {
        var f = folders[i];
        if (f.parents && f.parents.length > 0) {
          var parentId = f.parents[0];
          if (!topology.dict[parentId]) {
            topology.dict[parentId] = [];
          }
          topology.dict[parentId].push({ id: f.id, name: f.name });
        }
      }

      return { success: true, topology: topology };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });
}

function Drive_getPendingStats() {
  return Logger.run('DRIVE_SYNC', 'Get Pending Stats', function () {
    SheetManager.assertActiveSheet('DRIVE_SYNC');
    var stats = SheetManager.getActionStats('DRIVE_SYNC', ['Create', 'Update', 'Delete', 'Clone']);
    stats.total = (stats.Create || 0) + (stats.Update || 0) + (stats.Delete || 0) + (stats.Clone || 0);
    return _App_ok('Pending stats loaded.', {
      creates: stats.Create || 0,
      updates: stats.Update || 0,
      deletes: stats.Delete || 0,
      clones: stats.Clone || 0,
      total: stats.total
    });
  });
}


function Drive_pullFromDrive(targetFolderId, isShallow) {
  return Logger.run('DRIVE_SYNC', 'Pull from Drive', function () {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return "⚠️ System is busy. Please try again.";

    try {
      DRIVE_SYNC_START_TIME = Date.now();
      targetFolderId = targetFolderId || "root";

      var sheet = _DriveSync_ensureSheetExistsAndActivate();

      var maxRows = sheet.getMaxRows();
      var maxCols = sheet.getMaxColumns();
      if (maxRows > 1) {
        sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
      }

      if (sheet.getLastRow() < 1) {
        _DriveSync_initializeHeaders(sheet);
      }

      var allItems = [];
      var folderMap = new Map();

      // Recursive Fetch with Error Guard
      function recursiveFetch(parentId) {
        _DriveSync_checkTimeLimit();
        try {
          var query = "'" + parentId + "' in parents and trashed = false";
          var fields = "files(id, name, description, starred, mimeType, parents, modifiedTime, webViewLink, size, permissions(type, role, emailAddress))";
          var items = _DriveSync_fetchAllItems(query, fields);

          items.forEach(function (item) {
            item._traversalParentId = parentId; // Track which folder we found this in
            allItems.push(item);
            if (item.mimeType === 'application/vnd.google-apps.folder') {
              folderMap.set(item.id, { name: item.name, parentId: parentId });
              if (!isShallow) {
                recursiveFetch(item.id);
              }
            }
          });
        } catch (e) {
          if (e.message && e.message.indexOf("Time limit") !== -1) throw e;
          Logger.warn('DRIVE_SYNC', 'recursiveFetch', "Error fetching folder " + parentId + ": " + e.message);
        }
      }

      var rootObj = { id: targetFolderId, name: "Root", parents: [] };
      try {
        var drv = Drive.Drives.get(targetFolderId);
        if (drv && drv.name) rootObj.name = drv.name;
        else rootObj = Drive.Files.get(targetFolderId, { fields: "id, name, parents", supportsAllDrives: true });
      } catch (e) {
        rootObj = Drive.Files.get(targetFolderId, { fields: "id, name, parents", supportsAllDrives: true });
      }

      var rootParent = (rootObj.parents && rootObj.parents.length > 0) ? rootObj.parents[0] : null;
      folderMap.set(rootObj.id, { name: rootObj.name, parentId: rootParent });

      // Resolve full path of the target folder (start point)
      var targetFolderFullPath = "";
      var rootFolderName = "My Drive"; // Default fallback
      var foundSharedDriveRoot = false;

      try {
        if (targetFolderId !== "root") {
          var driveObj = Drive.Drives.get(targetFolderId);
          if (driveObj && driveObj.name) {
            rootFolderName = driveObj.name;
            foundSharedDriveRoot = true;
          }
        }
      } catch (e) { }

      if (!foundSharedDriveRoot) {
        try {
          var actualRoot = Drive.Files.get("root", { fields: "name", supportsAllDrives: true });
          if (actualRoot && actualRoot.name) rootFolderName = actualRoot.name;
        } catch (e) { Logger.warn('DRIVE_SYNC', 'Path Resolution', "Could not fetch root name, using default."); }
      }

      if (targetFolderId === "root" || foundSharedDriveRoot) {
        targetFolderFullPath = "/" + rootFolderName;
      } else {
        var parts = [];
        var curr = targetFolderId;
        var depth = 0;
        var foundRoot = false;
        while (curr && depth < 50) {
          if (curr === "root") {
            foundRoot = true;
            break;
          }
          try {
            var isDrive = false;
            try {
              var drv2 = Drive.Drives.get(curr);
              if (drv2 && drv2.name) {
                rootFolderName = drv2.name;
                foundRoot = true;
                isDrive = true;
              }
            } catch (e2) { }
            if (isDrive) break;

            var f = Drive.Files.get(curr, { fields: "name, parents", supportsAllDrives: true });
            parts.unshift(f.name);
            curr = (f.parents && f.parents.length) ? f.parents[0] : null;
            depth++;
          } catch (e) { break; }
        }
        targetFolderFullPath = (foundRoot ? "/" + rootFolderName : "") + "/" + parts.join("/");
      }

      var isPartialPull = false;
      try {
        recursiveFetch(targetFolderId);
      } catch (timeoutEx) {
        if (timeoutEx.message && timeoutEx.message.indexOf("Time limit") !== -1) {
          isPartialPull = true;
        } else {
          throw timeoutEx;
        }
      }

      var rows = [];

      var getPath = function (itemId, currentPath) {
        if (!currentPath) currentPath = [];
        var item = folderMap.get(itemId);
        if (!item || itemId === targetFolderId) {
          var relative = currentPath.join("/");
          return targetFolderFullPath + (relative ? "/" + relative : "");
        }
        currentPath.unshift(item.name);
        return getPath(item.parentId, currentPath);
      };

      var headers = BeaverEngine.getTool('DRIVE_SYNC').HEADERS;
      for (var i = 0; i < allItems.length; i++) {
        var item = allItems[i];
        var parentId = item._traversalParentId || ((item.parents && item.parents.length > 0) ? item.parents[0] : "");
        var path = parentId ? getPath(parentId) : targetFolderFullPath;

        var perms = _DriveSync_parsePermissions(item.permissions);

        var row = new Array(headers.length);
        row[DRIVE_SYNC_COL.ACTION] = "";
        row[DRIVE_SYNC_COL.NAME] = item.name;
        row[DRIVE_SYNC_COL.DESC] = item.description || "";
        row[DRIVE_SYNC_COL.STARRED] = item.starred || false;
        row[DRIVE_SYNC_COL.TYPE] = _DriveSync_getFriendlyType(item.mimeType);

        row[DRIVE_SYNC_COL.SIZE] = _DriveSync_formatBytes(item.size);
        row[DRIVE_SYNC_COL.OWNER] = perms.owners.join(", ");
        row[DRIVE_SYNC_COL.EDITORS] = perms.editors.join(", ");
        row[DRIVE_SYNC_COL.VIEWERS] = perms.viewers.join(", ");
        row[DRIVE_SYNC_COL.IS_PUBLIC] = perms.isPublic;

        row[DRIVE_SYNC_COL.PATH] = path;
        row[DRIVE_SYNC_COL.MIME] = item.mimeType;
        row[DRIVE_SYNC_COL.MODIFIED] = item.modifiedTime;
        row[DRIVE_SYNC_COL.ITEM_ID] = item.id;
        row[DRIVE_SYNC_COL.PARENT_ID] = parentId;
        row[DRIVE_SYNC_COL.URL] = item.webViewLink;
        rows.push(row);
      }

      if (rows.length > 0) {
        var rowParams = { start: 2, total: rows.length };
        var range = sheet.getRange(rowParams.start, 1, rowParams.total, rows[0].length);
        range.setValues(rows);

        _App_applyBodyFormatting(sheet, rows.length, BeaverEngine.getTool('DRIVE_SYNC').FORMAT_CONFIG);

        var msg = "Successfully pulled " + rows.length + " items from " + (targetFolderId === "root" ? "Root Drive" : rootObj.name) + ".";
        if (isPartialPull) {
          msg = "⚠️ Partial Pull: " + msg + " (Execution Time Limit Reached. Run again to continue)";
        }
        Logger.info('DRIVE_SYNC', 'Pull Complete', msg);
        return msg;
      } else {
        return "Target folder is empty.";
      }

    } finally {
      lock.releaseLock();
    }
  });
}

function Drive_runPushSequence() {
  return Logger.run('DRIVE_SYNC', 'Push Sequence', function () {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return ["⚠️ System is busy. Please try again."];

    var logs = [];
    function log(msg) { logs.push("[" + new Date().toLocaleTimeString() + "] " + msg); }

    try {
      DRIVE_SYNC_START_TIME = Date.now();
      log("Starting Push Sequence...");

      var sheet = _App_assertActiveSheet(SHEET_NAMES.DRIVE_SYNC);
      _DriveSync_validateHeaders(sheet);

      var data = sheet.getDataRange().getValues();
      var pendingRows = [];

      for (var i = 1; i < data.length; i++) {
        if (data[i][DRIVE_SYNC_COL.ACTION] !== "") {
          pendingRows.push({ rowData: data[i], rowIndex: i + 1 });
        }
      }

      log("Found " + pendingRows.length + " pending actions.");

      pendingRows.sort(function (a, b) {
        var score = function (r) {
          var act = r.rowData[DRIVE_SYNC_COL.ACTION];
          var type = r.rowData[DRIVE_SYNC_COL.TYPE];
          if (act === 'Create' && type === 'Folder') return 1;
          if (act === 'Clone') return 2;
          if (act === 'Create') return 3;
          if (act === 'Update') return 4;
          return 5;
        };
        return score(a) - score(b);
      });

      var fullSheetData = sheet.getDataRange().getValues();
      var isPartialPush = false;

      for (var k = 0; k < pendingRows.length; k++) {
        try {
          _DriveSync_checkTimeLimit();
        } catch (timeoutEx) {
          log(timeoutEx.message);
          isPartialPush = true;
          break;
        }

        var item = pendingRows[k];
        var rowNum = item.rowIndex; 
        var arrayIndex = rowNum - 1; 

        try {
          var statusMsg = "";
          var action = item.rowData[DRIVE_SYNC_COL.ACTION];

          if (!action) continue;

          var resultValues = {};

          if (action === 'Create') statusMsg = _DriveSync_handleCreate(item.rowData, resultValues);
          else if (action === 'Clone') statusMsg = _DriveSync_handleClone(item.rowData, resultValues);
          else if (action === 'Update') statusMsg = _DriveSync_handleUpdate(item.rowData);
          else if (action === 'Delete') statusMsg = _DriveSync_handleDelete(item.rowData);

          fullSheetData[arrayIndex][DRIVE_SYNC_COL.ACTION] = "";

          if (resultValues.id) fullSheetData[arrayIndex][DRIVE_SYNC_COL.ITEM_ID] = resultValues.id;
          if (resultValues.url) fullSheetData[arrayIndex][DRIVE_SYNC_COL.URL] = resultValues.url;
          if (resultValues.mime) fullSheetData[arrayIndex][DRIVE_SYNC_COL.MIME] = resultValues.mime;
          if (resultValues.size) fullSheetData[arrayIndex][DRIVE_SYNC_COL.SIZE] = resultValues.size;

          Logger.info(BeaverEngine.getTool('DRIVE_SYNC').TITLE, 'Row ' + rowNum + ' (' + item.rowData[DRIVE_SYNC_COL.NAME] + ')', '✅ ' + statusMsg);
          log("Row " + rowNum + ": " + statusMsg);

        } catch (e) {
          Logger.error(BeaverEngine.getTool('DRIVE_SYNC').TITLE, 'Row ' + rowNum + ' (' + item.rowData[DRIVE_SYNC_COL.NAME] + ')', e);
          log("Row " + rowNum + " Error: " + e.message);
        }
      }

      if (pendingRows.length > 0) {
        sheet.getRange(1, 1, fullSheetData.length, fullSheetData[0].length).setValues(fullSheetData);
      }

      if (isPartialPush) {
        log("Sequence Paused: 5.5-minute time limit approached. Please run again to process remaining items.");
      } else {
        log("Sequence Complete.");
      }
      return logs;

    } finally {
      lock.releaseLock();
    }
  });
}

function Drive_setupSheet(sheet) {
  return Logger.run('DRIVE_SYNC', 'Setup Sheet', function () {
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      sheet.clear();
      const maxRows = sheet.getMaxRows();
      const maxCols = sheet.getMaxColumns();
      if (maxRows > 0 && maxCols > 0) sheet.getRange(1, 1, maxRows, maxCols).clearDataValidations();
    }

    _DriveSync_initializeHeaders(sheet);
    
    var cfg = BeaverEngine.getTool('DRIVE_SYNC');
    if (cfg.COL_WIDTHS) {
      cfg.COL_WIDTHS.forEach(function (w, i) {
        if (w !== null && w !== undefined) sheet.setColumnWidth(i + 1, w);
      });
    }
    // Schema-driven validation now handles this within _App_applyBodyFormatting
    
    return _App_ok("Sheet has been reset successfully.");
  });
}

/* ==========================================================================
   PRIVATE HELPER FUNCTIONS
   ========================================================================== */

function _DriveSync_validateHeaders(sheet) {
  var headers = BeaverEngine.getTool('DRIVE_SYNC').HEADERS;
  var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (currentHeaders[0] !== headers[0] || currentHeaders[DRIVE_SYNC_COL.ITEM_ID] !== headers[DRIVE_SYNC_COL.ITEM_ID]) {
    // Auto-fix corrupted headers instead of blocking the operation
    console.warn("Sheet headers appear corrupted. Auto-resetting headers...");
    _DriveSync_initializeHeaders(sheet);
    var cfg = BeaverEngine.getTool('DRIVE_SYNC');
    // Schema-driven validation handles this on sheet load.
  }
}

function _DriveSync_fetchAllItems(query, fields) {
  var items = [];
  var pageToken = null;
  do {
    var result = _App_callWithBackoff(function () {
      return Drive.Files.list({
        q: query,
        fields: "nextPageToken, " + fields,
        pageToken: pageToken,
        pageSize: 1000,
        includeItemsFromAllDrives: true,
        supportsAllDrives: true
      });
    });
    if (result.files) items = items.concat(result.files);
    pageToken = result.nextPageToken;
  } while (pageToken);
  return items;
}



function _DriveSync_formatBytes(bytes) {
  if (!bytes || bytes == 0) return "-";
  var k = 1024;
  var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  var i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function _DriveSync_parsePermissions(permissions) {
  var res = { owners: [], editors: [], viewers: [], isPublic: false };
  if (!permissions) return res;

  permissions.forEach(function (p) {
    if (p.type === 'anyone') res.isPublic = true;
    if (p.emailAddress) {
      if (p.role === 'owner') res.owners.push(p.emailAddress);
      else if (p.role === 'writer' || p.role === 'fileOrganizer') res.editors.push(p.emailAddress);
      else if (p.role === 'reader') res.viewers.push(p.emailAddress);
    }
  });
  return res;
}

function _DriveSync_parseEmailList(str) {
  if (!str) return [];
  return str.toString().split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(function (s) { return s !== ""; });
}

function _DriveSync_getFriendlyType(mimeType) {
  if (!mimeType) return 'File';
  if (mimeType === 'application/vnd.google-apps.folder') return 'Folder';
  if (mimeType === 'application/vnd.google-apps.spreadsheet') return 'Google Sheet';
  if (mimeType === 'application/vnd.google-apps.document') return 'Google Doc';
  if (mimeType === 'application/vnd.google-apps.presentation') return 'Google Slide';
  if (mimeType === 'application/vnd.google-apps.form') return 'Google Form';
  if (mimeType === 'application/pdf') return 'PDF';
  return 'File';
}

function _DriveSync_getMimeTypeFromFriendly(friendlyType) {
  switch (friendlyType) {
    case 'Folder': return 'application/vnd.google-apps.folder';
    case 'Google Sheet': return 'application/vnd.google-apps.spreadsheet';
    case 'Google Doc': return 'application/vnd.google-apps.document';
    case 'Google Slide': return 'application/vnd.google-apps.presentation';
    case 'Google Form': return 'application/vnd.google-apps.form';
    case 'PDF': return 'application/pdf';
    default: return 'application/vnd.google-apps.folder';
  }
}

function _DriveSync_initializeHeaders(sheet) {
  var headers = BeaverEngine.getTool('DRIVE_SYNC').HEADERS;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight(SHEET_THEME.LAYOUT.HEADER_WEIGHT)
    .setBackground(SHEET_THEME.HEADER)
    .setFontColor(SHEET_THEME.TEXT)
    .setBorder(true, true, true, true, true, true, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE);

  sheet.setFrozenRows(1);
}

// Stage 1: Data validations removed, using registry instead.

// --- Handlers ---

function _DriveSync_handleCreate(row, res) {
  var name = row[DRIVE_SYNC_COL.NAME];
  if (!name) throw new Error("Name is required");

  var desc = row[DRIVE_SYNC_COL.DESC];
  var starred = row[DRIVE_SYNC_COL.STARRED] === true;
  var friendlyType = row[DRIVE_SYNC_COL.TYPE];

  var parentId = row[DRIVE_SYNC_COL.PARENT_ID];
  var pathStr = row[DRIVE_SYNC_COL.PATH];

  // Priority: Path > ParentID > Root
  if (pathStr && pathStr.trim() !== "") {
    try {
      parentId = _DriveSync_resolveFolderIdFromPath(pathStr);
    } catch (e) {
      throw new Error("Path resolution failed: " + e.message);
    }
  } else if (!parentId) {
    parentId = DriveApp.getRootFolder().getId();
  }

  var mimeType = _DriveSync_getMimeTypeFromFriendly(friendlyType);

  var resource = { name: name, description: desc, starred: starred, parents: [parentId], mimeType: mimeType };
  var file = _App_callWithBackoff(function () { return Drive.Files.create(resource, null, { fields: 'id, webViewLink, mimeType', supportsAllDrives: true }); });

  res.id = file.id;
  res.url = file.webViewLink;
  res.mime = file.mimeType;

  return "Created (" + (friendlyType || 'Folder') + ")";
}

function _DriveSync_resolveFolderIdFromPath(pathString) {
  if (!pathString || pathString === "/" || pathString.trim() === "") return "root";

  // Normalize path: Remove leading/trailing slashes and split
  var parts = pathString.split("/").filter(function (p) { return p.trim() !== ""; });

  var possibleSharedDrive = null;
  if (parts.length > 0) {
    var first = parts[0].toLowerCase();
    if (first === "my drive" || first === "drive") {
      parts.shift();
    } else {
      try {
        var drivesResult = Drive.Drives.list({ fields: "drives(id, name)" });
        if (drivesResult && drivesResult.drives) {
          for (var d = 0; d < drivesResult.drives.length; d++) {
            if (drivesResult.drives[d].name.toLowerCase() === first) {
              possibleSharedDrive = drivesResult.drives[d];
              parts.shift();
              break;
            }
          }
        }
      } catch (e) { }
    }
  }

  var currentId = possibleSharedDrive ? possibleSharedDrive.id : "root";
  var resolvedSoFar = [];
  if (possibleSharedDrive) resolvedSoFar.push(possibleSharedDrive.name);

  for (var i = 0; i < parts.length; i++) {
    var folderName = parts[i];

    // Search for existing folder in current parent
    var query = "'" + currentId + "' in parents and name = '" + folderName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
    var folders = [];
    try {
      var result = Drive.Files.list({ q: query, fields: "files(id, name)", pageSize: 1, includeItemsFromAllDrives: true, supportsAllDrives: true });
      if (result.files && result.files.length > 0) {
        folders = result.files;
      }
    } catch (e) {
      throw new Error("Drive API error while searching for '" + folderName + "': " + e.message);
    }

    if (folders.length > 0) {
      currentId = folders[0].id;
      resolvedSoFar.push(folderName);
    } else {
      // Folder not found — inform the user instead of auto-creating
      var resolvedPath = resolvedSoFar.length > 0 ? resolvedSoFar.join("/") : "(root)";
      throw new Error(
        "Folder not found: '" + folderName + "' does not exist. " +
        "Resolved up to: " + resolvedPath + ". " +
        "Remaining path: " + parts.slice(i).join("/") + ". " +
        "Please create the folder first or correct the path."
      );
    }
  }

  return currentId;
}

function _DriveSync_handleClone(row, res) {
  var fileId = row[DRIVE_SYNC_COL.ITEM_ID];
  var nameInCell = row[DRIVE_SYNC_COL.NAME];
  var parentId = row[DRIVE_SYNC_COL.PARENT_ID];

  if (!fileId) throw new Error("Cannot Clone: Source Item ID is missing.");
  if (!nameInCell) throw new Error("Cannot Clone: Target Name is missing.");

  var resource = { name: nameInCell, parents: [parentId] };

  var file = _App_callWithBackoff(function () {
    return Drive.Files.copy(resource, fileId, { fields: 'id, webViewLink, mimeType, size', supportsAllDrives: true });
  });

  res.id = file.id;
  res.url = file.webViewLink;
  res.mime = file.mimeType;
  res.size = _DriveSync_formatBytes(file.size);

  return "Cloned";
}

function _DriveSync_handleUpdate(row) {
  var fileId = row[DRIVE_SYNC_COL.ITEM_ID];
  if (!fileId) throw new Error("Cannot Update: Item ID is missing.");

  var newName = row[DRIVE_SYNC_COL.NAME];
  var newDesc = row[DRIVE_SYNC_COL.DESC];
  var newStarred = row[DRIVE_SYNC_COL.STARRED];
  var newParentId = row[DRIVE_SYNC_COL.PARENT_ID];
  var pathStr = row[DRIVE_SYNC_COL.PATH];
  if (pathStr && pathStr.trim() !== "") {
    newParentId = _DriveSync_resolveFolderIdFromPath(pathStr);
  }

  var currentFile = _App_callWithBackoff(function () {
    return Drive.Files.get(fileId, { fields: 'parents, name, description, starred, permissions(id, role, emailAddress, type)', supportsAllDrives: true });
  });

  var changes = [];
  var resource = {};
  if (newName && newName !== currentFile.name) resource.name = newName;
  if (newDesc !== (currentFile.description || "")) resource.description = newDesc;
  if (newStarred !== (currentFile.starred || false)) resource.starred = newStarred;

  var currentParentId = (currentFile.parents && currentFile.parents.length) ? currentFile.parents[0] : null;
  var optionalArgs = {};
  var isMove = false;

  if (newParentId && currentParentId && newParentId !== currentParentId) {
    optionalArgs.addParents = newParentId;
    optionalArgs.removeParents = currentParentId;
    isMove = true;
  }
  optionalArgs.supportsAllDrives = true;

  if (Object.keys(resource).length > 0 || isMove) {
    _App_callWithBackoff(function () { Drive.Files.update(resource, fileId, null, optionalArgs); });
    if (Object.keys(resource).length > 0) changes.push("Properties");
    if (isMove) changes.push("Moved");
  }

  // Permissions
  var newEditors = _DriveSync_parseEmailList(row[DRIVE_SYNC_COL.EDITORS]);
  var newViewers = _DriveSync_parseEmailList(row[DRIVE_SYNC_COL.VIEWERS]);
  var targetIsPublic = row[DRIVE_SYNC_COL.IS_PUBLIC] === true;

  var currentEmailPerms = {};
  var publicPermId = null;

  if (currentFile.permissions) {
    currentFile.permissions.forEach(function (p) {
      if (p.type === 'anyone') publicPermId = p.id;
      else if (p.emailAddress) currentEmailPerms[p.emailAddress.toLowerCase()] = p;
    });
  }

  if (targetIsPublic && !publicPermId) {
    _App_callWithBackoff(function () { Drive.Permissions.create({ role: 'reader', type: 'anyone' }, fileId, { supportsAllDrives: true }); });
    changes.push("Made Public");
  } else if (!targetIsPublic && publicPermId) {
    _App_callWithBackoff(function () { Drive.Permissions.remove(fileId, publicPermId, { supportsAllDrives: true }); });
    changes.push("Made Private");
  }

  var permChanges = false;
  Object.keys(currentEmailPerms).forEach(function (email) {
    var p = currentEmailPerms[email];
    if (p.role === 'owner' || p.role === 'organizer') return;

    var shouldBeEditor = newEditors.indexOf(email) !== -1;
    var shouldBeViewer = newViewers.indexOf(email) !== -1;

    if (!shouldBeEditor && !shouldBeViewer) {
      _App_callWithBackoff(function () { Drive.Permissions.remove(fileId, p.id, { supportsAllDrives: true }); });
      permChanges = true;
    } else if (shouldBeEditor && p.role !== 'writer' && p.role !== 'fileOrganizer') {
      _App_callWithBackoff(function () { Drive.Permissions.update({ role: 'writer' }, fileId, p.id, { supportsAllDrives: true }); });
      permChanges = true;
    } else if (shouldBeViewer && p.role !== 'reader') {
      _App_callWithBackoff(function () { Drive.Permissions.update({ role: 'reader' }, fileId, p.id, { supportsAllDrives: true }); });
      permChanges = true;
    }
  });

  var allTargetEmails = newEditors.concat(newViewers);
  allTargetEmails.forEach(function (email) {
    if (currentEmailPerms[email]) return;
    var role = newEditors.indexOf(email) !== -1 ? 'writer' : 'reader';
    _App_callWithBackoff(function () {
      Drive.Permissions.create({ role: role, type: 'user', emailAddress: email }, fileId, { sendNotificationEmails: false, supportsAllDrives: true });
    });
    permChanges = true;
  });

  if (permChanges) changes.push("Permissions");

  return changes.length > 0 ? "Updated: " + changes.join(", ") : "No Changes Needed";
}

function _DriveSync_handleDelete(row) {
  var fileId = row[DRIVE_SYNC_COL.ITEM_ID];
  if (!fileId) throw new Error("Cannot Delete: Item ID is missing.");
  _App_callWithBackoff(function () { Drive.Files.update({ trashed: true }, fileId, null, { supportsAllDrives: true }); });
  return "Deleted (Trashed)";
}

function Drive_fillActivePath(folderId, pathString) {
  return Logger.run('DRIVE_SYNC', 'Fill Active Path', function () {
    var sheet = _App_assertActiveSheet(SHEET_NAMES.DRIVE_SYNC);

    var cell = sheet.getActiveCell();
    var row = cell.getRow();

    if (row < 2) throw new Error("Please select a row in the data area (Row 2 or below).");

    // Col.PATH is 10 (index), so column is 11
    // Col.PARENT_ID is 14 (index), so column is 15

    sheet.getRange(row, DRIVE_SYNC_COL.PATH + 1).setValue(pathString);
    sheet.getRange(row, DRIVE_SYNC_COL.PARENT_ID + 1).setValue(folderId);

    return "Updated Row " + row + " with path: " + pathString;
  });
}


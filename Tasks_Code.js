/**
 * Google Tasks <-> Sheets Two-Way Sync
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('TASKS', {
    REQUIRED_SERVICES: [ { name: 'Tasks API', test: function() { return typeof Tasks !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.TASKS,
    TITLE: '✅ Task Manager',
    MENU_LABEL: '✅ Google Tasks',
    MENU_ENTRYPOINT: 'Tasks_showSidebar',
    MENU_ORDER: 60,
    SIDEBAR_HTML: 'Tasks_Sidebar',
    SIDEBAR_WIDTH: 300,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [120, 150, 250, 320, 120, 120, null, null, null, null],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 4,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION' },
            { header: 'List Name', type: 'TEXT' },
            { header: 'Title', type: 'TEXT' },
            { header: 'Notes', type: 'TEXT' },
            { header: 'Due Date', type: 'DATE' },
            { header: 'Task Status', type: 'TEXT' },
            { header: 'Task ID', type: 'ID' },
            { header: 'Parent ID', type: 'ID' },
            { header: 'Version Token', type: 'ID' },
            { header: 'Last Sync', type: 'TEXT' }
        ]
    }
});

// ==========================================
// 1. CONFIGURATION & CONSTANTS
// ==========================================

function _TaskSync_getConfig() {
  return SyncEngine.getTool('TASKS');
}

// ==========================================
// 2. UTILITY HELPERS
// ==========================================

function _TaskSync_getColumnMap(sheet) {
  return SheetManager.getSheetHeaderMap(sheet);
}

function _TaskSync_colLetter(idx) {
  return String.fromCharCode(65 + idx);
}

/**
 * @deprecated — Use _App_throttle(tracker, calls) from 03_Core_Utils.js instead.
 * Kept as a thin proxy for backward compatibility.
 */
function _TaskSync_throttle(counter, calls) {
  _App_throttle(counter, calls);
}

function _TaskSync_getTargetSheet() {
  try {
    return SheetManager.getSheet('TASKS');
  } catch (e) {
    throw new Error('"' + _TaskSync_getConfig().SHEET_NAME + '" sheet not found. Please pull tasks first to create it.');
  }
}

// ==========================================
// 3. SIDEBAR HANDLERS
// ==========================================

/** @deprecated — Use _App_ensureSheetExists('TASKS') instead. */
function _TaskSync_ensureSheetExistsAndActivate() {
  return SheetManager.ensureSheet('TASKS');
}

/** Opens the Tasks sidebar and ensures the sheet exists. */
function Tasks_showSidebar() {
  return Logger.run('TASKS', 'Open Sidebar', function () {
    _App_launchTool('TASKS');
  });
}

function Tasks_getTaskLists() {
  return Logger.run('TASKS', 'Get Task Lists', function () {
    try {
      var items = Tasks.Tasklists.list().items || [];
      return _App_ok('Task lists loaded.', {
        lists: items.map(function (i) {
          return { id: i.id, title: i.title };
        })
      });
    } catch (e) {
      Logger.error(SyncEngine.getTool('TASKS').TITLE, 'Get Task Lists', e);
      throw new Error('Failed to fetch lists: ' + e.message);
    }
  });
}

function Tasks_pullRPC(selectedListIds, includeCompleted) {
  return Logger.run('TASKS', 'Pull Tasks', function () {
    _TaskSync_pullTasks(selectedListIds, includeCompleted);
    return _App_ok('Tasks pulled successfully!');
  });
}

function Tasks_pushRPC() {
  return Logger.run('TASKS', 'Push Changes', function () {
    _TaskSync_pushTasks();
    return _App_ok('Changes pushed successfully!');
  });
}

// ==========================================
// 4. PULL LOGIC
// ==========================================

function _TaskSync_pullTasks(selectedListIds, includeCompleted) {
  var sheet = _TaskSync_ensureSheetExistsAndActivate();

  if (sheet.getLastRow() > 1 && SheetManager.hasPendingActions('TASKS')) {
    throw new Error('Unsaved actions detected! Push your changes first or clear the Action column manually.');
  }

  var allListTasks = _TaskSync_fetchSelectedTasks(selectedListIds, includeCompleted);
  var rows = _TaskSync_transformForSheet(allListTasks);
  _TaskSync_renderSheet(sheet, rows);
}

function _TaskSync_fetchSelectedTasks(selectedIds, includeCompleted) {
  var taskLists = Tasks.Tasklists.list().items || [];

  if (selectedIds && selectedIds.length > 0) {
    var idSet = new Set(selectedIds);
    taskLists = taskLists.filter(function (l) { return idSet.has(l.id); });
  }

  var result = [];

  taskLists.forEach(function (list) {
    var options = { showHidden: true };
    if (includeCompleted) {
      options.showCompleted = true;
    }

    var rawTasks = [];
    var pageToken = null;

    do {
      options.pageToken = pageToken;
      var response = Tasks.Tasks.list(list.id, options);
      if (response.items) {
        rawTasks = rawTasks.concat(response.items);
      }
      pageToken = response.nextPageToken;
    } while (pageToken);

    if (!includeCompleted) {
      rawTasks = rawTasks.filter(function (t) { return t.status !== 'completed'; });
    }

    result.push({ listName: list.title, tasks: rawTasks });
  });

  return result;
}

function _TaskSync_transformForSheet(allListTasks) {
  var allRows = [];
  allListTasks.forEach(function (entry) {
    var sortedRows = _TaskSync_processHierarchy(entry.tasks, entry.listName);
    allRows = allRows.concat(sortedRows);
  });
  return allRows;
}

function _TaskSync_processHierarchy(tasks, listName) {
  var rows = [];
  var childrenMap = {};
  var roots = [];
  var taskIds = new Set(tasks.map(function (t) { return t.id; }));

  tasks.forEach(function (task) {
    if (task.parent && taskIds.has(task.parent)) {
      if (!childrenMap[task.parent]) childrenMap[task.parent] = [];
      childrenMap[task.parent].push(task);
    } else {
      roots.push(task);
    }
  });

  var sorter = function (a, b) {
    var dateA = a.due ? new Date(a.due).getTime() : 9999999999999;
    var dateB = b.due ? new Date(b.due).getTime() : 9999999999999;
    if (dateA !== dateB) return dateA - dateB;
    return (a.title || '').localeCompare(b.title || '');
  };

  function buildRows(taskList, depth) {
    taskList.sort(sorter);
    taskList.forEach(function (task) {
      var indent = depth > 0 ? '   '.repeat(depth) + '↳ ' : '';
      var displayTitle = indent + (task.title || '(No Title)');

      rows.push([
        '',                               // Action
        listName,                         // List Name
        displayTitle,                     // Title
        task.notes || '',                 // Notes
        task.due ? new Date(task.due) : '', // Due Date
        task.status,                      // Task Status
        task.id,                          // Task ID
        task.parent || '',                // Parent ID
        task.etag,                        // ETag
        new Date()                        // Last Sync
      ]);

      if (childrenMap[task.id]) {
        buildRows(childrenMap[task.id], depth + 1);
      }
    });
  }

  buildRows(roots, 0);
  return rows;
}

function _TaskSync_renderSheet(sheet, rows) {
  SheetManager.overwriteRows('TASKS', rows, {
    totalCols: _TaskSync_getConfig().HEADERS.length,
    formatConfig: _TaskSync_getConfig().FORMAT_CONFIG
  });

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearNote();
  }
}

// ==========================================
// 5. PUSH LOGIC
// ==========================================

function _TaskSync_pushTasks() {
  var sheet = _TaskSync_getTargetSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var map = _TaskSync_getColumnMap(sheet);
  if (map['ACTION'] === undefined || map['TASK ID'] === undefined) {
    throw new Error('Critical columns (Action, Task ID) missing. Please standardise headers.');
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var values = dataRange.getValues();

  var taskLists = Tasks.Tasklists.list().items || [];
  var listNameToId = {};
  taskLists.forEach(function (l) { listNameToId[l.title] = l.id; });

  var taskIdToListId = {};
  if (map['LIST NAME'] !== undefined && map['TASK ID'] !== undefined) {
    values.forEach(function (row) {
      var tid = row[map['TASK ID']];
      var lname = row[map['LIST NAME']];
      if (tid && lname && listNameToId[lname]) {
        taskIdToListId[tid] = listNameToId[lname];
      }
    });
  }

  var counter = { apiCalls: 0 };
  var processedCount = 0;
  var taskCache = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowIndex = i + 2;
    var action = row[map['ACTION']];

    if (!action) continue;

    var logMsg = '';
    var isError = false;
    var errorObj = null;

    try {
      var listName = row[map['LIST NAME']];
      var targetListId = listNameToId[listName];
      var taskId = row[map['TASK ID']];

      var rawTitle = String(row[map['TITLE']] || '');
      var cleanTitle = rawTitle.replace(/^(\s*↳\s*)+/, '').trim();

      var taskResource = {
        title: cleanTitle,
        notes: row[map['NOTES']],
        status: row[map['TASK STATUS']]
      };

      if (map['DUE DATE'] !== undefined) {
        if (row[map['DUE DATE']]) {
          var d = new Date(row[map['DUE DATE']]);
          if (!isNaN(d.getTime())) {
            taskResource.due = d.toISOString();
          }
        } else {
          taskResource.due = null;
        }
      }

      if (!targetListId && listName && action !== 'Remove') {
        var newList = Tasks.Tasklists.insert({ title: listName });
        targetListId = newList.id;
        listNameToId[listName] = targetListId;
        _TaskSync_throttle(counter, 1);
      }
      if (!targetListId) throw new Error('List not found: ' + listName);

      if (action === 'Create') {
        if (!cleanTitle) throw new Error('Title is empty');
        var parentId = map['PARENT ID'] !== undefined ? row[map['PARENT ID']] : null;
        var opt = parentId ? { parent: parentId } : {};

        var inserted = Tasks.Tasks.insert(taskResource, targetListId, opt);
        _TaskSync_throttle(counter, 1);

        values[i][map['TASK ID']] = inserted.id;
        values[i][map['VERSION TOKEN']] = inserted.etag;
        values[i][map['LAST SYNC']] = new Date();
        logMsg = 'Created successfully';
      }
      else if (action === 'Update') {
        if (!taskId) throw new Error('Missing Task ID');

        var storedEtag = map['VERSION TOKEN'] !== undefined ? row[map['VERSION TOKEN']] : null;
        if (storedEtag) {
          try {
            var live = Tasks.Tasks.get(targetListId, taskId);
            _TaskSync_throttle(counter, 1);
            if (live.etag !== storedEtag) {
              throw new Error('Conflict: Remote task changed. Pull to refresh.');
            }
          } catch (e) {
            if (e.message && e.message.includes('Not Found')) throw new Error('Task not found on server.');
            throw e;
          }
        }

        var updated = Tasks.Tasks.patch(taskResource, targetListId, taskId);
        _TaskSync_throttle(counter, 1);

        values[i][map['VERSION TOKEN']] = updated.etag;
        values[i][map['LAST SYNC']] = new Date();
        logMsg = 'Updated successfully';
      }
      else if (action === 'Remove') {
        if (!taskId) throw new Error('Missing Task ID');
        Tasks.Tasks.remove(targetListId, taskId);
        _TaskSync_throttle(counter, 1);
        logMsg = 'Removed successfully';
      }
      else if (action === 'Move') {
        if (!taskId) throw new Error('Missing Task ID');
        var sourceListId = taskIdToListId[taskId];
        if (!sourceListId) throw new Error('Source list unknown. Pull first.');

        if (sourceListId === targetListId) {
          var parentId = map['PARENT ID'] !== undefined ? row[map['PARENT ID']] : null;
          var moveParams = {};
          if (parentId) {
            moveParams.parent = parentId;
          }
          Tasks.Tasks.patch(taskResource, targetListId, taskId);
          Tasks.Tasks.move(targetListId, taskId, moveParams);
          _TaskSync_throttle(counter, 2);
          logMsg = 'Moved (In-List) successfully';
        } else {
          var oldToNewMap = {}; // Capture all ID migrations from the recursive move
          var insertedId = _TaskSync_deepMoveTask(sourceListId, taskId, targetListId, null, counter, taskCache, oldToNewMap);

          if (!insertedId) {
            logMsg = 'Skipped (Parent moved)';
            if (map['STATUS'] !== undefined) {
              values[i][map['STATUS']] = logMsg;
            }
            if (map['ACTION'] !== undefined) {
              values[i][map['ACTION']] = '';
            }
            continue; // Prevent querying Tasks API with null
          }

          // ** Desync Fix **: Update all downstream rows in the array in-memory so they don't break
          for (var j = i + 1; j < values.length; j++) {
            if (map['TASK ID'] !== undefined && oldToNewMap[values[j][map['TASK ID']]]) {
              values[j][map['TASK ID']] = oldToNewMap[values[j][map['TASK ID']]];
            }
            if (map['PARENT ID'] !== undefined && oldToNewMap[values[j][map['PARENT ID']]]) {
              values[j][map['PARENT ID']] = oldToNewMap[values[j][map['PARENT ID']]];
            }
          }

          var inserted = Tasks.Tasks.get(targetListId, insertedId);
          _TaskSync_throttle(counter, 1);

          values[i][map['TASK ID']] = inserted.id;
          values[i][map['VERSION TOKEN']] = inserted.etag;
          values[i][map['LAST SYNC']] = new Date();
          logMsg = 'Moved to ' + listName;
        }
      }

      processedCount++;

    } catch (e) {
      isError = true;
      errorObj = e;
      logMsg = 'ERROR: ' + e.message;
      console.error('Row ' + rowIndex + ': ' + e.message);
    }

    var reference = 'Row ' + rowIndex + ' (' + (row[map['TITLE']] || 'Unknown') + ')';
    if (isError) {
      Logger.error(SyncEngine.getTool('TASKS').TITLE, reference, errorObj || logMsg);
    } else if (logMsg) {
      Logger.info(SyncEngine.getTool('TASKS').TITLE, reference, logMsg);
    }

    if (!isError && map['ACTION'] !== undefined) {
      values[i][map['ACTION']] = '';
    }
  }

  dataRange.setValues(values);
  SpreadsheetApp.flush();
}

function _TaskSync_getCachedTasks(listId, cache, counter) {
  if (cache[listId]) return cache[listId];
  var options = { showHidden: true };
  var rawTasks = [];
  var pageToken = null;
  do {
    options.pageToken = pageToken;
    var response = Tasks.Tasks.list(listId, options);
    _TaskSync_throttle(counter, 1);
    if (response.items) {
      rawTasks = rawTasks.concat(response.items);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  cache[listId] = rawTasks;
  return rawTasks;
}

function _TaskSync_deepMoveTask(sourceListId, taskId, targetListId, newParentId, counter, taskCache, oldToNewMap) {
  try {
    var originalTask = Tasks.Tasks.get(sourceListId, taskId);
    _TaskSync_throttle(counter, 1);

    var copyRes = {
      title: originalTask.title,
      notes: originalTask.notes,
      status: originalTask.status
    };

    if (originalTask.due) {
      copyRes.due = originalTask.due;
    }

    var opt = newParentId ? { parent: newParentId } : {};
    var inserted = Tasks.Tasks.insert(copyRes, targetListId, opt);
    _TaskSync_throttle(counter, 1);

    if (oldToNewMap) {
      oldToNewMap[taskId] = inserted.id; // Record the mapping
    }

    var allSourceTasks = _TaskSync_getCachedTasks(sourceListId, taskCache, counter);
    var children = allSourceTasks.filter(function (t) { return t.parent === taskId; });

    children.forEach(function (child) {
      _TaskSync_deepMoveTask(sourceListId, child.id, targetListId, inserted.id, counter, taskCache, oldToNewMap);
    });

    Tasks.Tasks.remove(sourceListId, taskId);
    _TaskSync_throttle(counter, 1);

    return inserted.id;
  } catch (e) {
    if (e.message && e.message.includes('Not Found')) {
      // Task doesn't exist, ignore moving inner recursion or duplicate move
      return null;
    }
    throw e;
  }
}

// ==========================================
// 6. FORMATTING
// ==========================================

// Conditional formatting is now handled by _App_applyBodyFormatting via _App_applyConditionalRules

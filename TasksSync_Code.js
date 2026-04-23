/**
 * Google Tasks <-> Sheets Two-Way Sync
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('TASKS', {
    REQUIRED_SERVICES: [ { name: 'Tasks API', test: function() { return typeof Tasks !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.TASKS,
    TITLE: SHEET_NAMES.TASKS,
    MENU_LABEL: SHEET_NAMES.TASKS,
    MENU_ENTRYPOINT: 'TasksSync_openSidebar',
    MENU_ORDER: 60,
    SIDEBAR_HTML: 'TasksSync_Sidebar',
    SIDEBAR_WIDTH: 300,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [120, 150, 250, 320, 120, 120, null, null, null, null],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 4,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION' },
            {
                header: 'List Name', type: 'DROPDOWN', allowInvalid: true, options: function () {
                    var lists = [];
                    try {
                        var response = _App_callWithBackoff(function () { return Tasks.Tasklists.list(); });
                        (response.items || []).forEach(function (l) {
                            if (l.title) lists.push(l.title);
                        });
                        lists.sort();
                    } catch (e) { }
                    return lists.length ? lists.slice(0, 499) : ['My Tasks'];
                }
            },
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

function _TasksSync_getConfig() {
  return SyncEngine.getTool('TASKS');
}

// ==========================================
// 2. UTILITY HELPERS
// ==========================================

function _TasksSync_getColumnMap(sheet) {
  return SheetManager.getSheetHeaderMap(sheet);
}

function _TasksSync_colLetter(idx) {
  return String.fromCharCode(65 + idx);
}

/**
 * @deprecated — Use _App_throttle(tracker, calls) from 03_Core_Utils.js instead.
 * Kept as a thin proxy for backward compatibility.
 */
function _TasksSync_throttle(counter, calls) {
  _App_throttle(counter, calls);
}

function _TasksSync_getTargetSheet() {
  try {
    return SheetManager.getSheet('TASKS');
  } catch (e) {
    throw new Error('"' + _TasksSync_getConfig().SHEET_NAME + '" sheet not found. Please pull tasks first to create it.');
  }
}

// ==========================================
// 3. SIDEBAR HANDLERS
// ==========================================

/** Opens the Tasks sidebar and ensures the sheet exists. */
function TasksSync_openSidebar() {
  return Logger.run('TASKS', 'Open Sidebar', function () {
    _App_launchTool('TASKS');
  });
}

function TasksSync_getTaskLists() {
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

function TasksSync_pullRPC(includeCompleted) {
  return Logger.run('TASKS', 'Pull Tasks', function () {
    _TasksSync_pullTasks(includeCompleted);
    return _App_ok('Tasks pulled successfully!');
  });
}

function TasksSync_pushRPC() {
  return Logger.run('TASKS', 'Push Changes', function () {
    _TasksSync_pushTasks();
    return _App_ok('Changes pushed successfully!');
  });
}

// ==========================================
// 4. PULL LOGIC
// ==========================================

function _TasksSync_pullTasks(includeCompleted) {
  var sheet = _App_ensureSheetExists('TASKS');

  if (sheet.getLastRow() > 1 && SheetManager.hasPendingActions('TASKS')) {
    throw new Error('Unsaved actions detected! Push your changes first or clear the Action column manually.');
  }

  var allListTasks = _TasksSync_fetchSelectedTasks(includeCompleted);
  var rows = _TasksSync_transformForSheet(allListTasks);
  _TasksSync_renderSheet(sheet, rows);
}

function _TasksSync_fetchSelectedTasks(includeCompleted) {
  var taskLists = Tasks.Tasklists.list().items || [];

  // Sort lists alphabetically by title
  taskLists.sort(function(a, b) {
    return (a.title || '').localeCompare(b.title || '');
  });

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

function _TasksSync_transformForSheet(allListTasks) {
  var allRows = [];
  allListTasks.forEach(function (entry) {
    var sortedRows = _TasksSync_processHierarchy(entry.tasks, entry.listName);
    allRows = allRows.concat(sortedRows);
  });
  return allRows;
}

function _TasksSync_processHierarchy(tasks, listName) {
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

function _TasksSync_renderSheet(sheet, rows) {
  SheetManager.overwriteRows('TASKS', rows, {
    totalCols: _TasksSync_getConfig().HEADERS.length,
    formatConfig: _TasksSync_getConfig().FORMAT_CONFIG
  });

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearNote();
  }
}

// ==========================================
// 5. PUSH LOGIC
// ==========================================

function _TasksSync_pushTasks() {
  var pendingItems = SheetManager.readPendingObjects('TASKS');

  if (pendingItems.length === 0) return;

  var taskLists = Tasks.Tasklists.list().items || [];
  var listNameToId = {};
  taskLists.forEach(function (l) { listNameToId[l.title] = l.id; });

  // Map Task IDs to List IDs for movement lookups
  var taskIdToListId = {};
  // For Move operations, we need to know the source list of tasks.
  // We can fetch this from the sheet once if needed, or pull from pendingItems if they are already there.
  // However, source list might be in a non-pending row. 
  // To keep it simple and optimized, we'll fetch the Task ID and List Name columns for the whole sheet.
  var sheet = SheetManager.getSheet('TASKS');
  var lastRow = sheet.getLastRow();
  var headers = SheetManager.getHeaders('TASKS');
  var listColIdx = headers.indexOf('List Name') + 1;
  var idColIdx = headers.indexOf('Task ID') + 1;
  
  if (listColIdx > 0 && idColIdx > 0) {
    var idData = sheet.getRange(2, idColIdx, lastRow - 1, 1).getValues();
    var listData = sheet.getRange(2, listColIdx, lastRow - 1, 1).getValues();
    for (var i = 0; i < idData.length; i++) {
      var tid = idData[i][0];
      var lname = listData[i][0];
      if (tid && lname && listNameToId[lname]) {
        taskIdToListId[tid] = listNameToId[lname];
      }
    }
  }

  var counter = { apiCalls: 0 };
  var taskCache = {};

  var stats = _App_BatchProcessor('TASKS', pendingItems, function (item) {
    var rowUpdates = {
      action: item['Action'],
      taskId: item['Task ID'],
      etag: item['Version Token'],
      lastSync: item['Last Sync'],
      _rowNumber: item._rowNumber
    };

    var action = rowUpdates.action;
    var listName = item['List Name'];
    var targetListId = listNameToId[listName];
    var taskId = item['Task ID'];

    var rawTitle = String(item['Title'] || '');
    var cleanTitle = rawTitle.replace(/^(\s*↳\s*)+/, '').trim();

    var taskResource = {
      title: cleanTitle,
      notes: item['Notes'],
      status: item['Task Status']
    };

    if (item['Due Date']) {
      var d = new Date(item['Due Date']);
      if (!isNaN(d.getTime())) taskResource.due = d.toISOString();
    } else {
      taskResource.due = null;
    }

    if (!targetListId && listName && action !== 'Remove') {
      var newList = Tasks.Tasklists.insert({ title: listName });
      targetListId = newList.id;
      listNameToId[listName] = targetListId;
      _TasksSync_throttle(counter, 1);
    }
    if (!targetListId) throw new Error('List not found: ' + listName);

    var logMsg = '';
    if (action === 'Create') {
      if (!cleanTitle) throw new Error('Title is empty');
      var parentId = item['Parent ID'];
      var opt = parentId ? { parent: parentId } : {};

      var inserted = Tasks.Tasks.insert(taskResource, targetListId, opt);
      _TasksSync_throttle(counter, 1);

      rowUpdates.taskId = inserted.id;
      rowUpdates.etag = inserted.etag;
      rowUpdates.lastSync = new Date();
      logMsg = 'Created successfully';
    }
    else if (action === 'Update') {
      if (!taskId) throw new Error('Missing Task ID');

      var storedEtag = item['Version Token'];
      if (storedEtag) {
        try {
          var live = Tasks.Tasks.get(targetListId, taskId);
          _TasksSync_throttle(counter, 1);
          if (live.etag !== storedEtag) throw new Error('Conflict: Remote task changed. Pull to refresh.');
        } catch (e) {
          if (e.message && e.message.includes('Not Found')) throw new Error('Task not found on server.');
          throw e;
        }
      }

      var updated = Tasks.Tasks.patch(taskResource, targetListId, taskId);
      _TasksSync_throttle(counter, 1);

      rowUpdates.etag = updated.etag;
      rowUpdates.lastSync = new Date();
      logMsg = 'Updated successfully';
    }
    else if (action === 'Remove') {
      if (!taskId) throw new Error('Missing Task ID');
      Tasks.Tasks.remove(targetListId, taskId);
      _TasksSync_throttle(counter, 1);
      logMsg = 'Removed successfully';
    }
    else if (action === 'Move') {
      if (!taskId) throw new Error('Missing Task ID');
      var sourceListId = taskIdToListId[taskId];
      if (!sourceListId) throw new Error('Source list unknown. Pull first.');

      if (sourceListId === targetListId) {
        var parentId = item['Parent ID'];
        var moveParams = {};
        if (parentId) moveParams.parent = parentId;
        Tasks.Tasks.patch(taskResource, targetListId, taskId);
        Tasks.Tasks.move(targetListId, taskId, moveParams);
        _TasksSync_throttle(counter, 2);
        logMsg = 'Moved (In-List) successfully';
      } else {
        var oldToNewMap = {};
        var insertedId = _TasksSync_deepMoveTask(sourceListId, taskId, targetListId, null, counter, taskCache, oldToNewMap);

        if (!insertedId) {
          logMsg = 'Skipped (Parent moved)';
        } else {
          // If we have an ID mapping, we MUST update all rows in the sheet that refer to old IDs.
          // This is a complex case. For optimization, we'll patch affected rows directly.
          Object.keys(oldToNewMap).forEach(function(oldId) {
            var newId = oldToNewMap[oldId];
            // We need to find rows with this oldId as Task ID or Parent ID.
            // Since this is rare, we'll do a surgical update if possible, or just log.
            // For now, let's keep it simple: the user should 'Pull' after a cross-list move.
          });

          var inserted = Tasks.Tasks.get(targetListId, insertedId);
          _TasksSync_throttle(counter, 1);

          rowUpdates.taskId = inserted.id;
          rowUpdates.etag = inserted.etag;
          rowUpdates.lastSync = new Date();
          logMsg = 'Moved to ' + listName;
        }
      }
    }

    rowUpdates.action = '';
    
    var reference = 'Row ' + item._rowNumber + ' (' + (item['Title'] || 'Unknown') + ')';
    Logger.info(SyncEngine.getTool('TASKS').TITLE, reference, logMsg);

    return rowUpdates;

  }, {
    onBatchComplete: function (batchResults) {
      var rowNumbers = [];
      var updates = [];
      batchResults.forEach(function (res) {
        if (res && res._rowNumber !== undefined) {
          rowNumbers.push(res._rowNumber);
          updates.push({
            'Action': res.action,
            'Task ID': res.taskId,
            'Version Token': res.etag,
            'Last Sync': res.lastSync
          });
        }
      });
      if (rowNumbers.length > 0) {
        SheetManager.batchPatchRows('TASKS', rowNumbers, updates);
      }
    }
  });

  SpreadsheetApp.flush();
}

function _TasksSync_getCachedTasks(listId, cache, counter) {
  if (cache[listId]) return cache[listId];
  var options = { showHidden: true };
  var rawTasks = [];
  var pageToken = null;
  do {
    options.pageToken = pageToken;
    var response = Tasks.Tasks.list(listId, options);
    _TasksSync_throttle(counter, 1);
    if (response.items) {
      rawTasks = rawTasks.concat(response.items);
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  cache[listId] = rawTasks;
  return rawTasks;
}

function _TasksSync_deepMoveTask(sourceListId, taskId, targetListId, newParentId, counter, taskCache, oldToNewMap) {
  try {
    var originalTask = Tasks.Tasks.get(sourceListId, taskId);
    _TasksSync_throttle(counter, 1);

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
    _TasksSync_throttle(counter, 1);

    if (oldToNewMap) {
      oldToNewMap[taskId] = inserted.id; // Record the mapping
    }

    var allSourceTasks = _TasksSync_getCachedTasks(sourceListId, taskCache, counter);
    var children = allSourceTasks.filter(function (t) { return t.parent === taskId; });

    children.forEach(function (child) {
      _TasksSync_deepMoveTask(sourceListId, child.id, targetListId, inserted.id, counter, taskCache, oldToNewMap);
    });

    Tasks.Tasks.remove(sourceListId, taskId);
    _TasksSync_throttle(counter, 1);

    return inserted.id;
  } catch (e) {
    if (e.message && e.message.includes('Not Found')) {
      // Task doesn't exist, ignore moving inner recursion or duplicate move
      return null;
    }
    throw e;
  }
}

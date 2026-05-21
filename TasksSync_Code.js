/**
 * Google Tasks Sync
 * Version: 1.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('TASKS_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Tasks API', test: function() { return typeof Tasks !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.TASKS_SYNC,
    TITLE: SHEET_NAMES.TASKS_SYNC,
    MENU_LABEL: SHEET_NAMES.TASKS_SYNC,
    MENU_ENTRYPOINT: 'TasksSync_openSidebar',
    MENU_ORDER: 45,
    SIDEBAR_HTML: 'TasksSync_Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 2,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Task List Name', type: 'DROPDOWN', options: function() {
                try {
                  var lists = _App_callWithBackoff(function () {
                    return Tasks.Tasklists.list().items;
                  });
                  return (lists || []).map(function(t) { return t.title; });
                } catch(e) {
                  return [];
                }
              }
            },
            { header: 'Task Title', type: 'TEXT' },
            { header: 'Description', type: 'TEXT' },
            { header: 'Due Date', type: 'DATE' },
            { header: 'Completed?', type: 'DROPDOWN', options: ['Completed', 'Not Completed'] },
            { header: 'Task ID', type: 'ID' },
            { header: 'Task List ID', type: 'ID' }
        ]
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Tasks sidebar and ensures the sheet exists. */
function TasksSync_openSidebar() {
  return Logger.run('TASKS_SYNC', 'Open Sidebar', function () {
    _App_launchTool('TASKS_SYNC');
  });
}

// --- API FOR SIDEBAR ---

/** Check for unsaved changes before pull */
function TasksSync_checkForUnsavedChanges() {
  return Logger.run('TASKS_SYNC', 'Check Unsaved', function () {
    var hasChanges = false;
    try {
      hasChanges = SheetManager.hasPendingActions('TASKS_SYNC');
    } catch (e) {
      // If sheet doesn't exist yet, there are no changes
    }
    var response = _App_ok('Check complete.', hasChanges);
    response.hasChanges = hasChanges;
    return response;
  });
}

// --- PULL WORKFLOW ---

/** Imports tasks from all task lists into the sheet */
function TasksSync_pullTasks() {
  return Logger.run('TASKS_SYNC', 'Pull Tasks', function () {
    var TARGET_SHEET_NAME = SHEET_NAMES.TASKS_SYNC;
    _App_ensureSheetExists('TASKS_SYNC');

    var allLists = _App_callWithBackoff(function () {
      return Tasks.Tasklists.list().items;
    }) || [];

    var outputObjects = [];

    allLists.forEach(function (list) {
      try {
        var tasksResult = _App_callWithBackoff(function () {
          return Tasks.Tasks.list(list.id, { showCompleted: true, showHidden: true });
        });
        var tasks = tasksResult.items || [];
        tasks.forEach(function (t) {
          var formattedDue = "";
          if (t.due) {
            var d = new Date(t.due);
            if (!isNaN(d.getTime())) {
              formattedDue = Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy");
            }
          }

          outputObjects.push({
            'Action': "",
            'Status': "",
            'Task List Name': list.title,
            'Task Title': t.title || "",
            'Description': t.notes || "",
            'Due Date': formattedDue,
            'Completed?': t.status === 'completed' ? 'Completed' : 'Not Completed',
            'Task ID': t.id,
            'Task List ID': list.id
          });
        });
      } catch (err) {
        Logger.warn('TASKS_SYNC', 'Pull List Error', 'List ' + list.title + ': ' + err.message);
      }
    });

    // Sort by Task List Name, then Task Title
    outputObjects.sort(function (a, b) {
      var nameA = (a['Task List Name'] || "").toLowerCase();
      var nameB = (b['Task List Name'] || "").toLowerCase();
      if (nameA < nameB) return -1;
      if (nameA > nameB) return 1;

      var titleA = (a['Task Title'] || "").toLowerCase();
      var titleB = (b['Task Title'] || "").toLowerCase();
      if (titleA < titleB) return -1;
      if (titleA > titleB) return 1;
      return 0;
    });

    // Populate sheet
    SheetManager.overwriteObjects('TASKS_SYNC', outputObjects);

    var summary = 'Successfully imported ' + outputObjects.length + " tasks into '" + TARGET_SHEET_NAME + "'.";
    return _App_ok(summary);
  });
}

// --- PUSH WORKFLOW ---

/** Commits row-level CREATE, UPDATE, DELETE changes to Google Tasks in bulk */
function TasksSync_pushChanges() {
  return Logger.run('TASKS_SYNC', 'Push Changes', function () {
    var pendingItems = SheetManager.readPendingObjects('TASKS_SYNC');

    if (pendingItems.length === 0) return _App_ok("No pending actions found.");

    var allLists = _App_callWithBackoff(function () {
      return Tasks.Tasklists.list().items;
    }) || [];

    var listMapByName = new Map();
    var listMapById = new Map();
    allLists.forEach(function (l) {
      listMapByName.set(l.title, l.id);
      listMapById.set(l.id, l);
    });

    var stats = _App_BatchProcessor('TASKS_SYNC', pendingItems, function (item) {
      var rowUpdates = {
        action: item['Action'],
        taskId: item['Task ID'] ? String(item['Task ID']) : null,
        listId: item['Task List ID'] ? String(item['Task List ID']) : null,
        status: "",
        _rowNumber: item._rowNumber
      };

      var action = rowUpdates.action.toString().toUpperCase();
      var targetListName = item['Task List Name'];
      var targetListId = listMapByName.get(targetListName);

      var taskData = {
        title: item['Task Title'],
        description: item['Description'],
        dueDate: item['Due Date'],
        completedVal: item['Completed?']
      };

      if (action !== "DELETE" && !taskData.title) throw new Error("⚠️ Data Error: Missing Task Title");

      switch (action) {
        case "CREATE":
          if (!targetListName) throw new Error("⚠️ Data Error: Missing Task List Name");
          if (!targetListId) throw new Error("⚠️ Data Error: Task List '" + targetListName + "' not found");

          var dueStr = null;
          if (taskData.dueDate) {
            var d = new Date(taskData.dueDate);
            if (isNaN(d.getTime())) throw new Error("⚠️ Data Error: Invalid Due Date format");
            d.setUTCHours(0, 0, 0, 0);
            dueStr = d.toISOString();
          }

          var taskResource = {
            title: taskData.title,
            notes: taskData.description || "",
            status: taskData.completedVal === 'Completed' ? 'completed' : 'needsAction'
          };
          if (dueStr) taskResource.due = dueStr;

          var newTask = _App_callWithBackoff(function () {
            return Tasks.Tasks.insert(taskResource, targetListId);
          });

          rowUpdates.taskId = newTask.id;
          rowUpdates.listId = targetListId;
          rowUpdates.status = _App_formatStatus('SUCCESS', "Created");
          rowUpdates.action = "";
          break;

        case "UPDATE":
          if (!rowUpdates.taskId) throw new Error("⚠️ Data Error: Missing Task ID");
          if (!rowUpdates.listId) throw new Error("⚠️ Data Error: Missing Task List ID");

          if (targetListName && !targetListId) throw new Error("⚠️ Data Error: Task List '" + targetListName + "' not found");

          // Identity Check: If target task list name doesn't match current Task List ID, perform MOVE
          if (targetListId && rowUpdates.listId && targetListId !== rowUpdates.listId) {
            rowUpdates = _TasksSync_processMove(rowUpdates, targetListId, taskData);
            break;
          }

          var dueStrUpdate = null;
          if (taskData.dueDate) {
            var dUpdate = new Date(taskData.dueDate);
            if (isNaN(dUpdate.getTime())) throw new Error("⚠️ Data Error: Invalid Due Date format");
            dUpdate.setUTCHours(0, 0, 0, 0);
            dueStrUpdate = dUpdate.toISOString();
          }

          var updateResource = {
            title: taskData.title,
            notes: taskData.description || "",
            status: taskData.completedVal === 'Completed' ? 'completed' : 'needsAction',
            due: dueStrUpdate
          };

          _App_callWithBackoff(function () {
            return Tasks.Tasks.patch(updateResource, rowUpdates.listId, rowUpdates.taskId);
          });

          rowUpdates.status = _App_formatStatus('SUCCESS', "Updated");
          rowUpdates.action = "";
          break;

        case "DELETE":
          if (!rowUpdates.taskId) throw new Error("⚠️ Data Error: Missing Task ID");
          if (!rowUpdates.listId) throw new Error("⚠️ Data Error: Missing Task List ID");

          try {
            _App_callWithBackoff(function () {
              Tasks.Tasks.remove(rowUpdates.listId, rowUpdates.taskId);
            });
            rowUpdates.status = _App_formatStatus('SUCCESS', "Deleted");
            rowUpdates.action = "";
          } catch (e) {
            if (e.message.indexOf('404') !== -1 || e.message.indexOf('not found') !== -1) {
              rowUpdates.status = _App_formatStatus('WARNING', "Already Deleted");
              rowUpdates.action = "";
            } else {
              throw e;
            }
          }
          break;

        default:
          rowUpdates.status = _App_formatStatus('WARNING', "Unknown Action '" + action + "'");
      }

      return rowUpdates;
    }, {
      onBatchComplete: function (batchResults) {
        var rowNumbers = [];
        var patchData = [];
        batchResults.forEach(function (res) {
          if (res && res._rowNumber !== undefined) {
            rowNumbers.push(res._rowNumber);
            if (res.isError) {
              patchData.push(_App_makeStatusPatch(res._rowNumber, 'ERROR', res.error));
            } else {
              patchData.push(_App_makeRowPatch(res._rowNumber, {
                'Action': res.action,
                'Status': res.status,
                'Task ID': res.taskId,
                'Task List ID': res.listId
              }));
            }
          }
        });
        if (rowNumbers.length > 0) {
          SheetManager.batchPatchRows('TASKS_SYNC', rowNumbers, patchData);
        }
      }
    });

    return _App_ok("Sync Complete. Processed: " + stats.processedCount);
  });
}

// --- INTERNAL HELPERS ---

/** Handles moving a task from one list to another by copying then deleting the old one */
function _TasksSync_processMove(rowUpdates, targetListId, taskData) {
  var dueStr = null;
  if (taskData.dueDate) {
    var d = new Date(taskData.dueDate);
    if (!isNaN(d.getTime())) {
      d.setUTCHours(0, 0, 0, 0);
      dueStr = d.toISOString();
    }
  }

  var taskResource = {
    title: taskData.title,
    notes: taskData.description || "",
    status: taskData.completedVal === 'Completed' ? 'completed' : 'needsAction'
  };
  if (dueStr) taskResource.due = dueStr;

  // Insert into target list
  var newTask = _App_callWithBackoff(function () {
    return Tasks.Tasks.insert(taskResource, targetListId);
  });

  var deleteWarning = "";
  // Delete from original list
  if (rowUpdates.listId && rowUpdates.taskId) {
    try {
      _App_callWithBackoff(function () {
        Tasks.Tasks.remove(rowUpdates.listId, rowUpdates.taskId);
      });
    } catch (e) {
      deleteWarning = " (⚠️ Could not delete old task)";
    }
  }

  rowUpdates.taskId = newTask.id;
  rowUpdates.listId = targetListId;
  rowUpdates.status = _App_formatStatus('SUCCESS', "Moved") + deleteWarning;
  rowUpdates.action = "";

  return rowUpdates;
}

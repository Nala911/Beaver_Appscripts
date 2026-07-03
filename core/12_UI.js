function onOpen() {
  return Logger.run('SYSTEM', 'Initialize UI', function () {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Workspace Sync Tools');

    _App_getMenuTools().forEach(function (cfg) {
      if (cfg.MENU_ENTRYPOINT) {
        menu.addItem(cfg.MENU_LABEL, cfg.MENU_ENTRYPOINT);
      }
    });

    menu.addToUi();
  });
}

function onInstall(e) {
  onOpen(e);
}

// Global settings and Theme Config have been moved to 00_Config_Constants.js and 01_Config_Theme.js
// so that Apps Script evaluates them before the rest of the files.


// ==========================================
// Dialog Server-Side Functions
// ==========================================

function UI_getProgress(toolKey) {
  return Logger.run('SYSTEM', 'Get Progress', function () {
    return _App_ok('Progress', _App_getProgress(toolKey));
  });
}

function UI_getHelpItems(toolKey) {
  return Logger.run('SYSTEM', 'Get Help Items', function () {
    var cfg = SyncEngine.getTool(toolKey);
    return _App_ok('Help Items retrieved', cfg.HELP_ITEMS || null);
  });
}

// ==========================================
// Centralized Validation / Status Checks
// ==========================================

/**
 * Checks for unsaved changes dynamically across any sheet.
 * Called from client sidebars via SyncSidebar wrappers.
 */
function UI_checkForUnsavedChanges(toolKey) {
  return Logger.run(toolKey, 'Check Unsaved', function () {
    var hasChanges = false;
    try {
      hasChanges = SheetManager.hasPendingActions(toolKey);
    } catch (e) {
      // Sheet does not exist or isn't scaffolded yet
    }
    var response = _App_ok('Check complete.', hasChanges);
    response.hasChanges = hasChanges;
    return response;
  });
}

/**
 * Centralized global action dispatcher to route sidebar calls to the SyncEngine.
 * Exposes a single, unified entrypoint for client-side invocations.
 */
function SyncEngine_executeAction(toolKey, actionName, argsJson) {
  var args = argsJson ? JSON.parse(argsJson) : [];
  return SyncEngine.runAction(toolKey, actionName, args);
}


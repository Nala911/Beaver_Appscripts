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

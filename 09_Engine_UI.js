// ==========================================
// _App_openSidebar — Universal Sidebar Opener
// ==========================================
/**
 * Opens a tool's sidebar, ensuring the sheet exists first.
 */
function _App_openSidebar(toolKey, postCreateCallback) {
    var cfg = SyncEngine.getTool(toolKey);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(cfg.SHEET_NAME);

    if (!sheet) {
        sheet = _App_ensureSheetExists(toolKey, postCreateCallback);
    } else {
        sheet.activate();
    }

    var html = HtmlService.createTemplateFromFile(cfg.SIDEBAR_HTML).evaluate()
        .setTitle(cfg.TITLE)
        .setWidth(cfg.SIDEBAR_WIDTH || 300);
    SpreadsheetApp.getUi().showSidebar(html);
}

function _App_launchTool(toolKey, postCreateCallback) {
    return Logger.run(toolKey, 'Launch Tool', function () {
        var cfg = SyncEngine.getTool(toolKey);
        var launchMode = cfg.LAUNCH_MODE || TOOL_LAUNCH_MODES.SIDEBAR;

        if (launchMode === TOOL_LAUNCH_MODES.MODAL) {
            var html = HtmlService.createTemplateFromFile(cfg.MODAL_HTML || cfg.SIDEBAR_HTML).evaluate()
                .setTitle(cfg.TITLE)
                .setWidth(cfg.MODAL_WIDTH || cfg.SIDEBAR_WIDTH || 300)
                .setHeight(cfg.MODAL_HEIGHT || 600);
            SpreadsheetApp.getUi().showModalDialog(html, cfg.TITLE);
            return _App_ok('Modal opened successfully.');
        }

        _App_openSidebar(toolKey, postCreateCallback);
        return _App_ok('Sidebar opened successfully.');
    });
}

function _App_getMenuTools() {
    return Object.keys(SyncEngine.getAllTools())
        .map(function(key) { return SyncEngine.getTool(key); })
        .filter(function(cfg) { return !!cfg.MENU_LABEL; })
        .sort(function(a, b) {
            if (a.MENU_ORDER !== b.MENU_ORDER) return a.MENU_ORDER - b.MENU_ORDER;
            return String(a.MENU_LABEL).localeCompare(String(b.MENU_LABEL));
        });
}
// ==========================================
// _App_ensureSheetExists — Universal Sheet Scaffolding
// ==========================================
/**
 * Creates a tool sheet if it doesn't exist, with headers, column widths,
 * frozen rows/cols, data validations, and buffer body formatting.
 */
function _App_ensureSheetExists(toolKey, postCreateCallback) {
    return Logger.run(toolKey, 'Scaffold Sheet', function () {
        var cfg = SyncEngine.getTool(toolKey);
        if (!_App_canScaffoldSheet(cfg)) {
            throw new Error("Tool '" + toolKey + "' does not define a sheet schema and cannot be scaffolded automatically.");
        }

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(cfg.SHEET_NAME);
        var isNew = !sheet;

        if (isNew) {
            sheet = ss.insertSheet(cfg.SHEET_NAME);
            Logger.info(cfg.TITLE, 'Scaffold', "Created new sheet: " + cfg.SHEET_NAME);
        }

        // Always ensure headers and basic sheet setup are correct (idempotent)
        _App_applyHeaderFormatting(sheet, cfg.HEADERS);

        if (cfg.FROZEN_ROWS > 0) sheet.setFrozenRows(cfg.FROZEN_ROWS);
        if (cfg.FROZEN_COLS > 0) sheet.setFrozenColumns(cfg.FROZEN_COLS);

        if (cfg.COL_WIDTHS) {
            cfg.COL_WIDTHS.forEach(function (w, i) {
                if (w !== null && w !== undefined) sheet.setColumnWidth(i + 1, w);
            });
        }

        if (cfg.FORMAT_CONFIG) {
            _App_applyBodyFormatting(sheet, sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0, cfg.FORMAT_CONFIG);
        }

        if (isNew && typeof postCreateCallback === 'function') {
            try { postCreateCallback(sheet); }
            catch (e) { 
                Logger.warn(cfg.TITLE, 'Post-Scaffold Callback', e.message);
            }
        }

        sheet.activate();
        return sheet;
    });
}

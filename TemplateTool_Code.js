/**
 * Template Backend Tool
 * Version: 1.0 (Plugin Architecture — registers with SyncEngine)
 * 
 * Instructions:
 * 1. Duplicate this file and rename it (e.g. `MyNewTool_Code.js`).
 * 2. Update the Tool Key throughout from `TEMPLATE_TOOL` to your key.
 * 3. Add your `SHEET_NAME` to `00_Config_Constants.js` inside the `SHEET_NAMES` object.
 * 4. Create the corresponding HTML file `MyNewTool_Sidebar.html` from `TemplateTool_Sidebar.html`.
 */

// --- TOOL REGISTRATION ---
SyncEngine.registerTool('TEMPLATE_TOOL', {
    IS_TEMPLATE: true, // Mark as template to skip system audits
    SHEET_NAME: '⚙️ Template Tool', 
    TITLE: '⚙️ Template Tool',
    // MENU_LABEL removed to keep custom menu clean
    MENU_ENTRYPOINT: 'TemplateTool_openSidebar', 
    MENU_ORDER: 90, 
    SIDEBAR_HTML: 'TemplateTool_Sidebar', // Name of the .html file
    SIDEBAR_WIDTH: 300,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [100, 200, 150], // Initial column widths
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 0,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        // Defines the columns and validation rules (headers automatically generated)
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['START', 'STOP', 'PROCESS'] },
            { header: 'Item Name', type: 'TEXT' },
            { header: 'Status', type: 'TEXT' }
        ]
    }
});

// --- PUBLIC ENTRY POINTS ---

/**
 * Triggered from the Custom Menu. Opens the Sidebar and prepares the sheet.
 */
function TemplateTool_openSidebar() {
    return Logger.run('TEMPLATE_TOOL', 'Open Sidebar', function () {
        // Leverages standard engine logic to unhide/create the sheet and inject HTML
        _App_launchTool('TEMPLATE_TOOL');
    });
}

/**
 * Triggered from the Sidebar. Must return { success: boolean, message: string }
 */
function TemplateTool_processAction(payload) {
    return Logger.run('TEMPLATE_TOOL', 'Process Action', function () {
        
        // Example: access user properties
        // var mySavedKey = _App_getProperty(APP_PROPS.SOME_KEY);

        // Try extracting spreadsheet data
        var sheetObj = _App_ensureSheetExists('TEMPLATE_TOOL');
        var dataRange = sheetObj.getDataRange();
        var data = dataRange.getValues();

        if (data.length <= 1) {
            return { success: true, message: "No data to process." };
        }

        // Example: External App call with exponential backoff
        /*
        _App_callWithBackoff(function() {
            DriveApp.getFilesByName('some name');
        });
        */

        return { 
            success: true, 
            message: "Successfully executed the template action!" 
        };
    });
}

// --- INTERNAL HELPERS ---

/**
 * Private helper. Prefix with `_ToolName_`
 */
function _TemplateTool_internalHelper(param) {
    // Perform internal logic
    return param;
}

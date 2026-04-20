/**
 * Template Backend Tool
 * Version: 1.0 (Plugin Architecture — registers with App.Engine)
 * 
 * Instructions:
 * 1. Duplicate this file and rename it (e.g. `MyNewTool_Code.js`).
 * 2. Update the Tool Key throughout from `TEMPLATE_TOOL` to your key.
 * 3. Add your `SHEET_NAME` to `00_Config_Constants.js` inside the `SHEET_NAMES` object.
 * 4. Create the corresponding HTML file `MyNewTool_Sidebar.html` from `Template_Tool_Sidebar.html`.
 */

// --- TOOL REGISTRATION ---
App.Engine.registerTool('TEMPLATE_TOOL', {
    IS_TEMPLATE: true, // Mark as template to skip system audits
    SHEET_NAME: '⚙️ Template Tool', 
    TITLE: '⚙️ Template Tool',
    MENU_ENTRYPOINT: 'TemplateTool_openSidebar', 
    MENU_ORDER: 90, 
    SIDEBAR_HTML: 'Template_Tool_Sidebar', // Name of the .html file
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
    },

    /**
     * TOOL SERVICE ACTIONS
     * These methods are called from the Sidebar via SyncSidebar.call('TEMPLATE_TOOL', 'methodName')
     */
    service: {
        /**
         * Triggered from the Sidebar. Must return { success: boolean, message: string }
         */
        processAction: function(payload) {
            // Example: access tool-specific persistent preferences
            var prefs = App.Engine.getPrefs('TEMPLATE_TOOL');
            var mySavedOption = prefs.myOption || "Default Value";

            // To save:
            // prefs.myOption = "New Value";
            // App.Engine.setPrefs('TEMPLATE_TOOL', prefs);

            // Try extracting spreadsheet data
            // Instead of manually reading ranges, use the ExecutionService for automatic progress tracking, backoff, and logging
            var stats = ExecutionService.processPendingRows('TEMPLATE_TOOL', function(rowObj) {
                
                // Example: External App call with exponential backoff if doing raw API requests
                /*
                _App_callWithBackoff(function() {
                    DriveApp.getFilesByName(rowObj['Item Name']);
                });
                */
                
                // Return an object containing updates for the row
                // The system will automatically mark it as success or failure
                return {
                    'Action': "",
                    'Status': "✅ Processed " + rowObj['Item Name']
                };
            });

            if (stats.total === 0) {
                return _App_ok("No pending actions to process.");
            }

            return _App_ok("Successfully processed " + stats.success + " out of " + stats.total + " actions!");
        }
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
 * @deprecated Use SyncSidebar.call('TEMPLATE_TOOL', 'processAction') from frontend.
 * Keeping for backward compatibility during migration if needed, but should be removed eventually.
 */
function TemplateTool_processAction(payload) {
    return App_exec('TEMPLATE_TOOL', 'processAction', payload);
}

// --- INTERNAL HELPERS ---

/**
 * Private helper. Prefix with `_ToolName_`
 */
function _TemplateTool_internalHelper(param) {
    // Perform internal logic
    return param;
}

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
        
        // 1. Fetch only pending actions (Optimized)
        var pendingItems = SheetManager.readPendingObjects('TEMPLATE_TOOL');

        if (pendingItems.length === 0) {
            return { success: true, message: "No pending actions to process." };
        }

        // 2. Process in batches (handles timeouts and progress tracking)
        var stats = _App_BatchProcessor('TEMPLATE_TOOL', pendingItems, function (item) {
            // Perform logic for each row
            // Access data via headers: item['Item Name']
            
            Logger.info('TEMPLATE_TOOL', 'Row ' + item._rowNumber, 'Processed ' + item['Item Name']);

            // Return updates for this row
            return { 
                action: "", 
                status: "✅ Done", 
                _rowNumber: item._rowNumber 
            };
        }, {
            onBatchComplete: function (batchResults) {
                // Write back updates to the sheet efficiently
                var rowNumbers = batchResults.map(r => r._rowNumber);
                var patchData = batchResults.map(r => ({ 'Action': r.action, 'Status': r.status }));
                SheetManager.batchPatchRows('TEMPLATE_TOOL', rowNumbers, patchData);
            }
        });

        return { 
            success: true, 
            message: "Successfully processed " + stats.processedCount + " items!" 
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

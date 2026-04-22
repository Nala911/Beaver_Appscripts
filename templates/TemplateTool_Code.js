/**
 * Template Backend Tool
 * Version: 2.0 (Plugin Architecture — registers with SyncEngine)
 * 
 * Instructions:
 * 1. Copy the files from the `templates/` directory to the root.
 * 2. Rename them (e.g., `MyNewTool_Code.js` and `MyNewTool_Sidebar.html`).
 * 3. Update the Tool Key throughout from `TEMPLATE_TOOL` to your key.
 * 4. Add your `SHEET_NAME` to `00_Config_Constants.js` inside the `SHEET_NAMES` object.
 */

// --- TOOL REGISTRATION ---
SyncEngine.registerTool('TEMPLATE_TOOL', {
    // 1. Required Services (Checks for Advanced APIs)
    REQUIRED_SERVICES: [
        /* { name: 'Drive API', test: function() { return typeof Drive !== 'undefined'; } } */
    ],
    IS_TEMPLATE: true, // Mark as template to skip system audits
    SHEET_NAME: '⚙️ Template Tool', 
    TITLE: '⚙️ Template Tool',
    MENU_ENTRYPOINT: 'TemplateTool_openSidebar', 
    MENU_ORDER: 90, 
    SIDEBAR_HTML: 'TemplateTool_Sidebar', 
    SIDEBAR_WIDTH: 300,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [100, 200, 150], 
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 0,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['START', 'STOP', 'PROCESS'] },
            { header: 'Item Name', type: 'TEXT' },
            { header: 'Status', type: 'TEXT' }
        ]
    }
});

/* ==========================================================================
   CONFIGURATION / ALIASES
   ========================================================================== */

/**
 * Column-index aliases for easier reference. 
 * Update these if your COL_SCHEMA changes.
 */
var TEMPLATE_TOOL_COL = {
    ACTION: 0,
    NAME: 1,
    STATUS: 2
};

// --- SIDEBAR & SHEET SETUP ---

/**
 * Triggered from the Custom Menu. Opens the Sidebar and prepares the sheet.
 */
function TemplateTool_openSidebar() {
    return Logger.run('TEMPLATE_TOOL', 'Open Sidebar', function () {
        // Leverages standard engine logic to unhide/create the sheet and inject HTML
        _App_launchTool('TEMPLATE_TOOL');
    });
}

/* ==========================================================================
   CORE LOGIC
   ========================================================================== */

/**
 * Triggered from the Sidebar. Must return { success: boolean, message: string }
 */
function TemplateTool_processAction(payload) {
    return Logger.run('TEMPLATE_TOOL', 'Process Action', function () {
        
        // 1. Fetch only pending actions (Optimized)
        var pendingItems = SheetManager.readPendingObjects('TEMPLATE_TOOL');

        if (pendingItems.length === 0) {
            return _App_ok("No pending actions to process.");
        }

        // 2. Process in batches (handles timeouts and progress tracking)
        var stats = _App_BatchProcessor('TEMPLATE_TOOL', pendingItems, function (item) {
            // Perform logic for each row
            // Access data via headers: item['Item Name']
            
            Logger.info('TEMPLATE_TOOL', 'Row ' + item._rowNumber, 'Processed ' + item['Item Name']);

            // Return updates for this row
            var rowUpdates = {};
            rowUpdates['Action'] = "";
            rowUpdates['Status'] = "✅ Done";
            
            return { 
                _rowNumber: item._rowNumber,
                updates: rowUpdates
            };
        }, {
            onBatchComplete: function (batchResults) {
                // Write back updates to the sheet efficiently
                var rowNumbers = batchResults.map(function(r) { return r._rowNumber; });
                var patchData = batchResults.map(function(r) { return r.updates; });
                SheetManager.batchPatchRows('TEMPLATE_TOOL', rowNumbers, patchData);
            }
        });

        var msg = "Successfully processed " + stats.processedCount + " items!";
        if (stats.timeLimitReached) {
            msg += " (Execution Time Limit Reached. Run again to continue)";
        }

        return _App_ok(msg, { stats: stats });
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

/**
 * Gmail Filters Tool
 * Version: 1.0 (Plugin Architecture)
 * 
 * Allows users to manage Gmail filters directly from the spreadsheet.
 */

// --- TOOL REGISTRATION ---
SyncEngine.registerTool('GMAIL_FILTERS', {
    REQUIRED_SERVICES: [ { name: 'Gmail API', test: function() { return typeof Gmail !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.GMAIL_FILTERS, 
    TITLE: SHEET_NAMES.GMAIL_FILTERS,
    MENU_LABEL: SHEET_NAMES.GMAIL_FILTERS,
    MENU_ENTRYPOINT: 'GmailFilters_openSidebar', 
    MENU_ORDER: 30, 
    SIDEBAR_HTML: 'GmailFilters_Sidebar',
    SIDEBAR_WIDTH: 320,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 2,
    COL_WIDTHS: [100, 150, 150, 150, 150, 150, 200, 200, 100, 200, 200, 200, 100, 100, 100],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 0,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'TEXT' },
            { header: 'Filter ID', type: 'TEXT' },
            { header: 'Criteria: From', type: 'TEXT' },
            { header: 'Criteria: To', type: 'TEXT' },
            { header: 'Criteria: Subject', type: 'TEXT' },
            { header: 'Criteria: Includes Words', type: 'TEXT' },
            { header: 'Criteria: Excludes Words', type: 'TEXT' },
            { header: 'Criteria: Has Attachment', type: 'BOOLEAN' },
            { header: 'Action: Add Labels', type: 'TEXT' },
            { header: 'Action: Remove Labels', type: 'TEXT' },
            { header: 'Action: Forward To', type: 'TEXT' },
            { header: 'Action: Mark Read', type: 'BOOLEAN' },
            { header: 'Action: Mark Important', type: 'BOOLEAN' },
            { header: 'Action: Delete', type: 'BOOLEAN' }
        ]
    }
});

// --- PUBLIC ENTRY POINTS ---

/**
 * Opens the Sidebar and prepares the sheet.
 */
function GmailFilters_openSidebar() {
    return Logger.run('GMAIL_FILTERS', 'Open Sidebar', function () {
        _App_launchTool('GMAIL_FILTERS');
    });
}

/**
 * Pulls all Gmail filters into the spreadsheet.
 */
function GmailFilters_pullFilters() {
    return Logger.run('GMAIL_FILTERS', 'Pull Filters', function () {
        var labels = _GmailFilters_getLabelMap();
        var filtersResponse = _App_callWithBackoff(function() {
            return Gmail.Users.Settings.Filters.list('me');
        });

        var filters = filtersResponse.filter || [];
        var rows = filters.map(function(f) {
            var criteria = f.criteria || {};
            var action = f.action || {};

            return {
                'Action': '',
                'Status': 'Synced',
                'Filter ID': f.id,
                'Criteria: From': criteria.from || '',
                'Criteria: To': criteria.to || '',
                'Criteria: Subject': criteria.subject || '',
                'Criteria: Includes Words': criteria.query || '',
                'Criteria: Excludes Words': criteria.negatedQuery || '',
                'Criteria: Has Attachment': !!criteria.hasAttachment,
                'Action: Add Labels': _GmailFilters_resolveLabelIds(action.addLabelIds, labels.idToName),
                'Action: Remove Labels': _GmailFilters_resolveLabelIds(action.removeLabelIds, labels.idToName),
                'Action: Forward To': action.forward || '',
                'Action: Mark Read': action.removeLabelIds && action.removeLabelIds.indexOf('UNREAD') !== -1,
                'Action: Mark Important': action.addLabelIds && action.addLabelIds.indexOf('IMPORTANT') !== -1,
                'Action: Delete': action.addLabelIds && action.addLabelIds.indexOf('TRASH') !== -1
            };
        });

        // Clear existing and write new
        SheetManager.overwriteObjects('GMAIL_FILTERS', rows);

        return _App_ok("Successfully pulled " + rows.length + " filters.");
    });
}

/**
 * Processes pending actions (CREATE, UPDATE, DELETE).
 */
function GmailFilters_processAction() {
    return Logger.run('GMAIL_FILTERS', 'Process Action', function () {
        var pendingItems = SheetManager.readPendingObjects('GMAIL_FILTERS');
        if (pendingItems.length === 0) {
            return { success: true, message: "No pending actions to process." };
        }

        var labelMap = _GmailFilters_getLabelMap();

        var stats = _App_BatchProcessor('GMAIL_FILTERS', pendingItems, function (item) {
            var actionType = item['Action'];
            var filterId = item['Filter ID'];
            var resultStatus = "✅ Success";
            var resultAction = "";
            var newFilterId = filterId;

            try {
                if (actionType === 'DELETE' || actionType === 'UPDATE') {
                    if (!filterId) throw new Error("Missing Filter ID for " + actionType);
                    _App_callWithBackoff(function() {
                        Gmail.Users.Settings.Filters.remove('me', filterId);
                    });
                    if (actionType === 'DELETE') {
                        return { action: "", status: "✅ Deleted", _rowNumber: item._rowNumber };
                    }
                }

                if (actionType === 'CREATE' || actionType === 'UPDATE') {
                    var filterResource = _GmailFilters_constructFilterResource(item, labelMap.nameToId);
                    var createdFilter = _App_callWithBackoff(function() {
                        return Gmail.Users.Settings.Filters.create(filterResource, 'me');
                    });
                    newFilterId = createdFilter.id;
                    resultStatus = (actionType === 'UPDATE') ? "✅ Updated" : "✅ Created";
                }

            } catch (e) {
                Logger.error('GMAIL_FILTERS', 'Row ' + item._rowNumber, e.message);
                return { action: actionType, status: "❌ " + e.message, _rowNumber: item._rowNumber };
            }

            return { 
                action: resultAction, 
                status: resultStatus, 
                'Filter ID': newFilterId,
                _rowNumber: item._rowNumber 
            };
        }, {
            onBatchComplete: function (batchResults) {
                var rowNumbers = batchResults.map(r => r._rowNumber);
                var patchData = batchResults.map(r => {
                    var patch = { 'Action': r.action, 'Status': r.status };
                    if (r['Filter ID']) patch['Filter ID'] = r['Filter ID'];
                    return patch;
                });
                SheetManager.batchPatchRows('GMAIL_FILTERS', rowNumbers, patchData);
            }
        });

        return _App_ok("Processed " + stats.processedCount + " filters.");
    });
}

// --- INTERNAL HELPERS ---

/**
 * Fetches Gmail labels and creates mapping objects.
 */
function _GmailFilters_getLabelMap() {
    var response = _App_callWithBackoff(function() {
        return Gmail.Users.Labels.list('me');
    });
    
    var nameToId = {};
    var idToName = {};

    (response.labels || []).forEach(function(l) {
        nameToId[l.name] = l.id;
        idToName[l.id] = l.name;
    });

    return { nameToId: nameToId, idToName: idToName };
}

/**
 * Resolves Label IDs to comma-separated Names.
 */
function _GmailFilters_resolveLabelIds(ids, idToName) {
    if (!ids || !Array.isArray(ids)) return '';
    return ids.map(function(id) {
        return idToName[id] || id;
    }).join(', ');
}

/**
 * Constructs a Gmail Filter resource from sheet data.
 */
function _GmailFilters_constructFilterResource(item, nameToId) {
    var criteria = {};
    if (item['Criteria: From']) criteria.from = item['Criteria: From'];
    if (item['Criteria: To']) criteria.to = item['Criteria: To'];
    if (item['Criteria: Subject']) criteria.subject = item['Criteria: Subject'];
    if (item['Criteria: Includes Words']) criteria.query = item['Criteria: Includes Words'];
    if (item['Criteria: Excludes Words']) criteria.negatedQuery = item['Criteria: Excludes Words'];
    if (item['Criteria: Has Attachment']) criteria.hasAttachment = true;

    var action = {};
    var addLabelIds = [];
    var removeLabelIds = [];

    // Process Labels
    if (item['Action: Add Labels']) {
        item['Action: Add Labels'].split(',').forEach(function(n) {
            var id = nameToId[n.trim()];
            if (id) addLabelIds.push(id);
        });
    }
    if (item['Action: Remove Labels']) {
        item['Action: Remove Labels'].split(',').forEach(function(n) {
            var id = nameToId[n.trim()];
            if (id) removeLabelIds.push(id);
        });
    }

    // Process Boolean Flags
    if (item['Action: Mark Read']) {
        if (removeLabelIds.indexOf('UNREAD') === -1) removeLabelIds.push('UNREAD');
    }
    if (item['Action: Mark Important']) {
        if (addLabelIds.indexOf('IMPORTANT') === -1) addLabelIds.push('IMPORTANT');
    }
    if (item['Action: Delete']) {
        if (addLabelIds.indexOf('TRASH') === -1) addLabelIds.push('TRASH');
    }

    if (item['Action: Forward To']) action.forward = item['Action: Forward To'];
    if (addLabelIds.length > 0) action.addLabelIds = addLabelIds;
    if (removeLabelIds.length > 0) action.removeLabelIds = removeLabelIds;

    return { criteria: criteria, action: action };
}

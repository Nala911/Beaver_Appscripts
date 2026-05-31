/**
 * Gmail Filters Tool
 * Version: 1.0 (Plugin Architecture)
 * 
 * Allows users to manage Gmail filters directly from the spreadsheet.
 */

// --- TOOL REGISTRATION ---
SyncEngine.registerTool('GMAIL_FILTERS', {
    REQUIRED_SERVICES: [{ name: 'Gmail API', test: function () { return typeof Gmail !== 'undefined'; } }],
    SHEET_NAME: SHEET_NAMES.GMAIL_FILTERS,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.GMAIL_FILTERS,
    MENU_LABEL: SHEET_NAMES.GMAIL_FILTERS,
    MENU_ENTRYPOINT: 'GmailFilters_openSidebar',
    MENU_ORDER: 30,
    SIDEBAR_HTML: 'GmailFilters_Sidebar',
    SIDEBAR_WIDTH: 320,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Criteria: From', type: 'TEXT' },
            { header: 'Criteria: To', type: 'TEXT' },
            { header: 'Criteria: Subject', type: 'TEXT' },
            { header: 'Criteria: Includes Words', type: 'TEXT' },
            { header: 'Criteria: Excludes Words', type: 'TEXT' },
            { header: 'Criteria: Has Attachment', type: 'CHECKBOX' },
            { header: 'Action: Skip the Inbox (Archive it)', type: 'CHECKBOX' },
            { header: 'Action: Mark as read', type: 'CHECKBOX' },
            { header: 'Action: Star it', type: 'CHECKBOX' },
            {
                header: 'Action: Labels', type: 'DROPDOWN', allowInvalid: true, options: function () {
                    var labels = [];
                    try {
                        var response = _App_callWithBackoff(function () { return Gmail.Users.Labels.list('me'); });
                        var systemLabelIds = ['INBOX', 'UNREAD', 'STARRED', 'TRASH', 'SPAM', 'IMPORTANT', 'CHAT', 'DRAFT', 'GREEN_CIRCLE', 'SENT', 'YELLOW_STAR'];
                        (response.labels || []).forEach(function (l) {
                            if (systemLabelIds.indexOf(l.id) === -1 && !l.id.startsWith('CATEGORY_')) {
                                labels.push(l.name);
                            }
                        });
                        labels.sort();
                    } catch (e) { }
                    return labels.length ? labels.slice(0, 499) : ['None'];
                }
            },
            { header: 'Action: Forward to', type: 'TEXT' },
            { header: 'Action: Delete it', type: 'CHECKBOX' },
            { header: 'Action: Never send it to Spam', type: 'CHECKBOX' },
            { header: 'Action: Always mark it as important', type: 'CHECKBOX' },
            { header: 'Action: Never mark it as important', type: 'CHECKBOX' },
            { header: 'Action: Also apply filter to previous mails', type: 'CHECKBOX' },
            { header: 'Filter ID', type: 'ID' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to manage Gmail filters:</p><ol><li><strong>Define Criteria:</strong> Set filter matching rules under <code>Criteria:</code> columns (From, Subject, etc.).</li><li><strong>Set Action:</strong> Set output settings (Skip Inbox, Labels, Delete, etc.).</li><li><strong>Push:</strong> Click <strong>Push Changes</strong> in the sidebar.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Understand filter criteria and action flags.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Criteria:</strong> The filter rules that match incoming messages (e.g. <code>Criteria: From</code>).</li><li><strong>Action: Labels:</strong> Custom label to apply to matching messages.</li><li><strong>Retroactive:</strong> Tick <code>Also apply filter to previous mails</code> to run the filter matching on past emails.</li></ul>"
            }
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
        return _App_withDocumentLock('GMAIL_FILTERS_PULL', function () {
            var labels = _GmailFilters_getLabelMap();
        var filtersResponse = _App_callWithBackoff(function () {
            return Gmail.Users.Settings.Filters.list('me');
        });

        var filters = (filtersResponse && filtersResponse.filter) || [];
        var rows = filters.map(function (f) {
            var criteria = f.criteria || {};
            var action = f.action || {};
            var addLabelIds = action.addLabelIds || [];
            var removeLabelIds = action.removeLabelIds || [];

            // Filter out system labels for the "Labels" column (Take only the first custom label)
            var systemLabelIds = ['INBOX', 'UNREAD', 'STARRED', 'TRASH', 'SPAM', 'IMPORTANT'];
            var userLabelIds = addLabelIds.filter(function (id) {
                return systemLabelIds.indexOf(id) === -1 && !id.startsWith('CATEGORY_');
            });
            var singleLabelId = userLabelIds.length > 0 ? userLabelIds[0] : null;

            return {
                'Action': '',
                'Status': 'Synced',
                'Criteria: From': criteria.from || '',
                'Criteria: To': criteria.to || '',
                'Criteria: Subject': criteria.subject || '',
                'Criteria: Includes Words': criteria.query || '',
                'Criteria: Excludes Words': criteria.negatedQuery || '',
                'Criteria: Has Attachment': !!criteria.hasAttachment,
                'Action: Skip the Inbox (Archive it)': removeLabelIds.indexOf('INBOX') !== -1,
                'Action: Mark as read': removeLabelIds.indexOf('UNREAD') !== -1,
                'Action: Star it': addLabelIds.indexOf('STARRED') !== -1,
                'Action: Labels': singleLabelId ? labels.idToName[singleLabelId] || singleLabelId : '',
                'Action: Forward to': action.forward || '',
                'Action: Delete it': addLabelIds.indexOf('TRASH') !== -1,
                'Action: Never send it to Spam': removeLabelIds.indexOf('SPAM') !== -1,
                'Action: Always mark it as important': addLabelIds.indexOf('IMPORTANT') !== -1,
                'Action: Never mark it as important': removeLabelIds.indexOf('IMPORTANT') !== -1,
                'Action: Also apply filter to previous mails': false,
                'Filter ID': f.id
            };
        });

            // Clear existing and write new
            SheetManager.overwriteObjects('GMAIL_FILTERS', rows);

            return _App_ok("Successfully pulled " + rows.length + " filters.");
        });
    });
}

/**
 * Checks if the Gmail Filters sheet has pending actions.
 */


/**
 * Processes pending actions (CREATE, UPDATE, DELETE).
 */
function GmailFilters_processAction() {
    return Logger.run('GMAIL_FILTERS', 'Process Action', function () {
        return _App_withDocumentLock('GMAIL_FILTERS_PUSH', function () {
            var pendingItems = SheetManager.readPendingObjects('GMAIL_FILTERS');
            if (pendingItems.length === 0) {
                return _App_ok("No pending actions to process.");
            }

        var labelMap = _GmailFilters_getLabelMap();

        var stats = _App_BatchProcessor('GMAIL_FILTERS', pendingItems, function (item) {
            var actionType = item['Action'];
            var filterId = item['Filter ID'];
            var resultStatus = _App_formatStatus('SUCCESS', "Success");
            var resultAction = "";
            var newFilterId = filterId;

            if (actionType === 'DELETE' || actionType === 'UPDATE') {
                    if (!filterId) throw new Error("Missing Filter ID for " + actionType);
                    try {
                        _App_callWithBackoff(function () {
                            Gmail.Users.Settings.Filters.remove('me', filterId);
                        });
                    } catch (e) {
                        // Special handling for Apps Script quirk: DELETE/REMOVE often returns 204 (No Content)
                        // which Apps Script occasionally misinterprets as an "Empty response" error.
                        if (e.message.indexOf("Empty response") !== -1) {
                            // Ignored "Empty response" for filter delete/update (likely success)
                        } else {
                            throw e;
                        }
                    }
                    if (actionType === 'DELETE') {
                        return { action: "", status: _App_formatStatus('SUCCESS', "Deleted"), _rowNumber: item._rowNumber };
                    }
                }

                if (actionType === 'CREATE' || actionType === 'UPDATE') {
                    var filterResource = _GmailFilters_constructFilterResource(item, labelMap.nameToId);
                    var createdFilter = _App_callWithBackoff(function () {
                        return Gmail.Users.Settings.Filters.create(filterResource, 'me');
                    });
                    newFilterId = createdFilter.id;
                    resultStatus = (actionType === 'UPDATE') ? _App_formatStatus('SUCCESS', "Updated") : _App_formatStatus('SUCCESS', "Created");

                // Handle retroactive application
                if (item['Action: Also apply filter to previous mails']) {
                    var searchQuery = _GmailFilters_buildSearchQuery(filterResource.criteria);
                    _GmailFilters_applyToExistingMessages(searchQuery, filterResource.action.addLabelIds || [], filterResource.action.removeLabelIds || []);
                    resultStatus += " (+ Applied to existing)";
                }
            }

            return {
                action: resultAction,
                status: resultStatus,
                'Filter ID': newFilterId,
                _rowNumber: item._rowNumber
            };
        }, {
            onBatchComplete: function (batchResults) {
                _App_batchPatchResults('GMAIL_FILTERS', batchResults, function (res) {
                    var fields = {};
                    if (res['Filter ID'] !== undefined) {
                        fields['Filter ID'] = res['Filter ID'];
                    }
                    return fields;
                });
            }
        });

            return _App_ok("Processed " + stats.processedCount + " filters.");
        });
    });
}

/**
 * Checks for labels in the pending rows that don't exist in Gmail.
 * Used for pre-push confirmation in the sidebar.
 */
function GmailFilters_getMissingLabels() {
    return Logger.run('GMAIL_FILTERS', 'Check Missing Labels', function () {
        var pendingItems = SheetManager.readPendingObjects('GMAIL_FILTERS');
        if (pendingItems.length === 0) return _App_ok('No pending actions.', []);

        var labelsInSheet = [];
        pendingItems.forEach(function (item) {
            var action = (item['Action'] || '').toString().toUpperCase();
            if (action === 'CREATE' || action === 'UPDATE') {
                var label = item['Action: Labels'] ? item['Action: Labels'].trim() : '';
                if (label && labelsInSheet.indexOf(label) === -1) {
                    labelsInSheet.push(label);
                }
            }
        });

        if (labelsInSheet.length === 0) return _App_ok('No labels to check.', []);

        var labelMap = _GmailFilters_getLabelMap();
        var missing = labelsInSheet.filter(function (name) {
            return !labelMap.nameToId[name];
        });

        return _App_ok('Missing labels identified.', missing);
    });
}

// --- INTERNAL HELPERS ---

/**
 * Fetches Gmail labels and creates mapping objects.
 */
function _GmailFilters_getLabelMap() {
    var response = _App_callWithBackoff(function () {
        return Gmail.Users.Labels.list('me');
    });

    var nameToId = {};
    var idToName = {};

    (response.labels || []).forEach(function (l) {
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
    return ids.map(function (id) {
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

    // Process Boolean Flags -> addLabelIds
    if (item['Action: Star it']) addLabelIds.push('STARRED');
    if (item['Action: Delete it']) addLabelIds.push('TRASH');
    if (item['Action: Always mark it as important']) addLabelIds.push('IMPORTANT');

    // Process Boolean Flags -> removeLabelIds
    if (item['Action: Skip the Inbox (Archive it)']) removeLabelIds.push('INBOX');
    if (item['Action: Mark as read']) removeLabelIds.push('UNREAD');
    if (item['Action: Never send it to Spam']) removeLabelIds.push('SPAM');
    if (item['Action: Never mark it as important']) removeLabelIds.push('IMPORTANT');

    // Process Label with Auto-Creation (Treat as single label)
    if (item['Action: Labels']) {
        var labelName = item['Action: Labels'].trim();
        if (labelName) {
            var id = nameToId[labelName];
            if (!id) {
                // Auto-create missing label
                var newLabel = _App_callWithBackoff(function () {
                    return Gmail.Users.Labels.create({ name: labelName }, 'me');
                });
                id = newLabel.id;
                nameToId[labelName] = id; // Update map for subsequent rows in same batch
            }
            if (addLabelIds.indexOf(id) === -1) addLabelIds.push(id);
        }
    }

    if (item['Action: Forward to']) action.forward = item['Action: Forward to'];
    if (addLabelIds.length > 0) action.addLabelIds = addLabelIds;
    if (removeLabelIds.length > 0) action.removeLabelIds = removeLabelIds;

    return { criteria: criteria, action: action };
}

/**
 * Builds a Gmail search query string from filter criteria.
 */
function _GmailFilters_buildSearchQuery(criteria) {
    var queryParts = [];
    if (criteria.from) queryParts.push('from:(' + criteria.from + ')');
    if (criteria.to) queryParts.push('to:(' + criteria.to + ')');
    if (criteria.subject) queryParts.push('subject:(' + criteria.subject + ')');
    if (criteria.query) queryParts.push(criteria.query);
    if (criteria.negatedQuery) queryParts.push('-(' + criteria.negatedQuery + ')');
    if (criteria.hasAttachment) queryParts.push('has:attachment');
    return queryParts.join(' ').trim();
}

/**
 * Applies labels to up to 1000 existing messages matching the query.
 */
function _GmailFilters_applyToExistingMessages(query, addLabelIds, removeLabelIds) {
    if (!query) return;

    var response = _App_callWithBackoff(function () {
        return Gmail.Users.Messages.list('me', { q: query, maxResults: 1000 });
    });

    if (response.messages && response.messages.length > 0) {
        var messageIds = response.messages.map(function (m) { return m.id; });
        _App_callWithBackoff(function () {
            Gmail.Users.Messages.batchModify({
                ids: messageIds,
                addLabelIds: addLabelIds.length > 0 ? addLabelIds : undefined,
                removeLabelIds: removeLabelIds.length > 0 ? removeLabelIds : undefined
            }, 'me');
        });
    }
}


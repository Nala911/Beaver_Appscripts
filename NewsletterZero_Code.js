/**
 * Newsletter Zero (The Inbox Purge)
 * Version: 1.0 (Plugin Architecture)
 * 
 * Scans the inbox for bulk senders and newsletters, providing a bulk unsubscribe/archive interface.
 */

// --- TOOL REGISTRATION ---
SyncEngine.registerTool('NEWSLETTER_ZERO', {
    REQUIRED_SERVICES: [{ name: 'Gmail API', test: function () { return typeof Gmail !== 'undefined'; } }],
    SHEET_NAME: SHEET_NAMES.NEWSLETTER_ZERO,
    TITLE: SHEET_NAMES.NEWSLETTER_ZERO,
    MENU_LABEL: SHEET_NAMES.NEWSLETTER_ZERO,
    MENU_ENTRYPOINT: 'NewsletterZero_openSidebar',
    MENU_ORDER: 35,
    SIDEBAR_HTML: 'NewsletterZero_Sidebar',
    SIDEBAR_WIDTH: 350,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 1,
    COL_WIDTHS: [120, 200, 250, 150, 100, 400],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 0,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['UNSUBSCRIBE', 'ARCHIVE ALL', 'DELETE ALL', 'IGNORE'] },
            { header: 'Sender Name', type: 'TEXT' },
            { header: 'Sender Email', type: 'TEXT' },
            { header: 'Category', type: 'TEXT' },
            { header: 'Email Count', type: 'TEXT' },
            { header: 'Unsubscribe Details', type: 'TEXT', italic: true }
        ]
    }
});

// --- PUBLIC ENTRY POINTS ---

/**
 * Opens the Sidebar and prepares the sheet.
 */
function NewsletterZero_openSidebar() {
    return Logger.run('NEWSLETTER_ZERO', 'Open Sidebar', function () {
        _App_launchTool('NEWSLETTER_ZERO');
    });
}

/**
 * Scans the inbox for newsletters and bulk senders.
 */
function NewsletterZero_pullSenders(daysToScan) {
    return Logger.run('NEWSLETTER_ZERO', 'Scan Inbox', function () {
        daysToScan = daysToScan || 30;
        _App_setProperty(APP_PROPS.NEWSLETTER_ZERO_SCAN_DAYS, daysToScan);
        
        var query = 'newer_than:' + daysToScan + 'd';
        var threads = GmailApp.search(query, 0, 500); // Sample up to 500 threads
        var senderMap = {};

        _App_setProgress('NEWSLETTER_ZERO', 'Scanning ' + threads.length + ' threads...', 10);

        threads.forEach(function (thread, index) {
            if (index % 50 === 0) {
                _App_setProgress('NEWSLETTER_ZERO', 'Analyzing thread ' + (index + 1) + '/' + threads.length, 10 + (index / threads.length * 80));
            }

            var messages = thread.getMessages();
            var firstMsg = messages[0];
            var from = firstMsg.getFrom();
            var rawContent = firstMsg.getRawContent();
            
            // Extract Email and Name
            var emailMatch = from.match(/<([^>]+)>/);
            var email = emailMatch ? emailMatch[1] : from;
            var name = from.split('<')[0].replace(/"/g, '').trim() || email;

            // Detection Logic
            var unsubscribeHeader = _NewsletterZero_getUnsubscribeHeader(rawContent);
            var isBulk = !!unsubscribeHeader;
            
            // Heuristic: If we've replied, it's probably personal
            if (isBulk && !_NewsletterZero_hasUserReplied(thread)) {
                if (!senderMap[email]) {
                    senderMap[email] = {
                        name: name,
                        email: email,
                        count: 0,
                        category: _NewsletterZero_getCategory(firstMsg),
                        unsubscribe: unsubscribeHeader
                    };
                }
                senderMap[email].count += messages.length;
            }
        });

        var rows = Object.keys(senderMap).map(function (email) {
            var data = senderMap[email];
            return {
                'Action': '',
                'Sender Name': data.name,
                'Sender Email': data.email,
                'Category': data.category,
                'Email Count': data.count,
                'Unsubscribe Details': data.unsubscribe || 'None found'
            };
        });

        // Sort by count descending
        rows.sort(function (a, b) { return b['Email Count'] - a['Email Count']; });

        SheetManager.overwriteObjects('NEWSLETTER_ZERO', rows);
        return _App_ok("Found " + rows.length + " bulk senders.");
    });
}

/**
 * Processes the selected actions (Unsubscribe, Archive, Delete).
 */
function NewsletterZero_processActions() {
    return Logger.run('NEWSLETTER_ZERO', 'Process Actions', function () {
        var pendingItems = SheetManager.readPendingObjects('NEWSLETTER_ZERO');
        if (pendingItems.length === 0) {
            return _App_ok("No pending actions to process.");
        }

        var stats = _App_BatchProcessor('NEWSLETTER_ZERO', pendingItems, function (item) {
            var actionType = item['Action'];
            var email = item['Sender Email'];
            var unsubscribeDetails = item['Unsubscribe Details'];
            var resultStatus = "✅ Success";

            try {
                if (actionType === 'UNSUBSCRIBE') {
                    _NewsletterZero_doUnsubscribe(unsubscribeDetails);
                }

                if (actionType === 'ARCHIVE ALL' || actionType === 'DELETE ALL') {
                    _NewsletterZero_bulkCleanup(email, actionType === 'DELETE ALL');
                }

                return { action: "", status: resultStatus, _rowNumber: item._rowNumber };
            } catch (e) {
                return { action: actionType, status: "❌ " + e.message, _rowNumber: item._rowNumber };
            }
        }, {
            onBatchComplete: function (batchResults) {
                var rowNumbers = batchResults.map(r => r._rowNumber);
                var patchData = batchResults.map(r => ({ 'Action': r.action }));
                SheetManager.batchPatchRows('NEWSLETTER_ZERO', rowNumbers, patchData);
            }
        });

        return _App_ok("Processed " + stats.processedCount + " actions.");
    });
}

// --- INTERNAL HELPERS ---

/**
 * Extracts List-Unsubscribe header from raw message content.
 */
function _NewsletterZero_getUnsubscribeHeader(raw) {
    var match = raw.match(/^List-Unsubscribe:\s*(.+)$/im);
    return match ? match[1].trim() : null;
}

/**
 * Detects Gmail Category (Promotions, Social, etc.)
 */
function _NewsletterZero_getCategory(message) {
    try {
        var threadId = message.getThread().getId();
        var thread = Gmail.Users.Threads.get('me', threadId);
        var labelIds = thread.messages[0].labelIds || [];
        
        if (labelIds.indexOf('CATEGORY_PROMOTIONS') !== -1) return 'Promotions';
        if (labelIds.indexOf('CATEGORY_SOCIAL') !== -1) return 'Social';
        if (labelIds.indexOf('CATEGORY_UPDATES') !== -1) return 'Updates';
        if (labelIds.indexOf('CATEGORY_FORUMS') !== -1) return 'Forums';
        return 'General';
    } catch (e) {
        return 'General';
    }
}

/**
 * Checks if the user has ever replied in this thread.
 */
function _NewsletterZero_hasUserReplied(thread) {
    var messages = thread.getMessages();
    for (var i = 0; i < messages.length; i++) {
        if (messages[i].getFrom().indexOf(Session.getActiveUser().getEmail()) !== -1) {
            return true;
        }
    }
    return false;
}

/**
 * Executes the actual unsubscribe action.
 */
function _NewsletterZero_doUnsubscribe(details) {
    if (!details || details === 'None found') return;

    // List-Unsubscribe can contain multiple options like <mailto:..>, <http:..>
    var links = details.match(/<([^>]+)>/g);
    if (!links) return;

    var httpLink = null;
    var mailtoLink = null;

    links.forEach(function (l) {
        var url = l.replace(/[<>]/g, '');
        if (url.startsWith('http')) httpLink = url;
        if (url.startsWith('mailto')) mailtoLink = url;
    });

    // Prefer HTTP for automation
    if (httpLink) {
        _App_callWithBackoff(function () {
            UrlFetchApp.fetch(httpLink, { muteHttpExceptions: true });
        });
    } else if (mailtoLink) {
        var parts = mailtoLink.replace('mailto:', '').split('?');
        var to = parts[0];
        var subject = "Unsubscribe";
        if (parts[1] && parts[1].indexOf('subject=') !== -1) {
            subject = decodeURIComponent(parts[1].split('subject=')[1].split('&')[0]);
        }
        _App_callWithBackoff(function () {
            GmailApp.sendEmail(to, subject, "Please unsubscribe me from this list.");
        });
    }
}

/**
 * Archives or Deletes all threads from a specific sender.
 */
function _NewsletterZero_bulkCleanup(email, shouldDelete) {
    var threads = GmailApp.search('from:' + email);
    var batchSize = 100;
    
    for (var i = 0; i < threads.length; i += batchSize) {
        var batch = threads.slice(i, i + batchSize);
        if (shouldDelete) {
            _App_callWithBackoff(function () { GmailApp.moveThreadsToTrash(batch); });
        } else {
            _App_callWithBackoff(function () { GmailApp.moveThreadsToArchive(batch); });
        }
    }
}

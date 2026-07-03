/**
 * Mail Merge
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('MAIL_MERGE', {
    SHEET_NAME: SHEET_NAMES.MAIL_MERGE,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.MAIL_MERGE,
    MENU_LABEL: SHEET_NAMES.MAIL_MERGE,
    MENU_ENTRYPOINT: 'MailMerge_openSidebar',
    MENU_ORDER: 30,
    SIDEBAR_HTML: 'tools/MailMerge/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [
            { type: 'pending', actionCol: 'A', scope: 'actionOnly' },
            { type: 'success', statusCol: 'B', scope: 'fullRow' },
            { type: 'error', statusCol: 'B', scope: 'fullRow' }
        ],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['SEND', 'DRAFT'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'To', type: 'EMAIL_LIST' },
            { header: 'CC', type: 'EMAIL_LIST' },
            { header: 'BCC', type: 'EMAIL_LIST' },
            { header: 'Thread ID or Subject', type: 'TEXT' },
            { header: 'Attachments', type: 'TEXT' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these 3 steps to run your first merge:</p><ol><li><strong>Gmail Draft:</strong> Create a draft with tags like <code>{{First Name}}</code>.</li><li><strong>Select & Sync:</strong> Choose the draft above and click <strong>Pull Placeholders</strong>.</li><li><strong>Execute:</strong> Fill the data, set Action to <strong>SEND</strong>, and click <strong>Run</strong>.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Learn about Action, To, Thread ID, and Attachments.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Action:</strong> Set to <code>SEND</code> for immediate mailing or <code>DRAFT</code> to create Gmail drafts for review.</li><li><strong>To:</strong> Primary recipient email address.</li><li><strong>Thread ID or Subject:</strong> Paste a Gmail Thread ID to reply to a specific conversation.</li><li><strong>Attachments:</strong> Comma-separated list of Google Drive File IDs or URLs.</li></ul>"
            },
            {
                icon: "lightbulb",
                color: "var(--warning)",
                label: "Pro Tips",
                shortDesc: "Case-sensitivity, formatting, and quota limits.",
                tooltipId: "help-tips",
                tooltipContent: "<p><strong>Pro Tips</strong></p><ul><li><strong>Case Sensitive:</strong> Tags like <code>{{Name}}</code> must match the column header exactly.</li><li><strong>HTML Support:</strong> Any formatting (bold, links, images) in your Gmail draft is preserved.</li><li><strong>Status:</strong> The tool updates the Status column automatically after each row.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getQuota: function () {
            return _App_ok('Quota loaded.', { remaining: MailApp.getRemainingDailyQuota() });
        },
        getGmailDrafts: function () {
            try {
                var drafts = GmailApp.getDrafts();
                var validDrafts = [];
                var regex = /\{\{[^{}]+\}\}/;

                for (var i = 0; i < drafts.length; i++) {
                    var msg = drafts[i].getMessage();
                    var subject = msg.getSubject() || "";
                    var body = msg.getBody() || "";

                    if (regex.test(subject) || regex.test(body)) {
                        validDrafts.push({
                            id: drafts[i].getId(),
                            subject: subject || "(No Subject)"
                        });

                        if (validDrafts.length >= 10) {
                            break;
                        }
                    }
                }
                return _App_ok('Drafts loaded.', { drafts: validDrafts });
            } catch (e) {
                return _App_ok('No drafts available.', { drafts: [] });
            }
        },
        syncPlaceholders: function (draftId) {
            if (!draftId) return _App_fail("No draft selected.");
            try {
                var draft = GmailApp.getDraft(draftId);
                if (!draft) throw new Error("Draft not found.");
                var msg = draft.getMessage();
                var subject = msg.getSubject() || "";
                var body = msg.getBody() || "";

                var placeholders = [];
                var regex = /\{\{([^{}]+)\}\}/g;

                var match;
                while ((match = regex.exec(subject)) !== null) {
                    if (placeholders.indexOf(match[1]) === -1) placeholders.push(match[1]);
                }
                while ((match = regex.exec(body)) !== null) {
                    if (placeholders.indexOf(match[1]) === -1) placeholders.push(match[1]);
                }

                var syncResult = SheetManager.syncDynamicColumns('MAIL_MERGE', placeholders, {
                    dynamicColWidth: 150
                });

                return _App_ok('Synced ' + placeholders.length + ' placeholders.', {
                    placeholders: placeholders,
                    headers: syncResult.headers
                });
            } catch (e) {
                var toolConfig = SyncEngine.getTool('MAIL_MERGE') || { TITLE: 'MAIL_MERGE' };
                Logger.error(toolConfig.TITLE, 'Sync Placeholders', e);
                return _App_fail("Sync failed: " + e.message + (e.stack ? "\nTrace:\n" + e.stack : ""));
            }
        },
        executeActions: function (draftId) {
            var pendingRows = SheetManager.readPendingObjects('MAIL_MERGE', { useDisplayValues: true });

            if (pendingRows.length === 0) {
                return _App_ok("Nothing to do! No 'SEND' or 'DRAFT' actions pending.");
            }

            var template = null;
            try {
                var draft = GmailApp.getDraft(draftId);
                if (!draft) throw new Error("Draft not found.");
                var msg = draft.getMessage();
                template = {
                    subject: msg.getSubject(),
                    body: msg.getBody(),
                    attachments: msg.getAttachments()
                };
            } catch (e) {
                throw new Error("⚠️ Failed to load Draft: " + e.message);
            }

            var headers = SheetManager.getHeaders('MAIL_MERGE');
            var compiledPlaceholders = [];
            for (var colIndex = 2; colIndex < headers.length; colIndex++) {
                var header = headers[colIndex];
                if (!header) continue;
                compiledPlaceholders.push({
                    header: header,
                    regex: new RegExp('{{' + _App_escapeRegExp(header) + '}}', 'g')
                });
            }

            var stats = _App_BatchProcessor('MAIL_MERGE', pendingRows, function (item) {
                var rowUpdates = {
                    action: item['Action'],
                    status: "",
                    _rowNumber: item._rowNumber
                };

                var action = rowUpdates.action.toString().trim().toUpperCase();
                if (action !== "SEND" && action !== "DRAFT") return null;

                var targetTo = item['To'];
                var targetCc = item['CC'];
                var targetBcc = item['BCC'];
                var targetThreadId = item['Thread ID or Subject'];
                var targetAttachments = item['Attachments'];

                if (!targetTo && !targetThreadId) throw new Error("Missing Email To");

                var emailBody = template.body;
                var emailSubject = template.subject;

                compiledPlaceholders.forEach(function (pDef) {
                    var value = item[pDef.header];
                    var valStr = (value === undefined || value === null || value === "") ? "" : String(value);
                    var bodyVal = valStr.replace(/\r?\n/g, '<br>');

                    emailBody = emailBody.replace(pDef.regex, () => bodyVal);
                    emailSubject = emailSubject.replace(pDef.regex, () => valStr);
                });

                var remainingPlaceholders = [];
                var unmatched;
                var regexExtract = /\{\{([^{}]+)\}\}/g;
                while ((unmatched = regexExtract.exec(emailBody)) !== null) {
                    remainingPlaceholders.push(unmatched[1]);
                }
                while ((unmatched = regexExtract.exec(emailSubject)) !== null) {
                    remainingPlaceholders.push(unmatched[1]);
                }
                var allRemaining = [...new Set(remainingPlaceholders)];
                if (allRemaining.length > 0) {
                    throw new Error("Missing columns for: " + allRemaining.join(', '));
                }

                var finalAttachments = [...template.attachments];
                if (targetAttachments) {
                    var files = targetAttachments.split(',');
                    for (var f = 0; f < files.length; f++) {
                        var blob = _App_getDriveAttachment(files[f].trim());
                        if (blob) finalAttachments.push(blob);
                    }
                }

                rowUpdates.status = _App_sendOrDraftEmail({
                    action: action,
                    to: targetTo,
                    cc: targetCc,
                    bcc: targetBcc,
                    subject: emailSubject,
                    body: emailBody,
                    attachments: finalAttachments,
                    threadIdOrSubject: targetThreadId
                });
                rowUpdates.action = "";

                Logger.info(SyncEngine.getTool('MAIL_MERGE').TITLE, 'Row ' + item._rowNumber, rowUpdates.status);
                return rowUpdates;
            }, {
                onBatchComplete: function (batchResults) {
                    _App_batchPatchResults('MAIL_MERGE', batchResults);
                }
            });

            var finalMsg = "Successfully processed " + stats.processedCount + " emails.";
            if (stats.errorCount > 0) finalMsg += " (" + stats.errorCount + " errors)";
            if (stats.timeLimitReached) finalMsg = "⏳ Time limit reached. " + finalMsg;

            return _App_ok(finalMsg);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Mail Merge sidebar and ensures the sheet exists. */
function MailMerge_openSidebar() {
  return Logger.run('MAIL_MERGE', 'Open Sidebar', function () {
    _App_launchTool('MAIL_MERGE');
  });
}

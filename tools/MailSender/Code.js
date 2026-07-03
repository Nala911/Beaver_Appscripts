/**
 * Mail Sender
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('MAIL_SENDER', {
    SHEET_NAME: SHEET_NAMES.MAIL_SENDER,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.MAIL_SENDER,
    MENU_LABEL: SHEET_NAMES.MAIL_SENDER,
    MENU_ENTRYPOINT: 'MailSender_openSidebar',
    MENU_ORDER: 40,
    SIDEBAR_HTML: 'tools/MailSender/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['SEND', 'DRAFT'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'To', type: 'EMAIL_LIST' },
            { header: 'CC', type: 'EMAIL_LIST' },
            { header: 'BCC', type: 'EMAIL_LIST' },
            { header: 'Thread ID or Subject', type: 'TEXT' },
            { header: 'Attachments', type: 'TEXT' },
            { header: 'Email Subject', type: 'TEXT' },
            { header: 'Email Body', type: 'TEXT' },
            { header: 'PDF HTML', type: 'TEXT' },
            { header: 'PDF Name', type: 'TEXT' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to send custom emails:</p><ol><li><strong>Fill Content:</strong> Enter subjects and bodies directly into the sheet.</li><li><strong>Action:</strong> Set the Action column to <code>SEND</code> or <code>DRAFT</code>.</li><li><strong>Run:</strong> Click <strong>Send Custom Mail</strong> in the sidebar.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Learn about PDF HTML, Attachments, and BCC.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Email Subject / Body:</strong> The actual text sent to recipients.</li><li><strong>PDF HTML:</strong> Raw HTML that will be converted to a PDF attachment.</li><li><strong>PDF Name:</strong> The name of the generated PDF file.</li><li><strong>Attachments:</strong> Comma-separated Drive File IDs or URLs.</li></ul>"
            },
            {
                icon: "lightbulb",
                color: "var(--warning)",
                label: "Pro Tips",
                shortDesc: "HTML-to-PDF conversion and bulk sending.",
                tooltipId: "help-tips",
                tooltipContent: "<p><strong>Pro Tips</strong></p><ul><li><strong>HTML to PDF:</strong> Use <code>&lt;table&gt;</code>, <code>&lt;h1&gt;</code>, and inline CSS for professional PDF attachments.</li><li><strong>Bulk Sending:</strong> Ideal for unique, one-off messages that don't follow a fixed template.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getQuota: function () {
            var quota = MailApp.getRemainingDailyQuota();
            return _App_ok('Remaining quota: ' + quota, { remaining: quota });
        },
        executeActions: function () {
            var pendingRows = SheetManager.readPendingObjects('MAIL_SENDER', { useDisplayValues: true });

            if (pendingRows.length === 0) return _App_ok("Nothing to do! No 'SEND' or 'DRAFT' actions pending.");

            var stats = _App_BatchProcessor('MAIL_SENDER', pendingRows, function (item, index) {
                var rowUpdates = {
                    action: item['Action'],
                    _rowNumber: item._rowNumber
                };

                var action = rowUpdates.action.toString().trim().toUpperCase();
                if (action !== "SEND" && action !== "DRAFT") return null;

                var targetTo = item['To'];
                var targetCc = item['CC'];
                var targetBcc = item['BCC'];
                var targetThreadId = item['Thread ID or Subject'];
                var targetAttachments = item['Attachments'];
                var targetPdfHtml = item['PDF HTML'];
                var targetPdfName = item['PDF Name'];

                if (!targetTo && !targetThreadId) throw new Error("⚠️ Missing Email To");

                var emailSubject = item['Email Subject'];
                var emailBody = item['Email Body'] ? String(item['Email Body']).replace(/\r?\n/g, '<br>') : "";

                if (!emailSubject && !targetThreadId) throw new Error("⚠️ Missing Email Subject");
                if (!emailBody) throw new Error("⚠️ Missing Email Body");

                var finalAttachments = [];
                if (targetAttachments) {
                    var files = targetAttachments.split(',');
                    for (var f = 0; f < files.length; f++) {
                        var blob = _App_getDriveAttachment(files[f].trim());
                        if (blob) finalAttachments.push(blob);
                    }
                }

                if (targetPdfHtml) {
                    var defaultFileName = "document.pdf";
                    var fileName = targetPdfName ? targetPdfName.toString().trim() : defaultFileName;
                    if (!fileName.toLowerCase().endsWith(".pdf")) {
                        fileName += ".pdf";
                    }
                    var pdfBlob = Utilities.newBlob(targetPdfHtml, 'text/html', fileName).getAs('application/pdf');
                    finalAttachments.push(pdfBlob);
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
                return rowUpdates;

            }, {
                onBatchComplete: function (batchResults) {
                    _App_batchPatchResults('MAIL_SENDER', batchResults);
                }
            });

            var finalResult = stats.processedCount + " actions processed!";
            return _App_ok(finalResult);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Mail Sender sidebar and ensures the sheet exists. */
function MailSender_openSidebar() {
  return Logger.run('MAIL_SENDER', 'Open Sidebar', function () {
    _App_launchTool('MAIL_SENDER');
  });
}

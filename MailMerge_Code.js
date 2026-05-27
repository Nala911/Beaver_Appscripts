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
    SIDEBAR_HTML: 'MailMerge_Sidebar',
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
            { header: 'To', type: 'TEXT' },
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
    }
});



/** Opens the Mail Merge sidebar and ensures the sheet exists. */
function MailMerge_openSidebar() {
  return Logger.run('MAIL_MERGE', 'Open Sidebar', function () {
    _App_launchTool('MAIL_MERGE');
  });
}


function MailMerge_getQuota() {
  return Logger.run('MAIL_MERGE', 'Get Quota', function () {
    return _App_ok('Quota loaded.', { remaining: MailApp.getRemainingDailyQuota() });
  });
}

function MailMerge_getGmailDrafts() {
  return Logger.run('MAIL_MERGE', 'Get Gmail Drafts', function () {
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
  });
}

function MailMerge_syncPlaceholders(draftId) {
  return Logger.run('MAIL_MERGE', 'Sync Placeholders', function () {
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
  });
}

function _MailMerge_escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function _MailMerge_getDriveAttachment(fileIdOrUrl) {
  try {
    if (!fileIdOrUrl) return null;
    var fileId = fileIdOrUrl;
    // Extract ID if URL is provided
    var match = fileIdOrUrl.match(/[-\w]{25,}/);
    if (match) fileId = match[0];

    var file = DriveApp.getFileById(fileId);
    return file.getBlob();
  } catch (e) {
    throw new Error("Cannot find attachment in Drive (" + fileIdOrUrl + ")" + (e.stack ? "\nTrace:\n" + e.stack : ""));
  }
}
// Centralized validators from 04_Core_Validators are used instead.

function _MailMerge_mergeEmails(existingStr, newStr) {
  if (!newStr) return existingStr || "";
  var existingArr = (existingStr || "").split(',').map(function (e) { return e.trim(); }).filter(function (e) { return e; });
  var newArr = (newStr || "").split(',').map(function (e) { return e.trim(); }).filter(function (e) { return e; });
  newArr.forEach(function (em) {
    if (existingArr.indexOf(em) === -1) {
      existingArr.push(em);
    }
  });
  return existingArr.join(',');
}

function MailMerge_executeActions(draftId, startIndex) {
  return Logger.run('MAIL_MERGE', 'Execute Actions', function () {
    var start = startIndex || 0;
    var batchSize = 10; 

    var pendingRows = SheetManager.readPendingObjects('MAIL_MERGE', { useDisplayValues: true });

    if (pendingRows.length === 0) return _App_ok(start > 0 ? "Batch finished!" : "Nothing to do! No 'SEND' or 'DRAFT' actions pending.", { completed: true, message: start > 0 ? "Batch finished!" : "Nothing to do! No 'SEND' or 'DRAFT' actions pending." });
    if (start >= pendingRows.length) return _App_ok("Batch complete!", { completed: true, message: "Batch complete!" });

    var batchItems = pendingRows.slice(start, start + batchSize);
    var remainingPending = pendingRows.length - (start + batchItems.length);

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

    var stats = _App_BatchProcessor('MAIL_MERGE', batchItems, function (item) {
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
      if (targetTo && !_App_validateEmailList(targetTo)) throw new Error("Invalid Email To address");
      if (!_App_validateEmailList(targetCc)) throw new Error("Invalid CC address");
      if (!_App_validateEmailList(targetBcc)) throw new Error("Invalid BCC address");

      var emailBody = template.body;
      var emailSubject = template.subject;

      // Headers for dynamic placeholders
      var headers = SheetManager.getHeaders('MAIL_MERGE');

      for (var colIndex = 2; colIndex < headers.length; colIndex++) {
        var header = headers[colIndex];
        if (!header) continue;
        var safeHeader = _MailMerge_escapeRegExp(header);
        var placeholder = new RegExp('{{' + safeHeader + '}}', 'g');
        var value = item[header];
        var valStr = (value === undefined || value === null || value === "") ? "" : String(value);
        var bodyVal = valStr.replace(/\r?\n/g, '<br>');

        emailBody = emailBody.replace(placeholder, () => bodyVal);
        emailSubject = emailSubject.replace(placeholder, () => valStr);
      }

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
          var blob = _MailMerge_getDriveAttachment(files[f].trim());
          if (blob) finalAttachments.push(blob);
        }
      }

      if (action === "SEND") {
        var options = {
          htmlBody: emailBody,
          attachments: finalAttachments
        };

        if (targetThreadId) {
          var thread = null;
          try { thread = GmailApp.getThreadById(targetThreadId); } catch (ignore) { }

          if (!thread) {
            var safeSubject = targetThreadId.toString().replace(/['"]/g, '');
            var query = 'subject:("' + safeSubject + '")';
            var threads = GmailApp.search(query, 0, 1);
            if (threads && threads.length > 0) thread = threads[0];
          }
          if (!thread) throw new Error("Thread not found for ID or Subject");

          var messages = thread.getMessages();
          var lastMessage = messages[messages.length - 1];

          var existingTo = lastMessage.getTo();
          var existingCc = lastMessage.getCc();

          var newTo = _MailMerge_mergeEmails(existingTo, targetTo);
          var newCc = _MailMerge_mergeEmails(existingCc, targetCc);

          var replyOptions = {
            htmlBody: emailBody,
            attachments: finalAttachments,
            cc: newCc || "",
            bcc: targetBcc || ""
          };

          var draftReply = lastMessage.createDraftReplyAll("", replyOptions);
          draftReply.update(newTo || "", emailSubject, "", replyOptions);
          draftReply.send();
        } else {
          options.cc = targetCc;
          options.bcc = targetBcc;
          GmailApp.sendEmail(targetTo, emailSubject, "", options);
        }

        rowUpdates.status = _App_formatStatus('SUCCESS', "Sent (" + new Date().toLocaleString() + ")");
        rowUpdates.action = "";
      } else if (action === "DRAFT") {
        var options = {
          htmlBody: emailBody,
          attachments: finalAttachments
        };

        if (targetThreadId) {
          var thread = null;
          try { thread = GmailApp.getThreadById(targetThreadId); } catch (ignore) { }

          if (!thread) {
            var safeSubject = targetThreadId.toString().replace(/['"]/g, '');
            var query = 'subject:("' + safeSubject + '")';
            var threads = GmailApp.search(query, 0, 1);
            if (threads && threads.length > 0) thread = threads[0];
          }
          if (!thread) throw new Error("Thread not found for ID or Subject");

          var threadMessages = thread.getMessages();
          var lastMessage = threadMessages[threadMessages.length - 1];

          var existingTo = lastMessage.getTo();
          var existingCc = lastMessage.getCc();

          var newTo = _MailMerge_mergeEmails(existingTo, targetTo);
          var newCc = _MailMerge_mergeEmails(existingCc, targetCc);

          var replyOptions = {
            htmlBody: emailBody,
            attachments: finalAttachments,
            cc: newCc || "",
            bcc: targetBcc || ""
          };

          var draftReply = lastMessage.createDraftReplyAll("", replyOptions);
          draftReply.update(newTo || "", emailSubject, "", replyOptions);

          rowUpdates.status = _App_formatStatus('SUCCESS', "Reply Draft Created");
          rowUpdates.action = "";
        } else {
          options.cc = targetCc;
          options.bcc = targetBcc;
          GmailApp.createDraft(targetTo, emailSubject, "", options);
          rowUpdates.status = _App_formatStatus('SUCCESS', "Draft Created");
          rowUpdates.action = "";
        }
      }

      Logger.info(SyncEngine.getTool('MAIL_MERGE').TITLE, 'Row ' + item._rowNumber, rowUpdates.status);
      return rowUpdates;
    }, {
      onBatchComplete: function (batchResults) {
        _App_batchPatchResults('MAIL_MERGE', batchResults);
      }
    });

    return _App_ok('Processed mail merge batch.', {
      completed: stats.processedCount + stats.errorCount >= pendingRows.length - start,
      nextIndex: start + batchItems.length,
      remainingPending: remainingPending,
      processed: stats.processedCount
    });
  });
}

function MailMerge_getRemainingPendingCount() {
  return Logger.run('MAIL_MERGE', 'Get Pending Count', function () {
    return SheetManager.readPendingObjects('MAIL_MERGE').length;
  });
}

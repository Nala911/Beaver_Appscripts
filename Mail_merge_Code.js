/**
 * Mail Merge Toolkit
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('MAIL_MERGE', {
    SHEET_NAME: SHEET_NAMES.MAIL_MERGE,
    TITLE: '📧 Mail Merge Toolkit',
    MENU_LABEL: '📧 Mail Merge System',
    MENU_ENTRYPOINT: 'MailMerge_openSidebar',
    MENU_ORDER: 30,
    SIDEBAR_HTML: 'Mail_merge_HTML',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 1,
    COL_WIDTHS: [120, 200, 150, 150, 150, 250, 200],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 1,
        conditionalRules: [
            { type: 'pending', actionCol: 'A', scope: 'actionOnly' },
            { type: 'success', statusCol: 'G', scope: 'fullRow' },
            { type: 'error', statusCol: 'G', scope: 'fullRow' }
        ],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['SEND', 'DRAFT'] },
            { header: 'To', type: 'EMAIL_LIST' },
            { header: 'CC', type: 'EMAIL_LIST' },
            { header: 'BCC', type: 'EMAIL_LIST' },
            { header: 'Thread ID or Subject', type: 'TEXT' },
            { header: 'Attachments', type: 'LIST' },
            { header: 'Status', type: 'TEXT' }
        ]
    }
});

// Column-index aliases kept for backward compatibility within this file.
// Metadata (title, sidebar, headers, widths) now lives in SyncEngine.getTool('MAIL_MERGE').
var MAILMERGE_CFG = {
  COLUMNS: {
    ACTION: 0, EMAIL_TO: 1, CC: 2, BCC: 3, THREAD_ID: 4, ATTACHMENTS: 5, STATUS: 6
  },
  HEADER_ROW: 1
};

/** @deprecated — Use _App_ensureSheetExists('MAIL_MERGE') instead. */
function _MailMerge_ensureSheetExistsAndActivate() {
  return _App_ensureSheetExists('MAIL_MERGE');
}

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
        anchorHeader: 'Status',
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

function _MailMerge_validateEmails(emailsString) {
  if (!emailsString) return true; // Empty is fine for CC/BCC
  var emails = emailsString.split(',');
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i].trim();
    if (email && !emailRegex.test(email)) {
      return false;
    }
  }
  return true;
}

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
    var draft = GmailApp.getDraft(draftId);
    if (!draft) throw new Error("Draft not found.");
    var templateMsg = draft.getMessage();
    var templateSubject = templateMsg.getSubject();
    var templateBody = templateMsg.getBody();
    var templateAttachments = templateMsg.getAttachments();

    var startTime = Date.now();
    var maxExecutionTime = 300 * 1000; // 5 minutes limit per batch execution

    var toolCfg = SyncEngine.getTool('MAIL_MERGE');
    var headers = toolCfg.HEADERS;

    // Use ExecutionService for robust row processing
    var stats = ExecutionService.processPendingRows('MAIL_MERGE', function(rowObj) {
        if (Date.now() - startTime > maxExecutionTime) throw new Error("⏳ Time limit reached.");

        var action = String(rowObj['Action'] || '').toUpperCase();
        if (action !== "SEND" && action !== "DRAFT") return;

        var targetTo = rowObj['To'];
        var targetCc = rowObj['CC'];
        var targetBcc = rowObj['BCC'];
        var targetThreadId = rowObj['Thread ID or Subject'];
        var targetAttachments = rowObj['Attachments'];

        if (!targetTo && !targetThreadId) throw new Error("Missing Email To");

        // Merge logic
        var emailBody = templateBody;
        var emailSubject = templateSubject;

        headers.forEach(function(header) {
          if (!header) return;
          var safeHeader = _MailMerge_escapeRegExp(header);
          var placeholder = new RegExp('{{' + safeHeader + '}}', 'g');
          var value = rowObj[header];
          var valStr = (value === undefined || value === null || value === "") ? "" : String(value);
          
          emailSubject = emailSubject.replace(placeholder, valStr);
          emailBody = emailBody.replace(placeholder, valStr.replace(/\r?\n/g, '<br>'));
        });

        // Placeholder validation
        var remainingPlaceholders = [];
        var unmatched;
        var regexExtract = /\{\{([^{}]+)\}\}/g;
        while ((unmatched = regexExtract.exec(emailBody)) !== null) remainingPlaceholders.push(unmatched[1]);
        while ((unmatched = regexExtract.exec(emailSubject)) !== null) remainingPlaceholders.push(unmatched[1]);
        
        var allRemaining = [...new Set(remainingPlaceholders)];
        if (allRemaining.length > 0) throw new Error("Missing columns for: " + allRemaining.join(', '));

        // Attachments
        var finalAttachments = [...templateAttachments];
        if (targetAttachments) {
          var files = targetAttachments.split(',');
          for (var f = 0; f < files.length; f++) {
            var blob = _MailMerge_getDriveAttachment(files[f].trim());
            if (blob) finalAttachments.push(blob);
          }
        }

        if (action === "SEND") {
          var options = { htmlBody: emailBody, attachments: finalAttachments };
          if (targetThreadId) {
            var thread = null;
            try { thread = GmailApp.getThreadById(targetThreadId); } catch (e) {}
            if (!thread) {
              var query = 'subject:("' + targetThreadId.toString().replace(/['"]/g, '') + '")';
              var threads = GmailApp.search(query, 0, 1);
              if (threads && threads.length > 0) thread = threads[0];
            }
            if (!thread) throw new Error("Thread not found for ID or Subject");

            var messages = thread.getMessages();
            var lastMessage = messages[messages.length - 1];

            var replyOptions = {
              htmlBody: emailBody,
              attachments: finalAttachments,
              cc: _MailMerge_mergeEmails(lastMessage.getCc(), targetCc) || "",
              bcc: targetBcc || ""
            };

            var draftReply = lastMessage.createDraftReplyAll("", replyOptions);
            draftReply.update(_MailMerge_mergeEmails(lastMessage.getTo(), targetTo) || "", emailSubject, "", replyOptions);
            draftReply.send();

          } else {
            options.cc = targetCc;
            options.bcc = targetBcc;
            GmailApp.sendEmail(targetTo, emailSubject, "", options);
          }

          SheetManager.patchRow('MAIL_MERGE', rowObj._rowNumber, {
              'Action': '',
              'Status': "✅ Sent (" + new Date().toLocaleString() + ")"
          });

        } else if (action === "DRAFT") {
          var draftOptions = { htmlBody: emailBody, attachments: finalAttachments };
          if (targetThreadId) {
            var threadD = null;
            try { threadD = GmailApp.getThreadById(targetThreadId); } catch (e) {}
            if (!threadD) {
                var queryD = 'subject:("' + targetThreadId.toString().replace(/['"]/g, '') + '")';
                var threadsD = GmailApp.search(queryD, 0, 1);
                if (threadsD && threadsD.length > 0) threadD = threadsD[0];
            }
            if (!threadD) throw new Error("Thread not found for ID or Subject");

            var tMsgs = threadD.getMessages();
            var lMsg = tMsgs[tMsgs.length - 1];

            var rOptions = {
              htmlBody: emailBody,
              attachments: finalAttachments,
              cc: _MailMerge_mergeEmails(lMsg.getCc(), targetCc) || "",
              bcc: targetBcc || ""
            };

            var dReply = lMsg.createDraftReplyAll("", rOptions);
            dReply.update(_MailMerge_mergeEmails(lMsg.getTo(), targetTo) || "", emailSubject, "", rOptions);

            SheetManager.patchRow('MAIL_MERGE', rowObj._rowNumber, {
              'Action': '',
              'Status': "📝 Reply Draft Created"
            });
          } else {
            draftOptions.cc = targetCc;
            draftOptions.bcc = targetBcc;
            GmailApp.createDraft(targetTo, emailSubject, "", draftOptions);
            SheetManager.patchRow('MAIL_MERGE', rowObj._rowNumber, {
              'Action': '',
              'Status': "📝 Draft Created"
            });
          }
        }
    }, { limit: 10 }); // Keep UI responsiveness by processing in small chunks

    return _App_ok('Processed mail merge batch.', {
      completed: stats.processed === 0 && stats.errors === 0,
      processed: stats.processed,
      errors: stats.errors
    });
  });
}

function MailMerge_getRemainingPendingCount() {
  return Logger.run('MAIL_MERGE', 'Get Pending Count', function () {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MAIL_MERGE);
    if (!sheet) return 0;
    var maxRows = sheet.getMaxRows();
    if (maxRows < 2) return 0;
    var actionRange = sheet.getRange(2, MAILMERGE_CFG.COLUMNS.ACTION + 1, maxRows - 1);
    var values = actionRange.getValues();
    var count = 0;
    for (var i = 0; i < values.length; i++) {
      var act = String(values[i][0]).toUpperCase();
      if (act === "SEND" || act === "DRAFT") {
        count++;
      }
    }
    return count;
  });
}

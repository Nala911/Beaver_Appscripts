/**
 * Mail Merge Toolkit
 * Version: 5.0 (Plugin Architecture — registers with BeaverEngine)
 */

BeaverEngine.registerTool('MAIL_MERGE', {
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
            { header: 'To', type: 'TEXT' },
            { header: 'CC', type: 'EMAIL_LIST' },
            { header: 'BCC', type: 'EMAIL_LIST' },
            { header: 'Thread ID or Subject', type: 'TEXT' },
            { header: 'Attachments', type: 'TEXT' },
            { header: 'Status', type: 'TEXT' }
        ]
    }
});

// Column-index aliases kept for backward compatibility within this file.
// Metadata (title, sidebar, headers, widths) now lives in BeaverEngine.getTool('MAIL_MERGE').
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
  _App_launchTool('MAIL_MERGE');
}


function MailMerge_getQuota() {
  return _App_ok('Quota loaded.', { remaining: MailApp.getRemainingDailyQuota() });
}

function MailMerge_getGmailDrafts() {
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
}

function MailMerge_syncPlaceholders(draftId) {
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
    return _App_fail("Sync failed: " + e.message + (e.stack ? "\nTrace:\n" + e.stack : ""));
  }
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
    var sheet = _App_assertActiveSheet(SHEET_NAMES.MAIL_MERGE);

    var start = startIndex || 0;
    var batchSize = 10; 

    var dataRange = sheet.getDataRange();
    var data = dataRange.getDisplayValues();

    if (data.length < 2) return _App_ok("Sheet is empty.", { completed: true, message: "Sheet is empty." });

    var headers = data[0];
    var allRows = data.slice(1);

    var pendingRows = [];
    allRows.forEach(function (row, idx) {
      if (row[MAILMERGE_CFG.COLUMNS.ACTION] === "SEND" || row[MAILMERGE_CFG.COLUMNS.ACTION] === "DRAFT") {
        pendingRows.push({ data: row, originalIndex: idx });
      }
    });

    if (pendingRows.length === 0) return _App_ok(start > 0 ? "Batch finished!" : "Nothing to do! No 'SEND' or 'DRAFT' actions pending.", { completed: true, message: start > 0 ? "Batch finished!" : "Nothing to do! No 'SEND' or 'DRAFT' actions pending." });
    if (start >= pendingRows.length) return _App_ok("Batch complete!", { completed: true, message: "Batch complete!" });

    var batch = pendingRows.slice(start, start + batchSize);
    var remainingPending = pendingRows.length - (start + batch.length);

    var templateMsg, templateSubject, templateBody, templateAttachments = [];

    try {
      var draft = GmailApp.getDraft(draftId);
      if (!draft) throw new Error("Draft not found.");
      templateMsg = draft.getMessage();
      templateSubject = templateMsg.getSubject();
      templateBody = templateMsg.getBody();
      templateAttachments = templateMsg.getAttachments();
    } catch (e) {
      throw new Error("⚠️ Failed to load Draft: " + e.message);
    }

    var updatesCount = 0;

    var updatesOriginalIndices = [];
    var batchUpdates = batch.map(function (item) {
      var row = item.data;
      var originalIdx = item.originalIndex;
      updatesOriginalIndices.push(originalIdx);

      var rowUpdates = {
        action: row[MAILMERGE_CFG.COLUMNS.ACTION],
        status: ""
      };

      if (!rowUpdates.action) return null;

      var action = rowUpdates.action.toString().trim().toUpperCase();
      if (action !== "SEND" && action !== "DRAFT") return null;

      try {
        var targetTo = row[MAILMERGE_CFG.COLUMNS.EMAIL_TO];
        var targetCc = row[MAILMERGE_CFG.COLUMNS.CC];
        var targetBcc = row[MAILMERGE_CFG.COLUMNS.BCC];
        var targetThreadId = row[MAILMERGE_CFG.COLUMNS.THREAD_ID];
        var targetAttachments = row[MAILMERGE_CFG.COLUMNS.ATTACHMENTS];

        if (!targetTo && !targetThreadId) throw new Error("Missing Email To");
        if (targetTo && !_MailMerge_validateEmails(targetTo)) throw new Error("Invalid Email To address");
        if (!_MailMerge_validateEmails(targetCc)) throw new Error("Invalid CC address");
        if (!_MailMerge_validateEmails(targetBcc)) throw new Error("Invalid BCC address");

        var emailBody = templateBody;
        var emailSubject = templateSubject;

        for (var colIndex = 6; colIndex < headers.length; colIndex++) {
          var header = headers[colIndex];
          if (!header) continue;
          var safeHeader = _MailMerge_escapeRegExp(header);
          var placeholder = new RegExp('{{' + safeHeader + '}}', 'g');
          var value = row[colIndex];
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

        var finalAttachments = [...templateAttachments];
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
            try {
              thread = GmailApp.getThreadById(targetThreadId);
            } catch (ignore) { }

            if (!thread) {
              var safeSubject = targetThreadId.toString().replace(/['"]/g, '');
              var query = 'subject:("' + safeSubject + '")';
              var threads = GmailApp.search(query, 0, 1);
              if (threads && threads.length > 0) {
                thread = threads[0];
              }
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

          rowUpdates.status = "✅ Sent (" + new Date().toLocaleString() + ")";
          rowUpdates.action = "";
        } else if (action === "DRAFT") {
          var options = {
            htmlBody: emailBody,
            attachments: finalAttachments
          };

          if (targetThreadId) {
            var thread = null;
            try {
              thread = GmailApp.getThreadById(targetThreadId);
            } catch (ignore) { }

            if (!thread) {
              var safeSubject = targetThreadId.toString().replace(/['"]/g, '');
              var query = 'subject:("' + safeSubject + '")';
              var threads = GmailApp.search(query, 0, 1);
              if (threads && threads.length > 0) {
                thread = threads[0];
              }
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

            rowUpdates.status = "📝 Reply Draft Created";
            rowUpdates.action = "";
          } else {
            options.cc = targetCc;
            options.bcc = targetBcc;
            GmailApp.createDraft(targetTo, emailSubject, "", options);
            rowUpdates.status = "📝 Draft Created";
            rowUpdates.action = "";
          }
        }

        updatesCount++;

      } catch (e) {
        rowUpdates.status = e.message;
        Logger.error(BeaverEngine.getTool('MAIL_MERGE').TITLE, 'Row ' + (originalIdx + 2), e);
      }

      if (action === "SEND" || action === "DRAFT") {
          var isSuccess = rowUpdates.status.indexOf('✅') > -1 || rowUpdates.status.indexOf('📝') > -1;
          if (isSuccess) {
              Logger.info(BeaverEngine.getTool('MAIL_MERGE').TITLE, 'Row ' + (originalIdx + 2), rowUpdates.status);
          }
      }

      return rowUpdates;
    });

    batchUpdates.forEach(function (u, i) {
      if (u) {
        var rowNum = updatesOriginalIndices[i] + 2;
        sheet.getRange(rowNum, MAILMERGE_CFG.COLUMNS.ACTION + 1).setValue(u.action);
        sheet.getRange(rowNum, MAILMERGE_CFG.COLUMNS.STATUS + 1).setValue(u.status);
      }
    });

    return _App_ok('Processed mail merge batch.', {
      completed: false,
      nextIndex: start + batchUpdates.length,
      remainingPending: remainingPending,
      processed: batchUpdates.length
    });
  });
}

function MailMerge_getRemainingPendingCount() {
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
}

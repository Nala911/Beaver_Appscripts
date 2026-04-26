/**
 * Mail Merge
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('MAIL_MERGE', {
    SHEET_NAME: SHEET_NAMES.MAIL_MERGE,
    TITLE: SHEET_NAMES.MAIL_MERGE,
    MENU_LABEL: SHEET_NAMES.MAIL_MERGE,
    MENU_ENTRYPOINT: 'MailMerge_openSidebar',
    MENU_ORDER: 30,
    SIDEBAR_HTML: 'MailMerge_Sidebar',
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

      try {
        var targetTo = item['To'];
        var targetCc = item['CC'];
        var targetBcc = item['BCC'];
        var targetThreadId = item['Thread ID or Subject'];
        var targetAttachments = item['Attachments'];

        if (!targetTo && !targetThreadId) throw new Error("Missing Email To");
        if (targetTo && !_MailMerge_validateEmails(targetTo)) throw new Error("Invalid Email To address");
        if (!_MailMerge_validateEmails(targetCc)) throw new Error("Invalid CC address");
        if (!_MailMerge_validateEmails(targetBcc)) throw new Error("Invalid BCC address");

        var emailBody = template.body;
        var emailSubject = template.subject;

        // Headers for dynamic placeholders
        var headers = SheetManager.getHeaders('MAIL_MERGE');

        for (var colIndex = 6; colIndex < headers.length; colIndex++) {
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

          rowUpdates.status = SHEET_THEME.STATUS_PREFIXES.SUCCESS + "Sent (" + new Date().toLocaleString() + ")";
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

            rowUpdates.status = SHEET_THEME.STATUS_PREFIXES.SUCCESS + "Reply Draft Created";
            rowUpdates.action = "";
          } else {
            options.cc = targetCc;
            options.bcc = targetBcc;
            GmailApp.createDraft(targetTo, emailSubject, "", options);
            rowUpdates.status = SHEET_THEME.STATUS_PREFIXES.SUCCESS + "Draft Created";
            rowUpdates.action = "";
          }
        }

        Logger.info(SyncEngine.getTool('MAIL_MERGE').TITLE, 'Row ' + item._rowNumber, rowUpdates.status);
        return rowUpdates;

      } catch (e) {
        rowUpdates.status = e.message;
        Logger.error(SyncEngine.getTool('MAIL_MERGE').TITLE, 'Row ' + item._rowNumber, e);
        return rowUpdates;
      }
    }, {
      onBatchComplete: function (batchResults) {
        var rowNumbers = [];
        var updates = [];
        var prefixes = SHEET_THEME.STATUS_PREFIXES;
        
        batchResults.forEach(function (res) {
          if (res && res._rowNumber !== undefined) {
            rowNumbers.push(res._rowNumber);
            if (res.isError) {
              updates.push({ 'Action': res.action, 'Status': prefixes.ERROR + res.error });
            } else {
              updates.push({ 'Action': res.action, 'Status': res.status });
            }
          }
        });
        if (rowNumbers.length > 0) {
          SheetManager.batchPatchRows('MAIL_MERGE', rowNumbers, updates);
        }
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MAIL_MERGE);
    if (!sheet) return 0;
    var maxRows = sheet.getMaxRows();
    if (maxRows < 2) return 0;
    var headers = SheetManager.getHeaders('MAIL_MERGE');
    var actionColIndex = headers.indexOf('Action') + 1;
    if (actionColIndex < 1) return 0;
    
    var actionRange = sheet.getRange(2, actionColIndex, maxRows - 1);
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

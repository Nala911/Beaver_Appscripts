/**
 * Mail Sender
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('MAIL_SENDER', {
    SHEET_NAME: SHEET_NAMES.MAIL_SENDER,
    TITLE: SHEET_NAMES.MAIL_SENDER,
    MENU_LABEL: SHEET_NAMES.MAIL_SENDER,
    MENU_ENTRYPOINT: 'MailSender_openSidebar',
    MENU_ORDER: 40,
    SIDEBAR_HTML: 'MailSender_Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 1,
    COL_WIDTHS: [120, 200, 150, 150, 150, 250, 250, 300, 300, 200],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 0,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['SEND', 'DRAFT'] },
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
    }
});

// Column-index aliases kept for backward compatibility within this file.
// Metadata (title, sidebar, headers, widths) now lives in SyncEngine.getTool('MAIL_SENDER').
var MAIL_SENDER_CFG = {
  COLUMNS: {
    ACTION: 0, EMAIL_TO: 1, CC: 2, BCC: 3, THREAD_ID: 4,
    ATTACHMENTS: 5, EMAIL_SUBJECT: 6, EMAIL_BODY: 7, PDF_HTML: 8, PDF_NAME: 9
  },
  HEADER_ROW: 1
};

/** Opens the Mail Sender sidebar and ensures the sheet exists. */
function MailSender_openSidebar() {
  return Logger.run('MAIL_SENDER', 'Open Sidebar', function () {
    _App_launchTool('MAIL_SENDER');
  });
}

function MailSender_getQuota() {
  return Logger.run('MAIL_SENDER', 'Get Quota', function () {
    var quota = MailApp.getRemainingDailyQuota();
    return _App_ok('Remaining quota: ' + quota, quota);
  });
}



function _MailSender_escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function _MailSender_getDriveAttachment(fileIdOrUrl) {
  try {
    if (!fileIdOrUrl) return null;
    var fileId = fileIdOrUrl;
    // Extract ID if URL is provided
    var match = fileIdOrUrl.match(/[-\w]{25,}/);
    if (match) fileId = match[0];

    var file = DriveApp.getFileById(fileId);
    return file.getBlob();
  } catch (e) {
    throw new Error("Cannot find attachment in Drive (" + fileIdOrUrl + ")");
  }
}

function _MailSender_validateEmails(emailsString) {
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

function _MailSender_mergeEmails(existingStr, newStr) {
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

function MailSender_executeActions() {
  return Logger.run('MAIL_SENDER', 'Execute Actions', function () {
    var pendingRows = SheetManager.readPendingObjects('MAIL_SENDER', { useDisplayValues: true });

    if (pendingRows.length === 0) return _App_ok("Nothing to do! No 'SEND' or 'DRAFT' actions pending.");

    var stats = _App_BatchProcessor('MAIL_SENDER', pendingRows, function (item, index) {
      var rowUpdates = {
        action: item['Action'],
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
        var targetPdfHtml = item['PDF HTML'];
        var targetPdfName = item['PDF Name'];

        if (!targetTo && !targetThreadId) throw new Error("⚠️ Missing Email To");
        if (targetTo && !_MailSender_validateEmails(targetTo)) throw new Error("⚠️ Invalid Email To address");
        if (!_MailSender_validateEmails(targetCc)) throw new Error("⚠️ Invalid CC address");
        if (!_MailSender_validateEmails(targetBcc)) throw new Error("⚠️ Invalid BCC address");

        var emailSubject = item['Email Subject'];
        var emailBody = item['Email Body'] ? String(item['Email Body']).replace(/\r?\n/g, '<br>') : "";

        if (!emailSubject && !targetThreadId) throw new Error("⚠️ Missing Email Subject");
        if (!emailBody) throw new Error("⚠️ Missing Email Body");

        var finalAttachments = [];
        if (targetAttachments) {
          var files = targetAttachments.split(',');
          for (var f = 0; f < files.length; f++) {
            var blob = _MailSender_getDriveAttachment(files[f].trim());
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

        var resultStatus = "";
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
            if (!thread) throw new Error("⚠️ Thread not found for ID or Subject");

            var messages = thread.getMessages();
            var lastMessage = messages[messages.length - 1];

            var existingTo = lastMessage.getTo();
            var existingCc = lastMessage.getCc();

            var newTo = _MailSender_mergeEmails(existingTo, targetTo);
            var newCc = _MailSender_mergeEmails(existingCc, targetCc);

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

          resultStatus = "✅ Sent (" + new Date().toLocaleString() + ")";
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
            if (!thread) throw new Error("⚠️ Thread not found for ID or Subject");

            var threadMessages = thread.getMessages();
            var lastMessage = threadMessages[threadMessages.length - 1];

            var existingTo = lastMessage.getTo();
            var existingCc = lastMessage.getCc();

            var newTo = _MailSender_mergeEmails(existingTo, targetTo);
            var newCc = _MailSender_mergeEmails(existingCc, targetCc);

            var replyOptions = {
              htmlBody: emailBody,
              attachments: finalAttachments,
              cc: newCc || "",
              bcc: targetBcc || ""
            };

            var draftReply = lastMessage.createDraftReplyAll("", replyOptions);
            draftReply.update(newTo || "", emailSubject, "", replyOptions);

            resultStatus = "📝 Reply Draft Created";
            rowUpdates.action = "";
          } else {
            options.cc = targetCc;
            options.bcc = targetBcc;
            GmailApp.createDraft(targetTo, emailSubject, "", options);
            resultStatus = "📝 Draft Created";
            rowUpdates.action = "";
          }
        }

        var reference = 'Row ' + item._rowNumber;
        if (emailSubject) reference += ' (' + emailSubject + ')';
        Logger.info(SyncEngine.getTool('MAIL_SENDER').TITLE, reference, resultStatus);

        return rowUpdates;

      } catch (e) {
        var reference = 'Row ' + item._rowNumber;
        Logger.error(SyncEngine.getTool('MAIL_SENDER').TITLE, reference, e);
        return null;
      }
    }, {
      onBatchComplete: function (batchResults) {
        var rowNumbers = [];
        var patchData = [];
        batchResults.forEach(function (res) {
          if (res && res._rowNumber !== undefined) {
            rowNumbers.push(res._rowNumber);
            patchData.push({ 'Action': res.action });
          }
        });
        if (rowNumbers.length > 0) {
          SheetManager.batchPatchRows('MAIL_SENDER', rowNumbers, patchData);
        }
      }
    });

    var finalResult = stats.processedCount + " actions processed!";
    Logger.info(SyncEngine.getTool('MAIL_SENDER').TITLE, "Execute Actions", finalResult);
    return _App_ok(finalResult);
  });
}

function MailSender_getProgress() {
  return Logger.run('MAIL_SENDER', 'Get Progress', function () {
    return _App_ok('Progress', _App_getProgress('MAIL_SENDER'));
  });
}

/**
 * Mail_Sender Toolkit
 * Version: 5.0 (Plugin Architecture — registers with BeaverEngine)
 */

BeaverEngine.registerTool('MAIL_SENDER', {
    SHEET_NAME: SHEET_NAMES.MAIL_SENDER,
    TITLE: '📩 Mail Sender Toolkit',
    MENU_LABEL: '📩 Mail Sender',
    MENU_ENTRYPOINT: 'Mail_Sender_openSidebar',
    MENU_ORDER: 40,
    SIDEBAR_HTML: 'Mail_Sender-Sidebar',
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
// Metadata (title, sidebar, headers, widths) now lives in APP_REGISTRY.MAIL_SENDER.
var MAIL_SENDER_CFG = {
  COLUMNS: {
    ACTION: 0, EMAIL_TO: 1, CC: 2, BCC: 3, THREAD_ID: 4,
    ATTACHMENTS: 5, EMAIL_SUBJECT: 6, EMAIL_BODY: 7, PDF_HTML: 8, PDF_NAME: 9
  },
  HEADER_ROW: 1
};

/** @deprecated — Use _App_ensureSheetExists('MAIL_SENDER') instead. */
function _Mail_Sender_ensureSheetExistsAndActivate() {
  return _App_ensureSheetExists('MAIL_SENDER');
}

/** Opens the Mail Sender sidebar and ensures the sheet exists. */
function Mail_Sender_openSidebar() {
  _App_launchTool('MAIL_SENDER');
}

function Mail_Sender_getQuota() {
  return MailApp.getRemainingDailyQuota();
}



function _Mail_Sender_escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function _Mail_Sender_getDriveAttachment(fileIdOrUrl) {
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

function _Mail_Sender_validateEmails(emailsString) {
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

function _Mail_Sender_mergeEmails(existingStr, newStr) {
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

function Mail_Sender_executeActions() {
  return Logger.run('MAIL_SENDER', 'Execute Actions', function () {
    var sheet = _App_assertActiveSheet(SHEET_NAMES.MAIL_SENDER);

    var dataRange = sheet.getDataRange();
    var data = dataRange.getDisplayValues();

    if (data.length < 2) return "Sheet is empty.";

    var headers = data[0];
    var allRows = data.slice(1);

    var validRowsCount = allRows.filter(function (row) {
      if (!row[MAIL_SENDER_CFG.COLUMNS.ACTION]) return false;
      var action = row[MAIL_SENDER_CFG.COLUMNS.ACTION].toString().trim().toUpperCase();
      return action === "SEND" || action === "DRAFT";
    }).length;

    var updatesCount = 0;

    _App_setProgress('MAIL_SENDER', 0, validRowsCount);

    var updates = allRows.map(function (row, index) {
      var rowUpdates = {
        action: row[MAIL_SENDER_CFG.COLUMNS.ACTION],
        status: ""
      };

      if (!rowUpdates.action) return null;

      var action = rowUpdates.action.toString().trim().toUpperCase();
      if (action !== "SEND" && action !== "DRAFT") return null;

      try {
        var targetTo = row[MAIL_SENDER_CFG.COLUMNS.EMAIL_TO];
        var targetCc = row[MAIL_SENDER_CFG.COLUMNS.CC];
        var targetBcc = row[MAIL_SENDER_CFG.COLUMNS.BCC];
        var targetThreadId = row[MAIL_SENDER_CFG.COLUMNS.THREAD_ID];
        var targetAttachments = row[MAIL_SENDER_CFG.COLUMNS.ATTACHMENTS];
        var targetPdfHtml = row[MAIL_SENDER_CFG.COLUMNS.PDF_HTML];
        var targetPdfName = row[MAIL_SENDER_CFG.COLUMNS.PDF_NAME];

        if (!targetTo && !targetThreadId) throw new Error("⚠️ Missing Email To");
        if (targetTo && !_Mail_Sender_validateEmails(targetTo)) throw new Error("⚠️ Invalid Email To address");
        if (!_Mail_Sender_validateEmails(targetCc)) throw new Error("⚠️ Invalid CC address");
        if (!_Mail_Sender_validateEmails(targetBcc)) throw new Error("⚠️ Invalid BCC address");

        var emailSubject = row[MAIL_SENDER_CFG.COLUMNS.EMAIL_SUBJECT];
        var emailBody = row[MAIL_SENDER_CFG.COLUMNS.EMAIL_BODY] ? String(row[MAIL_SENDER_CFG.COLUMNS.EMAIL_BODY]).replace(/\r?\n/g, '<br>') : "";

        if (!emailSubject && !targetThreadId) throw new Error("⚠️ Missing Email Subject");
        if (!emailBody) throw new Error("⚠️ Missing Email Body");

        var finalAttachments = [];
        if (targetAttachments) {
          var files = targetAttachments.split(',');
          for (var f = 0; f < files.length; f++) {
            var blob = _Mail_Sender_getDriveAttachment(files[f].trim());
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
            if (!thread) throw new Error("⚠️ Thread not found for ID or Subject");

            var messages = thread.getMessages();
            var lastMessage = messages[messages.length - 1];

            var existingTo = lastMessage.getTo();
            var existingCc = lastMessage.getCc();

            var newTo = _Mail_Sender_mergeEmails(existingTo, targetTo);
            var newCc = _Mail_Sender_mergeEmails(existingCc, targetCc);

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
            if (!thread) throw new Error("⚠️ Thread not found for ID or Subject");

            var threadMessages = thread.getMessages();
            var lastMessage = threadMessages[threadMessages.length - 1];

            var existingTo = lastMessage.getTo();
            var existingCc = lastMessage.getCc();

            var newTo = _Mail_Sender_mergeEmails(existingTo, targetTo);
            var newCc = _Mail_Sender_mergeEmails(existingCc, targetCc);

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

        var reference = 'Row ' + (index + 2);
        if (emailSubject) reference += ' (' + emailSubject + ')';
        Logger.info(APP_REGISTRY.MAIL_SENDER.TITLE, reference, rowUpdates.status);

        updatesCount++;
        _App_setProgress('MAIL_SENDER', updatesCount, validRowsCount);

      } catch (e) {
        rowUpdates.status = e.message;
        var reference = 'Row ' + (index + 2);
        Logger.error(APP_REGISTRY.MAIL_SENDER.TITLE, reference, e);
      }

      return rowUpdates;
    });

    var firstUpdatedRow = Infinity;
    var lastUpdatedRow = -Infinity;

    updates.forEach(function (u, i) {
      if (u) {
        var rowNum = i + 2;
        if (rowNum < firstUpdatedRow) firstUpdatedRow = rowNum;
        if (rowNum > lastUpdatedRow) lastUpdatedRow = rowNum;
      }
    });

    if (firstUpdatedRow <= lastUpdatedRow) {
      var rangeToUpdate = sheet.getRange(firstUpdatedRow, MAIL_SENDER_CFG.COLUMNS.ACTION + 1, lastUpdatedRow - firstUpdatedRow + 1, 1);
      var currentValues = rangeToUpdate.getValues();
      updates.forEach(function (u, i) {
        if (u) {
          var arrayIdx = (i + 2) - firstUpdatedRow;
          currentValues[arrayIdx][0] = u.action;
        }
      });
      rangeToUpdate.setValues(currentValues);
      }

      _App_clearProgress('MAIL_SENDER');

      var finalResult = updatesCount + " actions processed!";
      Logger.info("Mail Sender", "Execute Actions", finalResult);
    return finalResult;
  });
}

function Mail_Sender_getProgress() {
  return _App_getProgress('MAIL_SENDER');
}

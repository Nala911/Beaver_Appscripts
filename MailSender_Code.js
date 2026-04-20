/**
 * Mail_Sender Toolkit
 * Version: 5.0 (Plugin Architecture — registers with App.Engine)
 */

App.Engine.registerTool('MAIL_SENDER', {
    SHEET_NAME: SHEET_NAMES.MAIL_SENDER,
    TITLE: '📩 Mail Sender Toolkit',
    MENU_LABEL: '📩 Mail Sender',
    MENU_ENTRYPOINT: 'Mail_Sender_openSidebar',
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
    },

    /**
     * MAIL SENDER SERVICE ACTIONS
     */
    service: {
        getQuota: function() {
            return Logger.run('MAIL_SENDER', 'Get Quota', function () {
                return MailApp.getRemainingDailyQuota();
            });
        },

        executeActions: function() {
            return Mail_Sender_executeActions();
        },

        getProgress: function() {
            return _App_getProgress('MAIL_SENDER');
        }
    }
});


/** Opens the Mail Sender sidebar and ensures the sheet exists. */
function Mail_Sender_openSidebar() {
  return Logger.run('MAIL_SENDER', 'Open Sidebar', function () {
    _App_launchTool('MAIL_SENDER');
  });
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
    var stats = ExecutionService.processPendingRows('MAIL_SENDER', function(rowObj) {
      var action = String(rowObj['Action'] || '').trim().toUpperCase();
      if (action !== "SEND" && action !== "DRAFT") throw new Error("Invalid action: " + action);

      var targetTo = rowObj['To'];
      var targetCc = rowObj['CC'];
      var targetBcc = rowObj['BCC'];
      var targetThreadId = rowObj['Thread ID or Subject'];
      var targetAttachments = rowObj['Attachments'];
      var targetPdfHtml = rowObj['PDF HTML'];
      var targetPdfName = rowObj['PDF Name'];
      var emailSubject = rowObj['Email Subject'];
      var emailBody = rowObj['Email Body'] ? String(rowObj['Email Body']).replace(/\r?\n/g, '<br>') : "";

      if (!targetTo && !targetThreadId) throw new Error("Missing Email To");
      if (targetTo && !_Mail_Sender_validateEmails(targetTo)) throw new Error("Invalid Email To address");
      if (!_Mail_Sender_validateEmails(targetCc)) throw new Error("Invalid CC address");
      if (!_Mail_Sender_validateEmails(targetBcc)) throw new Error("Invalid BCC address");

      if (!emailSubject && !targetThreadId) throw new Error("Missing Email Subject");
      if (!emailBody) throw new Error("Missing Email Body");

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

      var rowStatus = "";

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

        rowStatus = "✅ Sent (" + new Date().toLocaleString() + ")";
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
          rowStatus = "📝 Reply Draft Created";
        } else {
          options.cc = targetCc;
          options.bcc = targetBcc;
          GmailApp.createDraft(targetTo, emailSubject, "", options);
          rowStatus = "📝 Draft Created";
        }
      }

      SheetManager.patchRow('MAIL_SENDER', rowObj._rowNumber, {
        'Action': '',
        'Log': rowStatus
      });
    });

    if (stats.processed === 0 && stats.errors === 0) {
      return { success: true, message: "No pending 'SEND' or 'DRAFT' actions found." };
    }

    return { success: true, message: "Processed " + stats.processed + " actions. (" + stats.errors + " errors)" };
  });
}

function Mail_Sender_getProgress() {
  return _App_getProgress('MAIL_SENDER');
}

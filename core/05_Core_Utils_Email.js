// ==========================================
// Centralized Email Utilities
// ==========================================

/**
 * Unified helper to send an email or create a draft.
 * Supports thread replies and file attachments.
 * 
 * @param {Object} options
 * @param {string} options.action - "SEND" or "DRAFT"
 * @param {string} options.to - Primary recipient(s)
 * @param {string} options.cc - CC recipient(s)
 * @param {string} options.bcc - BCC recipient(s)
 * @param {string} options.subject - Email subject
 * @param {string} options.body - Email HTML body
 * @param {Blob[]} options.attachments - Attachment blobs
 * @param {string} [options.threadIdOrSubject] - Thread ID or thread subject to reply to
 * @returns {string} User-friendly result status message
 */
function _App_sendOrDraftEmail(options) {
    var action = (options.action || "").toString().trim().toUpperCase();
    var to = options.to || "";
    var cc = options.cc || "";
    var bcc = options.bcc || "";
    var subject = options.subject || "";
    var body = options.body || "";
    var attachments = options.attachments || [];
    var threadIdOrSubject = options.threadIdOrSubject || "";

    if (!to && !threadIdOrSubject) {
        throw new Error("⚠️ Missing Email To");
    }
    if (!subject && !threadIdOrSubject) {
        throw new Error("⚠️ Missing Email Subject");
    }

    var mailOptions = {
        htmlBody: body,
        attachments: attachments
    };

    if (threadIdOrSubject) {
        var thread = null;
        try { 
            thread = GmailApp.getThreadById(threadIdOrSubject); 
        } catch (ignore) {}

        if (!thread) {
            var safeSubject = threadIdOrSubject.toString().replace(/['"]/g, '');
            var query = 'subject:("' + safeSubject + '")';
            var threads = GmailApp.search(query, 0, 1);
            if (threads && threads.length > 0) thread = threads[0];
        }
        if (!thread) {
            throw new Error("⚠️ Thread not found for ID or Subject");
        }

        var messages = thread.getMessages();
        var lastMessage = messages[messages.length - 1];

        var existingTo = lastMessage.getTo();
        var existingCc = lastMessage.getCc();

        var newTo = _App_mergeEmails(existingTo, to);
        var newCc = _App_mergeEmails(existingCc, cc);

        var replyOptions = {
            htmlBody: body,
            attachments: attachments,
            cc: newCc || "",
            bcc: bcc || ""
        };

        var draftReply = lastMessage.createDraftReplyAll("", replyOptions);
        draftReply.update(newTo || "", subject, "", replyOptions);

        if (action === "SEND") {
            draftReply.send();
            return _App_formatStatus('SUCCESS', "Sent (" + _App_formatDateTime(new Date()) + ")");
        } else {
            return _App_formatStatus('SUCCESS', "Reply Draft Created");
        }
    } else {
        mailOptions.cc = cc;
        mailOptions.bcc = bcc;

        if (action === "SEND") {
            GmailApp.sendEmail(to, subject, "", mailOptions);
            return _App_formatStatus('SUCCESS', "Sent (" + _App_formatDateTime(new Date()) + ")");
        } else {
            GmailApp.createDraft(to, subject, "", mailOptions);
            return _App_formatStatus('SUCCESS', "Draft Created");
        }
    }
}

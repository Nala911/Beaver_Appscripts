// ==========================================
// Centralized Data Validators
// ==========================================
var SYSTEM_VALIDATORS = {
    EMAIL: function(val) { return typeof val === 'string' && val.indexOf('@') !== -1; },
    DATE: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    DATETIME: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    BOOLEAN: function(val) {
        if (typeof val === 'boolean') return true;
        if (typeof val === 'string') {
            var lower = val.toLowerCase();
            return lower === 'true' || lower === 'false';
        }
        return false;
    },
    DOCS_URL: function(val) {
        if (val === '' || val === null || val === undefined) return true;
        return typeof val === 'string' && val.indexOf('docs.google.com/document') !== -1;
    },
    DRIVE_URL: function(val) {
        if (val === '' || val === null || val === undefined) return true;
        return typeof val === 'string' && (val.indexOf('drive.google.com') !== -1 || val.indexOf('docs.google.com') !== -1);
    }
};

/**
 * Validates a single email address using a robust regex.
 * @param {string} email
 * @returns {boolean}
 */
function _App_validateEmail(email) {
    if (!email || typeof email !== 'string') return false;
    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email.trim());
}

/**
 * Validates a comma-separated list of email addresses.
 * @param {string} emailsString
 * @param {boolean} [allowEmpty] - If true, an empty string is considered valid.
 * @returns {boolean}
 */
function _App_validateEmailList(emailsString, allowEmpty) {
    var shouldAllowEmpty = allowEmpty !== false; // default to true, since CC/BCC etc are optional
    var val = (emailsString || '').toString().trim();
    if (val === '') return shouldAllowEmpty;

    var emails = val.split(',');
    for (var i = 0; i < emails.length; i++) {
        var email = emails[i].trim();
        if (email && !_App_validateEmail(email)) {
            return false;
        }
    }
    return true;
}



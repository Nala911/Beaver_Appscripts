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

/**
 * Unifies column cell validation by checking its value against schema column type definitions.
 * @param {string} type - Column schema type (e.g. 'EMAIL', 'DATE', 'BOOLEAN', 'URL', 'DROPDOWN')
 * @param {*} value - Cell value to check
 * @param {Object} [fieldConfig] - The schema field configuration object (contains header, options, etc.)
 * @returns {boolean} True if value is valid, false otherwise.
 */
function _App_validateValueByType(type, value, fieldConfig) {
    var valStr = (value === null || value === undefined) ? '' : String(value).trim();
    
    // ACTION, STATUS, ID, and READ_ONLY do not require data-type validation or are handled natively
    if (type === 'ACTION' || type === 'STATUS' || type === 'ID' || type === 'READ_ONLY') {
        return true;
    }
    
    switch (type) {
        case 'EMAIL':
            if (valStr === '') return true; // Optional fields are empty
            return _App_validateEmail(valStr);
            
        case 'EMAIL_LIST':
            return _App_validateEmailList(valStr, true);
            
        case 'DATE':
        case 'DATETIME':
            if (valStr === '') return true;
            return SYSTEM_VALIDATORS.DATE(value);
            
        case 'BOOLEAN':
            if (valStr === '') return true;
            return SYSTEM_VALIDATORS.BOOLEAN(value);
            
        case 'URL':
            if (valStr === '') return true;
            // Simple match for URLs
            var urlRegex = /^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$/i;
            return urlRegex.test(valStr);
            
        case 'DOCS_URL':
            return SYSTEM_VALIDATORS.DOCS_URL(valStr);
            
        case 'DRIVE_URL':
            return SYSTEM_VALIDATORS.DRIVE_URL(valStr);
            
        case 'DROPDOWN':
            if (valStr === '') return true;
            if (fieldConfig && fieldConfig.allowInvalid) return true;
            if (fieldConfig && fieldConfig.options) {
                var opts = fieldConfig.options;
                if (typeof opts === 'function') {
                    try {
                        opts = opts();
                    } catch (e) {
                        return true; // If dynamic evaluation fails, pass validation or gracefully ignore
                    }
                }
                if (Array.isArray(opts)) {
                    var lowerVal = valStr.toLowerCase();
                    return opts.some(function(opt) {
                        return String(opt).trim().toLowerCase() === lowerVal;
                    });
                }
            }
            return true;
            
        case 'TEXT':
        default:
            return true; // Standard text has no structural restrictions
    }
}




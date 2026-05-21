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


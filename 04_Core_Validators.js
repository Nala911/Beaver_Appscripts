// ==========================================
// Centralized Data Validators
// ==========================================
var SYSTEM_VALIDATORS = {
    EMAIL: function(val) { return typeof val === 'string' && val.indexOf('@') !== -1; },
    DATE: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    DATETIME: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    NUMBER: function(val) { return typeof val === 'number' || (!isNaN(parseFloat(val)) && isFinite(val)); },
    TEXT: function(val) { return typeof val === 'string' && val.trim().length > 0; },
    EMAIL_LIST: function(val) {
        if (!val) return true;
        var emails = typeof val === 'string' ? val.split(',') : (Array.isArray(val) ? val : []);
        return emails.every(function(e) { 
            var trimmed = String(e).trim();
            return trimmed === "" || (trimmed.indexOf('@') !== -1 && trimmed.indexOf('.') !== -1);
        });
    },
    ACTION: function(val) { return typeof val === 'string' && val.trim().length > 0; }
};

/**
 * SchemaValidator — Validates row objects against the tool's COL_SCHEMA.
 */
var SchemaValidator = (function() {

    function validateRow(toolKey, rowObj) {
        var cfg = SyncEngine.getTool(toolKey);
        var schema = (cfg.FORMAT_CONFIG && cfg.FORMAT_CONFIG.COL_SCHEMA) ? cfg.FORMAT_CONFIG.COL_SCHEMA : [];
        var errors = [];

        schema.forEach(function(col) {
            var val = rowObj[col.header];
            var isMissing = (val === undefined || val === null || (typeof val === 'string' && val.trim() === ''));

            // Check Required
            if (col.required && isMissing) {
                errors.push("Column '" + col.header + "' is required.");
                return;
            }

            // Skip further checks if empty and not required
            if (isMissing) return;

            // Type Validation
            if (col.type && SYSTEM_VALIDATORS[col.type]) {
                if (!SYSTEM_VALIDATORS[col.type](val)) {
                    errors.push("Column '" + col.header + "' must be of type " + col.type + ".");
                }
            }

            // Regex Validation
            if (col.regex) {
                var re = new RegExp(col.regex);
                if (!re.test(String(val))) {
                    errors.push("Column '" + col.header + "' does not match required pattern.");
                }
            }

            // Numeric Range
            if (col.type === 'NUMBER') {
                var num = parseFloat(val);
                if (col.min !== undefined && num < col.min) errors.push("Column '" + col.header + "' must be at least " + col.min + ".");
                if (col.max !== undefined && num > col.max) errors.push("Column '" + col.header + "' must be at most " + col.max + ".");
            }
        });

        return {
            isValid: errors.length === 0,
            errors: errors
        };
    }

    return {
        validateRow: validateRow
    };
})();
// ==========================================
// _App_logClientError / _App_logClientInfo
// ==========================================
/**
 * Logs a client-side error from a sidebar (Unified for all tools).
 * @param {string|Object} err Error message or object
 * @param {string} [context] Optional context (e.g. 'Button Click')
 */
function _App_logClientError(err, context) {
    var source = 'Client UI';
    var ref = context || 'Default';
    var msg = (typeof err === 'object' && err.message) ? err.message : String(err);
    if (typeof err === 'object' && err.stack) msg += '\nStack: ' + err.stack;
    
    // Only attempt to log if Logger framework is loaded
    if (typeof Logger !== 'undefined' && typeof Logger.error === 'function') {
        Logger.setRunId(); // Ensure a Run ID exists for this independent execution
        Logger.error(source, ref, msg);
        if (typeof Logger.flushLogs === 'function') Logger.flushLogs();
    } else {
        console.error("[Client Error] " + ref + ": " + msg);
    }
}
/**
 * Logs a client-side info message from a sidebar (Unified for all tools).
 * @param {string} message 
 * @param {string} [context]
 */
function _App_logClientInfo(message, context) {
    if (typeof Logger !== 'undefined' && typeof Logger.info === 'function') {
        Logger.setRunId(); // Ensure a Run ID exists for this independent execution
        Logger.info('Client UI', context || 'Default', message);
        if (typeof Logger.flushLogs === 'function') Logger.flushLogs();
    } else {
        console.log("[Client Info] " + (context || "UI") + ": " + message);
    }
}

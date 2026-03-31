// ==========================================
// Centralized Data Validators
// ==========================================
var SYSTEM_VALIDATORS = {
    EMAIL: function(val) { return typeof val === 'string' && val.indexOf('@') !== -1; },
    DATE: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); },
    DATETIME: function(val) { return (val instanceof Date) || !isNaN(Date.parse(val)); }
};
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
        Logger.info('Client UI', context || 'Default', message);
        if (typeof Logger.flushLogs === 'function') Logger.flushLogs();
    } else {
        console.log("[Client Info] " + (context || "UI") + ": " + message);
    }
}

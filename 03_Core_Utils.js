/**
 * CORE UTILITIES
 * ==========================================
 * Common helper functions for error handling, locking, and API retries.
 */

Object.assign(App.Utils, (function() {
    return {
        ok: function(msg, data, meta) {
            return { success: true, message: msg || 'Success', data: data || null, meta: meta || null };
        },
        fail: function(msg, data, meta) {
            return { success: false, message: msg || 'Operation failed', data: data || null, meta: meta || null };
        },
        withLock: function(lockName, callback, timeoutMs) {
            var lock = LockService.getDocumentLock();
            if (!lock.tryLock(timeoutMs || 30000)) {
                throw new Error('System is busy' + (lockName ? ' (' + lockName + ')' : '') + '.');
            }
            try { return callback(); } finally { lock.releaseLock(); }
        },
        include: function(filename) {
            return HtmlService.createHtmlOutputFromFile(filename).getContent();
        },
        callWithBackoff: function(func, retries) {
            var max = (retries !== undefined) ? retries : 5;
            for (var n = 0; n <= max; n++) {
                try { return func(); } catch (e) {
                    var m = (e.message || '').toLowerCase();
                    var retriable = m.indexOf('403')!==-1||m.indexOf('429')!==-1||m.indexOf('500')!==-1||m.indexOf('rate limit')!==-1;
                    if (retriable && n < max) {
                        Utilities.sleep((Math.pow(2, n) * 1000) + Math.round(Math.random() * 1000));
                    } else throw e;
                }
            }
        }
    };
})());

// Backward Compatibility Aliases
function _App_ok(m, d, mt) { return App.Utils.ok(m, d, mt); }
function _App_fail(m, d, mt) { return App.Utils.fail(m, d, mt); }
function _App_withDocumentLock(l, c, t) { return App.Utils.withLock(l, c, t); }
function _App_include(f) { return App.Utils.include(f); }
function _App_callWithBackoff(f, r) { return App.Utils.callWithBackoff(f, r); }

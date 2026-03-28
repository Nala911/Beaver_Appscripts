function _App_ok(message, data, meta) {
    return {
        success: true,
        message: message || 'Success',
        data: data || null,
        meta: meta || null
    };
}

function _App_fail(message, data, meta) {
    return {
        success: false,
        message: message || 'Operation failed',
        data: data || null,
        meta: meta || null
    };
}
function _App_withDocumentLock(lockName, callback, timeoutMs) {
    var lock = LockService.getDocumentLock();
    var waitMs = timeoutMs || 30000;

    if (!lock.tryLock(waitMs)) {
        throw new Error('System is busy with another operation' + (lockName ? ' (' + lockName + ')' : '') + '. Please try again.');
    }

    try {
        return callback();
    } finally {
        lock.releaseLock();
    }
}
function _App_include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// ==========================================
// _App_throttle — Unified API Rate Limiter
// ==========================================
/**
 * Tracks cumulative API calls and sleeps (1 s) whenever a multiple of the
 * limit is crossed — preventing Google 429 "Too Many Requests" errors.
 */
function _App_throttle(tracker, callsMade, limit) {
    var _limit = limit || 10;
    var prev = tracker.apiCalls;
    tracker.apiCalls += callsMade;
    if (Math.floor(tracker.apiCalls / _limit) > Math.floor(prev / _limit)) {
        Utilities.sleep(1000);
    }
}
// ==========================================
// _App_callWithBackoff — Unified Exponential Backoff Retry
// ==========================================
/**
 * Runs a function; retries with exponential backoff on transient Google API errors.
 */
function _App_callWithBackoff(func, retries) {
    var maxRetries = (retries !== undefined) ? retries : 5;
    for (var n = 0; n <= maxRetries; n++) {
        try {
            return func();
        } catch (e) {
            var msg = (e.message || '').toLowerCase();
            var isRetriable = (
                msg.indexOf('403') !== -1 || msg.indexOf('429') !== -1 ||
                msg.indexOf('500') !== -1 || msg.indexOf('502') !== -1 ||
                msg.indexOf('503') !== -1 ||
                msg.indexOf('rate limit') !== -1 || msg.indexOf('quota') !== -1 ||
                msg.indexOf('limit exceeded') !== -1 || msg.indexOf('too many') !== -1
            );
            if (isRetriable && n < maxRetries) {
                var waitMs = (Math.pow(2, n) * 1000) + Math.round(Math.random() * 1000);
                console.warn('[_App_callWithBackoff] Retry ' + (n + 1) + '/' + maxRetries + ' in ' + waitMs + 'ms — ' + e.message);
                Utilities.sleep(waitMs);
            } else {
                throw e;
            }
        }
    }
}

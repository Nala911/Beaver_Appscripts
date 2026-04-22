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

// ==========================================
// Execution Time Tracking — 6-min Limit Guard
// ==========================================
var _App_executionStartTime = 0;
var _App_executionLimitMs = 330 * 1000; // 5.5 minutes (330 seconds)

/**
 * Resets the global execution timer to current time.
 */
function _App_resetExecutionTimer() {
    _App_executionStartTime = Date.now();
}

/**
 * Returns true if the script is approaching the Google Apps Script 6-minute limit.
 */
function _App_isExecutionLimitApproaching() {
    if (_App_executionStartTime === 0) return false;
    return (Date.now() - _App_executionStartTime > _App_executionLimitMs);
}

// ==========================================
// _App_BatchProcessor — Unified Iteration Engine
// ==========================================

/**
 * A centralized utility for processing batches of items (rows) with automated
 * progress tracking, error handling, time-limit guarding, and logging.
 *
 * @param {string} toolKey      - Tool identifier from SyncEngine (e.g. 'MAIL_MERGE')
 * @param {Array} items         - Array of items to process (usually objects with data and originalIndex)
 * @param {Function} processFn  - Callback function(item, index) that processes one item. 
 *                                 Should return a result object (e.g. row update data) or throw an error.
 * @param {Object} [options]    - Optional configuration:
 *                                 - {number} batchSize: How many items to process before updating progress/checking time (default 10)
 *                                 - {boolean} stopOnFailure: If true, stops the entire batch if one item fails (default false)
 *                                 - {Function} onBatchComplete: function(results) called after each batch segment.
 *
 * @returns {Object} { 
 *   processedCount: number, 
 *   errorCount: number, 
 *   timeLimitReached: boolean, 
 *   results: Array 
 * }
 */
function _App_BatchProcessor(toolKey, items, processFn, options) {
    var opts = options || {};
    var batchSize = opts.batchSize || 10;
    var total = items.length;
    
    var stats = {
        processedCount: 0,
        errorCount: 0,
        timeLimitReached: false,
        results: []
    };

    if (total === 0) return stats;

    // Ensure timer is running if not already set
    if (_App_executionStartTime === 0) _App_resetExecutionTimer();

    for (var i = 0; i < total; i += batchSize) {
        // 1. Time-Limit Guard
        if (_App_isExecutionLimitApproaching()) {
            stats.timeLimitReached = true;
            break;
        }

        var segment = items.slice(i, i + batchSize);
        var segmentResults = [];

        // 3. Process Segment
        for (var j = 0; j < segment.length; j++) {
            var item = segment[j];
            var globalIndex = i + j;

            try {
                // Wrap in backoff retry for transient API issues
                var result = _App_callWithBackoff(function() {
                    return processFn(item, globalIndex);
                });
                
                segmentResults.push(result);
                stats.results.push(result);
                stats.processedCount++;
            } catch (err) {
                stats.errorCount++;
                Logger.error(toolKey, "Item Index " + (item.originalIndex !== undefined ? item.originalIndex : globalIndex), err);
                
                if (opts.stopOnFailure) {
                    _App_clearProgress(toolKey);
                    throw err;
                }
                // Push null or error info to keep results array aligned if needed
                segmentResults.push({ error: err.message });
                stats.results.push({ error: err.message });
            }

            // Optional: Frequent progress updates for UI responsiveness
            // _App_setProgress(toolKey, stats.processedCount + stats.errorCount, total);
        }

        // 2. Progress Tracking (CacheService) — Update after segment
        _App_setProgress(toolKey, stats.processedCount + stats.errorCount, total);

        // 4. Batch Lifecycle Hook
        if (opts.onBatchComplete) {
            opts.onBatchComplete(segmentResults);
        }
    }

    // 5. Cleanup
    _App_clearProgress(toolKey);
    return stats;
}

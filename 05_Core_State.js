// ==========================================
// Progress Tracking — Unified CacheService Wrappers
// ==========================================

/**
 * Helper to build a unique cache key incorporating the Spreadsheet ID
 * to prevent progress cross-contamination between different open documents.
 * @param {string} toolName
 * @returns {string}
 * @private
 */
function _App_getProgressKey_(toolName) {
    var ssId = "";
    try {
        ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
    } catch (e) {
        // Fallback if executed outside of an active spreadsheet context
    }
    return ssId + "_" + toolName + CACHE_KEYS.PROGRESS;
}

/**
 * Stores batch operation progress for sidebar polling.
 * @param {string} toolName     - Tool key e.g. 'MAIL_SENDER'
 * @param {number} current      - Items processed so far
 * @param {number} total        - Total items queued
 * @param {number} [ttlSec=600] - Cache TTL in seconds (default 10 min)
 */
function _App_setProgress(toolName, current, total, ttlSec) {
    CacheService.getUserCache().put(
        _App_getProgressKey_(toolName),
        JSON.stringify({ current: current, total: total }),
        ttlSec || 600
    );
}

/**
 * Returns cached progress or null if expired/not set.
 * @param {string} toolName
 * @returns {{ current: number, total: number }|null}
 */
function _App_getProgress(toolName) {
    var data = CacheService.getUserCache().get(_App_getProgressKey_(toolName));
    return data ? JSON.parse(data) : null;
}

/**
 * Removes progress state after an operation completes.
 * @param {string} toolName
 */
function _App_clearProgress(toolName) {
    CacheService.getUserCache().remove(_App_getProgressKey_(toolName));
}

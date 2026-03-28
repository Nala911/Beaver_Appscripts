// ==========================================
// Progress Tracking — Unified CacheService Wrappers
// ==========================================

/**
 * Stores batch operation progress for sidebar polling.
 * @param {string} toolName     - Tool key e.g. 'MAIL_SENDER'
 * @param {number} current      - Items processed so far
 * @param {number} total        - Total items queued
 * @param {number} [ttlSec=600] - Cache TTL in seconds (default 10 min)
 */
function _App_setProgress(toolName, current, total, ttlSec) {
    CacheService.getUserCache().put(
        toolName + CACHE_KEYS.PROGRESS,
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
    var data = CacheService.getUserCache().get(toolName + CACHE_KEYS.PROGRESS);
    return data ? JSON.parse(data) : null;
}

/**
 * Removes progress state after an operation completes.
 * @param {string} toolName
 */
function _App_clearProgress(toolName) {
    CacheService.getUserCache().remove(toolName + CACHE_KEYS.PROGRESS);
}

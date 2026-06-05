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

function _App_formatStatus(prefixKey, message) {
    var theme = (typeof SHEET_THEME !== 'undefined' && SHEET_THEME && SHEET_THEME.STATUS_PREFIXES)
        ? SHEET_THEME
        : DEFAULT_SHEET_THEME;
    var prefixes = theme.STATUS_PREFIXES || {};
    var key = String(prefixKey || '').toUpperCase();
    var prefix = prefixes[key] || '';
    var text = (message === undefined || message === null) ? '' : String(message);
    return prefix + text;
}

function _App_makeRowPatch(rowNumber, updates) {
    var patch = {};
    if (updates && typeof updates === 'object') {
        Object.keys(updates).forEach(function (key) {
            patch[key] = updates[key];
        });
    }
    patch._rowNumber = rowNumber;
    return patch;
}

function _App_makeStatusPatch(rowNumber, prefixKey, message, updates) {
    var patch = _App_makeRowPatch(rowNumber, updates);
    patch.Status = _App_formatStatus(prefixKey, message);
    return patch;
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
// ==========================================
// Centralized Google API Error Translation Engine
// ==========================================
/**
 * Parses raw JSON or Google Apps Script exception messages and returns a user-friendly, actionable description.
 *
 * @param {Error|string|Object} err - The original exception thrown by an API or system call
 * @returns {Object} { message: string, isFatal: boolean, category: string }
 */
function _App_translateApiError(err) {
    var rawMsg = '';
    if (err) {
        if (typeof err === 'string') rawMsg = err;
        else if (err.message) rawMsg = err.message;
        else rawMsg = String(err);
    }
    
    var lower = rawMsg.toLowerCase();
    var result = {
        message: rawMsg,
        isFatal: false,
        category: 'unknown'
    };

    // 1. Authorization / OAuth Scope Errors
    if (lower.indexOf('authentication') !== -1 ||
        lower.indexOf('authorization') !== -1 ||
        lower.indexOf('scopes') !== -1 ||
        lower.indexOf('403') !== -1 ||
        lower.indexOf('access_denied') !== -1 ||
        lower.indexOf('unauthorized') !== -1 ||
        lower.indexOf('access denied') !== -1) {
        
        result.message = "⚠️ Authorization Required: The script lacks necessary Google permissions. Please re-authorize the script and try again.";
        result.isFatal = true;
        result.category = 'auth';
    }
    // 2. Rate Limits & Quotas
    else if (lower.indexOf('quota') !== -1 ||
             lower.indexOf('limit exceeded') !== -1 ||
             lower.indexOf('too many') !== -1 ||
             lower.indexOf('429') !== -1 ||
             lower.indexOf('rate limit') !== -1) {
             
        result.message = "⚠️ Quota Exceeded: You have hit Google's daily API limits for this service. Please wait a few hours or until tomorrow to resume syncing.";
        result.isFatal = true;
        result.category = 'quota';
    }
    // 3. Document or Resource Not Found
    else if (lower.indexOf('not found') !== -1 ||
             lower.indexOf('404') !== -1 ||
             lower.indexOf('cannot find') !== -1 ||
             lower.indexOf('inaccessible') !== -1) {
             
        result.message = "⚠️ Resource Not Found: The specified event, file, calendar, or folder ID is invalid, was deleted, or is not shared with this Google Account.";
        result.isFatal = false;
        result.category = 'not_found';
    }
    // 4. Temporary / Transient Server Failure
    else if (lower.indexOf('500') !== -1 ||
             lower.indexOf('502') !== -1 ||
             lower.indexOf('503') !== -1 ||
             lower.indexOf('internal error') !== -1 ||
             lower.indexOf('service unavailable') !== -1) {
             
        result.message = "❌ Service Error: Google servers are temporarily overwhelmed. The system will auto-retry, but please wait and try again if it continues to fail.";
        result.isFatal = false;
        result.category = 'transient';
    }
    // 5. Native parameter mismatch
    else if (lower.indexOf("don't match the method signature") !== -1 ||
             lower.indexOf("invalid parameter") !== -1 ||
             lower.indexOf("400") !== -1) {
             
        result.message = "⚠️ Formatting Error: The data provided in the sheet doesn't match the required types (e.g. invalid date formats, missing compulsory fields).";
        result.isFatal = false;
        result.category = 'validation';
    }
    
    return result;
}

// ==========================================
// Centralized Sheet Row Validator
// ==========================================
/**
 * Validates a sheet row object against the tool's registered COL_SCHEMA rules.
 * @param {Object} item - Row object (usually from SheetManager.readPendingObjects)
 * @param {string} toolKey - Tool identifier (e.g. 'CALENDAR_SYNC')
 * @returns {string|null} Error string if validation fails, or null if row is fully valid.
 */
function _App_validateRowAgainstSchema(item, toolKey) {
    if (typeof SyncEngine === 'undefined' || !toolKey) return null;
    try {
        var toolCfg = SyncEngine.getTool(toolKey);
        if (!toolCfg || !toolCfg.FORMAT_CONFIG || !toolCfg.FORMAT_CONFIG.COL_SCHEMA) {
            return null;
        }
        var schema = toolCfg.FORMAT_CONFIG.COL_SCHEMA;
        for (var i = 0; i < schema.length; i++) {
            var col = schema[i];
            var colHeader = col.header;
            var colType = col.type;
            
            // Skip validation on standard structural columns
            if (colHeader === 'Action' || colHeader === 'Status') continue;
            
            if (item && Object.prototype.hasOwnProperty.call(item, colHeader)) {
                var val = item[colHeader];
                var isValid = _App_validateValueByType(colType, val, col);
                if (!isValid) {
                    return "⚠️ Data Error: Invalid format in column '" + colHeader + "' (Expected " + colType + ")";
                }
            }
        }
    } catch (e) {
        // Fallback gracefully
    }
    return null;
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
                // 3a. Pre-Validation Engine Check
                var schemaValidationError = _App_validateRowAgainstSchema(item, toolKey);
                if (schemaValidationError) {
                    throw new Error(schemaValidationError);
                }

                // Wrap in backoff retry for transient API issues
                var result = _App_callWithBackoff(function() {
                    return processFn(item, globalIndex);
                });
                
                segmentResults.push(result);
                stats.results.push(result);
                stats.processedCount++;
            } catch (err) {
                stats.errorCount++;
                
                var translated = _App_translateApiError(err);
                
                // Return an error object to the tool so it can write to the Status column
                var errObj = { isError: true, error: translated.message };
                if (item && item._rowNumber !== undefined) {
                    errObj._rowNumber = item._rowNumber;
                }
                
                segmentResults.push(errObj);
                stats.results.push(errObj);
                
                // If it is a fatal system/connection error (OAuth scopes / Rate limit), halt the execution
                if (translated.isFatal) {
                    _App_clearProgress(toolKey);
                    
                    // Flush what was successfully processed (and the current fatal item) before halting
                    if (opts.onBatchComplete && segmentResults.length > 0) {
                        try {
                            opts.onBatchComplete(segmentResults);
                        } catch (writeErr) {
                            // Suppress errors during flush write to prioritize bubbling up original fatal error
                        }
                    }
                    
                    throw new Error(translated.message);
                }
                
                if (opts.stopOnFailure) {
                    _App_clearProgress(toolKey);
                    throw err;
                }
            }

            // Optional: Frequent progress updates for UI responsiveness
            // _App_setProgress(toolKey, stats.processedCount + stats.errorCount, total);
        }

        // 2. Progress Tracking (CacheService) — Update after segment
        _App_setProgress(toolKey, stats.processedCount + stats.errorCount, total);

        // 4. Batch Lifecycle Hook
        if (opts.onBatchComplete && segmentResults.length > 0) {
            opts.onBatchComplete(segmentResults);
        }
    }

    // 5. Cleanup
    _App_clearProgress(toolKey);
    return stats;
}

/**
 * Centralized utility to batch patch rows in SheetManager from BatchProcessor results.
 *
 * @param {string} toolKey - Tool identifier (e.g., 'MAIL_MERGE')
 * @param {Array} batchResults - The results array from the batch segment
 * @param {Function} [successFieldsMapper] - Optional callback function(res) returning an object
 *                                            containing the middle/read-only columns to update on success.
 *                                            (Action and Status are automatically handled if not returned).
 */
function _App_batchPatchResults(toolKey, batchResults, successFieldsMapper) {
    var rowNumbers = [];
    var patchData = [];

    batchResults.forEach(function (res) {
        if (res && res._rowNumber !== undefined) {
            rowNumbers.push(res._rowNumber);
            if (res.isError) {
                // On error, write the error status. We keep the action so user can retry.
                patchData.push(_App_makeStatusPatch(res._rowNumber, 'ERROR', res.error));
            } else {
                // On success, construct updates
                var updates = {
                    'Action': res.action !== undefined ? res.action : '',
                    'Status': res.status !== undefined ? res.status : _App_formatStatus('SUCCESS', 'Processed')
                };
                if (successFieldsMapper) {
                    var customFields = successFieldsMapper(res);
                    if (customFields && typeof customFields === 'object') {
                        Object.keys(customFields).forEach(function (k) {
                            updates[k] = customFields[k];
                        });
                    }
                }
                patchData.push(_App_makeRowPatch(res._rowNumber, updates));
            }
        }
    });

    if (rowNumbers.length > 0) {
        SheetManager.batchPatchRows(toolKey, rowNumbers, patchData);
    }
}

/**
 * Converts a 1-based column number into its corresponding A-Z/AA-ZZ Excel-style column letter.
 * @param {number} col - 1-based column index
 * @returns {string} The column letter (e.g. 'A', 'Z', 'AA')
 */
function _App_getColumnLetter(col) {
    var temp, letter = '';
    while (col > 0) {
        temp = (col - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        col = (col - temp - 1) / 26;
    }
    return letter;
}

// ==========================================
// Centralized Document & Email Utilities
// ==========================================

/**
 * Escapes special regex characters in a string.
 */
function _App_escapeRegExp(string) {
  if (!string) return "";
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Retrieves a Drive file blob by URL or raw File ID.
 */
function _App_getDriveAttachment(fileIdOrUrl) {
  try {
    if (!fileIdOrUrl) return null;
    var fileId = fileIdOrUrl;
    var match = fileIdOrUrl.match(/[-\w]{25,}/);
    if (match) fileId = match[0];

    var file = DriveApp.getFileById(fileId);
    return file.getBlob();
  } catch (e) {
    throw new Error("Cannot find attachment in Drive (" + fileIdOrUrl + ")");
  }
}

/**
 * Merges two comma-separated lists of email addresses uniquely.
 */
function _App_mergeEmails(existingStr, newStr) {
  if (!newStr) return existingStr || "";
  var existingArr = (existingStr || "").split(',').map(function (e) { return e.trim(); }).filter(function (e) { return e; });
  var newArr = (newStr || "").split(',').map(function (e) { return e.trim(); }).filter(function (e) { return e; });
  newArr.forEach(function (em) {
    if (existingArr.indexOf(em) === -1) {
      existingArr.push(em);
    }
  });
  return existingArr.join(',');
}

/**
 * Formats a Date object or parseable date-time representation safely using the 
 * spreadsheet's active timezone. Falls back to Session.getScriptTimeZone() if needed.
 * @param {Date|string|number} date - Date representation to format
 * @param {string} [format] - Target format string, defaults to "MM/dd/yyyy HH:mm:ss"
 * @returns {string} Formatted date-time string
 */
function _App_formatDateTime(date, format) {
    if (!date) return "";
    var targetDate = (date instanceof Date) ? date : new Date(date);
    if (isNaN(targetDate.getTime())) return "";
    var tz = Session.getScriptTimeZone();
    try {
        tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    } catch (e) {
        // Fallback safely if run outside spreadsheet context
    }
    return Utilities.formatDate(targetDate, tz, format || "MM/dd/yyyy HH:mm:ss");
}



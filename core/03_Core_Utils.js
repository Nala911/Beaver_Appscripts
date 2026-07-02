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

function _App_include(filename) {
    try {
        return HtmlService.createHtmlOutputFromFile(filename).getContent();
    } catch (e) {
        if (filename.indexOf('/') === -1) {
            try {
                return HtmlService.createHtmlOutputFromFile('core/' + filename).getContent();
            } catch (e2) {
                try {
                    return HtmlService.createHtmlOutputFromFile('tools/' + filename).getContent();
                } catch (e3) {
                    throw new Error("Could not find file: '" + filename + "' (tried raw, core/, tools/)");
                }
            }
        }
        throw e;
    }
}

function _App_createTemplateFromFile(filename) {
    try {
        return HtmlService.createTemplateFromFile(filename);
    } catch (e) {
        if (filename.indexOf('/') === -1) {
            try {
                return HtmlService.createTemplateFromFile('tools/' + filename);
            } catch (e2) {
                try {
                    return HtmlService.createTemplateFromFile('core/' + filename);
                } catch (e3) {
                    throw new Error("Could not find template file: '" + filename + "' (tried raw, tools/, core/)");
                }
            }
        }
        throw e;
    }
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
            var is403Retriable = msg.indexOf('403') !== -1 && (
                msg.indexOf('rate limit') !== -1 ||
                msg.indexOf('quota') !== -1 ||
                msg.indexOf('limit exceeded') !== -1 ||
                msg.indexOf('too many') !== -1
            );
            var isRetriable = (
                is403Retriable || msg.indexOf('429') !== -1 ||
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
 * Consolidated utility to extract a Google Resource ID (Docs, Sheets, Slides, Forms, Drive files/folders)
 * from a full URL, or returns the input if it is already a raw ID.
 * @param {string} urlOrId - The URL or ID string.
 * @returns {string|null} The extracted ID or null if invalid/empty.
 */
function _App_extractIdFromUrl(urlOrId) {
    if (!urlOrId || typeof urlOrId !== 'string') return null;
    var trimmed = urlOrId.trim();
    if (!trimmed) return null;
    
    // Pattern matches typical Google ID formats (usually 25-100 characters of alphanumeric/hyphen/underscore)
    var idMatch = trimmed.match(/[-\w]{25,100}/);
    if (idMatch) {
        return idMatch[0];
    }
    
    // Fallback pattern for URLs containing /d/ID
    var dMatch = trimmed.match(/\/d\/([-\w]+)/);
    if (dMatch) {
        return dMatch[1];
    }
    
    // If it looks like a raw ID (no slashes), return it
    if (trimmed.indexOf('/') === -1) {
        return trimmed;
    }
    
    return null;
}

/**
 * Utility to execute multiple HTTP requests in parallel using UrlFetchApp.fetchAll.
 * Can be used by tools for parallel API queries or external webhooks.
 * @param {Object[]} requests - Array of request objects {url, method, headers, payload, etc.}
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse[]} Array of HTTPResponse objects.
 */
function _App_fetchParallel(requests) {
    if (!requests || requests.length === 0) return [];
    return UrlFetchApp.fetchAll(requests);
}

/**
 * Retrieves a Drive file blob by URL or raw File ID.
 */
function _App_getDriveAttachment(fileIdOrUrl) {
  try {
    if (!fileIdOrUrl) return null;
    var fileId = _App_extractIdFromUrl(fileIdOrUrl);
    if (!fileId) throw new Error("Invalid File ID");

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






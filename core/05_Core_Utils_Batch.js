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
                
                var statusMsg = translated.message;
                if (translated.category !== 'unknown' && translated.originalMessage && translated.originalMessage !== translated.message) {
                    var detailText = translated.originalMessage;
                    if (detailText.length > 150) {
                        detailText = detailText.substring(0, 147) + "...";
                    }
                    // Strip duplicate emojis and error labels from original message
                    detailText = detailText.replace(/^[⚠️❌⏳ℹ️✅]\s*/, '').replace(/^(?:Data|Formatting|API|System)\s+Error:\s*/i, '');
                    
                    var colonIdx = translated.message.indexOf(':');
                    var prefix = colonIdx !== -1 ? translated.message.substring(0, colonIdx + 1) : "❌ Error:";
                    statusMsg = prefix + " " + detailText;
                }
                
                // Return an error object to the tool so it can write to the Status column
                var errObj = { isError: true, error: statusMsg };
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

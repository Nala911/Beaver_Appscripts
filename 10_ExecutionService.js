/**
 * ExecutionService — Orchestrates row-level sync operations with built-in
 * throttling, exponential backoff, and schema validation.
 */
var ExecutionService = (function() {

    /**
     * Processes rows with pending actions for a specific tool.
     * @param {string} toolKey - The tool registry key.
     * @param {function} rowProcessor - Callback(rowObj, toolKey) that performs the sync action.
     * @param {Object} [options] - Optional settings (e.g., limit, skipValidation).
     */
    function processPendingRows(toolKey, rowProcessor, options) {
        var opts = options || {};
        var rows = SheetManager.readPendingActions(toolKey);
        if (rows.length === 0) return { processed: 0, skipped: 0, errors: 0 };

        if (opts.limit && opts.limit > 0) {
            rows = rows.slice(0, opts.limit);
        }

        var stats = { processed: 0, skipped: 0, errors: 0 };
        var tracker = { apiCalls: 0 };
        var throttleLimit = opts.throttleLimit || 10;
        var total = rows.length;

        // Start Progress Tracking
        _App_setProgress(toolKey, 0, total);

        rows.forEach(function(row, index) {
            try {
                // 1. Schema Validation (unless skipped)
                if (!opts.skipValidation) {
                    var valResult = SchemaValidator.validateRow(toolKey, row);
                    if (!valResult.isValid) {
                        SheetManager.patchRow(toolKey, row._rowNumber, {
                            'Action': 'ERROR',
                            'Log': valResult.errors.join(' ')
                        });
                        stats.errors++;
                        return;
                    }
                }

                // 2. Data Mapping (Casting to types)
                var typedRow = DataMapper.castRow(toolKey, row);

                // 3. Execute with Backoff & Throttling
                _App_callWithBackoff(function() {
                    rowProcessor(typedRow, toolKey);
                });

                stats.processed++;
                
                // 4. Update Progress
                _App_setProgress(toolKey, index + 1, total);

                // 5. API Throttling
                _App_throttle(tracker, 1, throttleLimit);

            } catch (e) {
                console.error("[ExecutionService] Error processing row " + row._rowNumber + ": " + e.message);
                SheetManager.patchRow(toolKey, row._rowNumber, {
                    'Action': 'ERROR',
                    'Log': e.message
                });
                stats.errors++;
            }
        });

        // Clear Progress Tracking
        _App_clearProgress(toolKey);

        return stats;
    }

    return {
        processPendingRows: processPendingRows
    };

})();

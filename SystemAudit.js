// ==========================================
// System Audit & Testing Logic
// ==========================================

/**
 * Extensible rules engine for auditing tools and environment.
 * Each rule receives (cfg, sheet, report) and mutates the report.
 */
var AuditRules = [
    {
        name: 'Registry Metadata Integrity',
        run: function (cfg, sheet, report) {
            var issues = SyncEngine.auditTool(cfg.TOOL_KEY || report.key);
            if (issues.length > 0) {
                report.status = 'ERROR';
                report.issues.push("Registry issues: " + issues.join(' | '));
            }

            if (cfg.LAUNCH_MODE === TOOL_LAUNCH_MODES.SIDEBAR && cfg.SIDEBAR_HTML) {
                try {
                    HtmlService.createTemplateFromFile(cfg.SIDEBAR_HTML);
                } catch (e) {
                    report.status = 'ERROR';
                    report.issues.push("Sidebar HTML '" + cfg.SIDEBAR_HTML + "' is missing or invalid.");
                }
            }

            if (cfg.LAUNCH_MODE === TOOL_LAUNCH_MODES.MODAL && (cfg.MODAL_HTML || cfg.SIDEBAR_HTML)) {
                try {
                    HtmlService.createTemplateFromFile(cfg.MODAL_HTML || cfg.SIDEBAR_HTML);
                } catch (e) {
                    report.status = 'ERROR';
                    report.issues.push("Modal HTML '" + (cfg.MODAL_HTML || cfg.SIDEBAR_HTML) + "' is missing or invalid.");
                }
            }
        }
    },
    {
        name: 'Check Sheet Existence',
        run: function (cfg, sheet, report) {
            if (!sheet) {
                report.status = 'WARN';
                report.issues.push("Sheet '" + cfg.SHEET_NAME + "' is missing (Not initialized).");
            }
        }
    },
    {
        name: 'Check Header Integrity',
        run: function (cfg, sheet, report) {
            if (!sheet) return; // Skip if no sheet
            if (cfg.HEADERS && Array.isArray(cfg.HEADERS)) {
                var lastCol = sheet.getLastColumn();
                if (lastCol === 0) {
                    report.status = 'WARN';
                    report.issues.push("Sheet is completely empty. Headers are missing.");
                } else if (lastCol < cfg.HEADERS.length) {
                    report.status = 'WARN';
                    report.issues.push("Sheet has fewer columns than expected (Expected " + cfg.HEADERS.length + ", found " + lastCol + ").");
                } else {
                    var actualHeaders = sheet.getRange(1, 1, 1, cfg.HEADERS.length).getValues()[0];
                    cfg.HEADERS.forEach(function (h, i) {
                        if (actualHeaders[i] !== h) {
                            report.status = 'WARN';
                            report.issues.push("Header mismatch at Col " + (i + 1) + ": Expected '" + h + "', found '" + actualHeaders[i] + "'.");
                        }
                    });
                }
            }
        }
    },
    {
        name: 'Dependency Checks',
        run: function (cfg, sheet, report) {
            if (cfg.REQUIRED_SERVICES && Array.isArray(cfg.REQUIRED_SERVICES)) {
                cfg.REQUIRED_SERVICES.forEach(function (svc) {
                    try {
                        if (!svc.test()) {
                            report.status = 'ERROR';
                            report.issues.push("Advanced Service '" + svc.name + "' is not enabled.");
                        }
                    } catch (err) {
                        report.status = 'WARN';
                        report.issues.push(svc.name + " test failed: " + err.message);
                    }
                });
            }
        }
    },
    {
        name: 'Deep Data Integrity',
        run: function (cfg, sheet, report) {
            if (!sheet || !cfg.FORMAT_CONFIG || !cfg.FORMAT_CONFIG.COL_SCHEMA) return;
            
            var schema = cfg.FORMAT_CONFIG.COL_SCHEMA;
            var numRows = sheet.getLastRow() - 1; // excluding header
            if (numRows < 1) return; // No data to check

            // Sample up to 50 rows
            var sampleRows = Math.min(numRows, 50);
            var data = sheet.getRange(2, 1, sampleRows, Math.max(sheet.getLastColumn(), 1)).getValues();
            
            var warnings = [];
            data.forEach(function(row, rowIndex) {
                schema.forEach(function(colDef, colIndex) {
                    if (colIndex >= row.length) return;
                    var cellValue = row[colIndex];
                    if (cellValue === '' || cellValue === null || cellValue === undefined) return; 
                    
                    // Unified Validation via SYSTEM_VALIDATORS
                    if (typeof SYSTEM_VALIDATORS !== 'undefined' && SYSTEM_VALIDATORS[colDef.type]) {
                        try {
                            if (!SYSTEM_VALIDATORS[colDef.type](cellValue)) {
                                warnings.push("Row " + (rowIndex + 2) + ", Col " + (colIndex + 1) + " ('" + colDef.header + "'): Invalid " + colDef.type + " format ('" + String(cellValue) + "').");
                            }
                        } catch (err) {
                            warnings.push("Validator crashed for " + colDef.type + " at Row " + (rowIndex + 2) + ": " + err.message);
                        }
                    }
                });
            });

            if (warnings.length > 0) {
                report.status = 'WARN';
                // Only show first 3 warnings to avoid extreme log bloat
                var displayWarnings = warnings.slice(0, 3).join(' | ');
                if (warnings.length > 3) displayWarnings += " | ...and " + (warnings.length - 3) + " more data issues.";
                report.issues.push("Data Integrity Issues: " + displayWarnings);
            }
        }
    },
    {
        name: 'External API & Auth Ping',
        run: function (cfg, sheet, report) {
            // If the tool defines an API_PING property/function, run it to verify Auth state
            if (typeof cfg.API_PING === 'function') {
                try {
                    var pingResult = cfg.API_PING();
                    if (pingResult !== true) {
                        report.status = 'WARN';
                        report.issues.push("API Ping returned irregular soft-fail: " + String(pingResult));
                    }
                } catch (e) {
                    report.status = 'ERROR';
                    report.issues.push("API Ping failed (Authentication or Service unreachable): " + e.message);
                }
            }
        }
    },
    {
        name: 'Tool-Specific Custom Audit',
        run: function (cfg, sheet, report) {
            // Allows tools to define their own custom validation logic
            if (typeof cfg.CUSTOM_AUDIT === 'function') {
                try {
                    cfg.CUSTOM_AUDIT(sheet, report);
                } catch (e) {
                    report.status = 'ERROR';
                    report.issues.push("CUSTOM_AUDIT crashed: " + e.message);
                }
            }
        }
    },
    {
        name: 'Performance Benchmarking Metrics',
        run: function (cfg, sheet, report) {
            if (!sheet) return;
            var maxRows = sheet.getMaxRows();
            var maxCols = sheet.getMaxColumns();
            var totalCells = maxRows * maxCols;
            
            // Warn if approaching 2M cells since Apps Script begins to slow down, limit is 10M cells.
            if (totalCells > 2000000) {
                 report.status = 'WARN';
                 report.issues.push("Sheet has " + totalCells + " cells. Approaching high limits, consider archiving old data.");
            }

            // Benchmark data retrieval for up to 100 rows
            var rBound = sheet.getLastRow();
            if (rBound > 1) {
                var start = Date.now();
                sheet.getRange(1, 1, Math.min(rBound, 100), Math.max(sheet.getLastColumn(), 1)).getValues();
                var end = Date.now();
                var duration = end - start;
                
                if (duration > 2000) { // arbitrary 2 second limit for a small sample
                    report.status = 'WARN';
                    report.issues.push("Data retrieval is slow (" + duration + "ms to read sample rows). Extrapolated performance may degrade user experience.");
                }
            }
        }
    }
];

// Single-run Global Level Rules 
var GlobalAuditRules = [
    {
        name: 'Environment & Configuration',
        run: function (summary, results) {
            try {
                var systemEnabled = _App_getProperty(APP_PROPS.SYSTEM_ENABLED);
                if (!systemEnabled || systemEnabled !== 'true') {
                    _addGlobalResult(summary, results, 'Environment Config', 'INFO', 'SYSTEM_ENABLED script property is missing or false. Pipeline logic is currently suspended.');
                }
                
                var debugEnabled = _App_getProperty(APP_PROPS.ENABLE_DEBUG_LOGGING);
                if (!debugEnabled || debugEnabled !== 'true') {
                    _addGlobalResult(summary, results, 'Environment Config', 'INFO', 'ENABLE_DEBUG_LOGGING is unset or false. Some developer logs will not be recorded.');
                }
                
            } catch (e) {
                _addGlobalResult(summary, results, 'Environment Config', 'ERROR', 'Error accessing properties service: ' + e.message, e);
            }
        }
    },
    {
        name: 'Quota & Limits',
        run: function (summary, results) {
            try {
                var remainingEmails = MailApp.getRemainingDailyQuota();
                if (remainingEmails < 100) {
                    _addGlobalResult(summary, results, 'Quota Check', 'WARN', 'Only ' + remainingEmails + ' emails remaining in daily quota. Mail operations may fail soon.');
                } else {
                    _addGlobalResult(summary, results, 'Quota Check', 'SUCCESS', 'Email quota is healthy (' + remainingEmails + ' remaining).');
                }
            } catch (e) {
                // Ignore API execution errors here gracefully if MailApp is restricted 
            }
        }
    },
    {
        name: 'Concurrency & Locking Health',
        run: function (summary, results) {
            try {
                var lock = LockService.getDocumentLock();
                if (lock.tryLock(500)) {
                    lock.releaseLock();
                    _addGlobalResult(summary, results, 'Locking Service', 'SUCCESS', 'Locking service is operational.');
                } else {
                    _addGlobalResult(summary, results, 'Locking Service', 'WARN', 'Could not acquire lock quickly. Document may be heavily contested.');
                }
            } catch (e) {
                _addGlobalResult(summary, results, 'Locking Service', 'ERROR', 'Locking service failure: ' + e.message, e);
            }
        }
    },
    {
        name: 'Trigger Health Assessment',
        run: function (summary, results) {
            try {
                var triggers = ScriptApp.getProjectTriggers();
                if (triggers.length === 0) {
                    _addGlobalResult(summary, results, 'Trigger Health', 'INFO', 'No triggers found in the project.');
                } else {
                    var triggerDetails = triggers.map(function(t) { return t.getHandlerFunction(); });
                    _addGlobalResult(summary, results, 'Trigger Health', 'SUCCESS', triggers.length + ' active triggers found.');
                    
                    if (triggers.length > 15) {
                        _addGlobalResult(summary, results, 'Trigger Health', 'WARN', 'Approaching trigger limit. You have ' + triggers.length + ' active triggers.');
                    }
                }
            } catch (e) {
                _addGlobalResult(summary, results, 'Trigger Health', 'ERROR', 'Error fetching triggers: ' + e.message, e);
            }
        }
    },
    {
        name: 'Theme & Storage Health',
        run: function (summary, results) {
            try {
                var themeStr = _App_getRawProperty(APP_PROPS.THEME);
                if (themeStr) {
                    deepMergeTheme_(DEFAULT_SHEET_THEME, JSON.parse(themeStr));
                }
            } catch (e) {
                _addGlobalResult(summary, results, 'Storage Health', 'ERROR', 'WorkspaceSync_SHEET_THEME contains malformed JSON. UI layers may crash: ' + e.message, e);
            }
        }
    }
];

function _addGlobalResult(summary, results, title, status, msg, errorObj) {
    results.push({
        key: 'GLOBAL', 
        title: title, 
        status: status, 
        issues: [msg]
    });
    if (status === 'ERROR') {
        summary.errors++;
        Logger.error(title, 'Audit Global Check', errorObj || msg);
    } else if (status === 'WARN') {
        summary.warnings++;
        Logger.warn(title, 'Audit Global Check', msg);
    } else {
        summary.passed++;
        Logger.info(title, 'Audit Global Check', msg);
    }
}

/**
 * Runs a comprehensive audit of all registered tools and the global environment.
 * Executes global rules followed by tool-specific rules.
 */
function Logger_runSystemAudit() {
    return Logger.run('LOGS', 'System Audit', function () {
        var tools = SyncEngine.getAllTools();
        var keys = SyncEngine.getToolKeys();
        var results = [];
        var summary = { passed: 0, warnings: 0, errors: 0 };

        Logger.info('System Audit', 'Global', "Starting rule-based audit for " + keys.length + " tools...");

        // 1. Run Global Rules
        GlobalAuditRules.forEach(function (rule) {
             rule.run(summary, results);
        });

        // 2. Run Tool-Specific Rules
        var ss = SpreadsheetApp.getActiveSpreadsheet();

        keys.forEach(function (key) {
            var cfg = tools[key];

            var report = { key: key, title: cfg.TITLE || key, status: 'SUCCESS', issues: [], errorObj: null };            
            try {
                var sheet = ss.getSheetByName(cfg.SHEET_NAME);
                AuditRules.forEach(function (rule) {
                    try {
                        rule.run(cfg, sheet, report);
                    } catch (e) {
                        report.status = 'ERROR';
                        report.errorObj = e;
                        report.issues.push("Rule '" + rule.name + "' crashed: " + e.message);
                    }
                });
            } catch(e) {
                report.status = 'ERROR';
                report.errorObj = e;
                report.issues.push("Audit Crash for " + key + ": " + e.message);
            }

            if (report.status === 'ERROR') {
                summary.errors++;
                Logger.error(report.title, 'Audit Test', report.errorObj || report.issues.join(' | '), { issues: report.issues, module: key });
            } else if (report.status === 'WARN') {
                summary.warnings++;
                Logger.warn(report.title, 'Audit Test', report.issues.join(' | '), { issues: report.issues, module: key });
            } else {
                summary.passed++;
                Logger.success(report.title, 'Audit Test', "All checks passed.");
            }

            results.push(report);
        });

        Logger.flushLogs();
        return {
            summary: summary,
            details: results
        };
    }, true);
}


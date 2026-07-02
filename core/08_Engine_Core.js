// ==========================================
// SyncEngine — Plugin Registration System
// ==========================================

var SyncEngine = (function() {
    var registry = {};

    function _validateToolConfig(key, config) {
        var issues = [];

        if (!config.SHEET_NAME) issues.push("Missing SHEET_NAME.");
        if (!config.TITLE) issues.push("Missing TITLE.");

        if (config.MENU_LABEL && !config.MENU_ENTRYPOINT) {
            issues.push("MENU_LABEL requires MENU_ENTRYPOINT.");
        }

        if (config.LAUNCH_MODE === TOOL_LAUNCH_MODES.SIDEBAR && !config.SIDEBAR_HTML && config.MENU_ENTRYPOINT) {
            issues.push("Sidebar tools require SIDEBAR_HTML.");
        }

        if (config.LAUNCH_MODE === TOOL_LAUNCH_MODES.MODAL && !(config.MODAL_HTML || config.SIDEBAR_HTML) && config.MENU_ENTRYPOINT) {
            issues.push("Modal tools require MODAL_HTML or SIDEBAR_HTML.");
        }

        if (config.FORMAT_CONFIG && config.FORMAT_CONFIG.COL_SCHEMA && !Array.isArray(config.FORMAT_CONFIG.COL_SCHEMA)) {
            issues.push("FORMAT_CONFIG.COL_SCHEMA must be an array.");
        } else if (config.FORMAT_CONFIG && Array.isArray(config.FORMAT_CONFIG.COL_SCHEMA)) {
            var schema = config.FORMAT_CONFIG.COL_SCHEMA;
            if (schema.length > 0 && schema[0].type !== 'ACTION') {
                issues.push("First column must be of type 'ACTION'.");
            }
            if (schema.length > 1 && schema[1].type !== 'STATUS') {
                issues.push("Second column must be of type 'STATUS'.");
            }
        }

        if (config.HELP_ITEMS) {
            if (typeof config.HELP_ITEMS !== 'object') {
                issues.push("HELP_ITEMS must be an object.");
            } else if (config.HELP_ITEMS.items && !Array.isArray(config.HELP_ITEMS.items)) {
                issues.push("HELP_ITEMS.items must be an array.");
            }
        }

        return issues;
    }

    /**
     * Registers a tool with the engine.
     * Automatically processes COL_SCHEMA to generate HEADERS and totalCols.
     */
    function registerTool(key, config) {
        config.TOOL_KEY = key;
        config.MENU_LABEL = config.MENU_LABEL || config.TITLE;
        config.MENU_ORDER = typeof config.MENU_ORDER === 'number' ? config.MENU_ORDER : 999;
        config.LAUNCH_MODE = config.LAUNCH_MODE || TOOL_LAUNCH_MODES.SIDEBAR;

        // Unifying defaults
        config.FROZEN_ROWS = typeof config.FROZEN_ROWS === 'number' ? config.FROZEN_ROWS : 1;
        config.FROZEN_COLS = typeof config.FROZEN_COLS === 'number' ? config.FROZEN_COLS : 2;

        // Post-process the config (generate HEADERS, totalCols, COL_WIDTHS from SCHEMA)
        if (config.FORMAT_CONFIG && config.FORMAT_CONFIG.COL_SCHEMA) {
            config.HEADERS = config.FORMAT_CONFIG.COL_SCHEMA.map(function(c) { return c.header; });
            config.FORMAT_CONFIG.totalCols = config.FORMAT_CONFIG.COL_SCHEMA.length;
            
            if (!config.COL_WIDTHS) {
                config.COL_WIDTHS = config.FORMAT_CONFIG.COL_SCHEMA.map(function(col) {
                    return col.width || DEFAULT_COL_WIDTHS[col.type] || 150;
                });
            }

            // Auto-inject default conditional formatting rules
            if (!config.FORMAT_CONFIG.conditionalRules) {
                config.FORMAT_CONFIG.conditionalRules = [];
            }

            var actionColIndex = -1;
            var statusColIndex = -1;
            config.FORMAT_CONFIG.COL_SCHEMA.forEach(function(col, idx) {
                if (col.type === 'ACTION') actionColIndex = idx + 1;
                if (col.type === 'STATUS') statusColIndex = idx + 1;
            });

            var hasPendingRule = config.FORMAT_CONFIG.conditionalRules.some(function(r) { return r.type === 'pending'; });
            var hasSuccessRule = config.FORMAT_CONFIG.conditionalRules.some(function(r) { return r.type === 'success'; });
            var hasWarningRule = config.FORMAT_CONFIG.conditionalRules.some(function(r) { return r.type === 'error'; });
            var hasErrorRule = config.FORMAT_CONFIG.conditionalRules.some(function(r) { return r.type === 'errorCross'; });

            if (actionColIndex !== -1 && !hasPendingRule) {
                var actionColLetter = _App_getColumnLetter(actionColIndex);
                config.FORMAT_CONFIG.conditionalRules.push({
                    type: 'pending',
                    actionCol: actionColLetter,
                    scope: 'actionOnly'
                });
            }

            if (statusColIndex !== -1) {
                var statusColLetter = _App_getColumnLetter(statusColIndex);
                if (!hasSuccessRule) {
                    config.FORMAT_CONFIG.conditionalRules.push({
                        type: 'success',
                        statusCol: statusColLetter,
                        scope: 'statusOnly'
                    });
                }
                if (!hasWarningRule) {
                    config.FORMAT_CONFIG.conditionalRules.push({
                        type: 'error',
                        statusCol: statusColLetter,
                        scope: 'statusOnly'
                    });
                }
                if (!hasErrorRule) {
                    config.FORMAT_CONFIG.conditionalRules.push({
                        type: 'errorCross',
                        statusCol: statusColLetter,
                        scope: 'statusOnly'
                    });
                }
            }
        }

        var issues = _validateToolConfig(key, config);
        if (issues.length > 0) {
            throw new Error("Tool '" + key + "' is misconfigured: " + issues.join(' '));
        }

        registry[key] = config;
    }

    /**
     * Retrieves a tool configuration by key.
     */
    function getTool(key) {
        var cfg = registry[key];
        if (!cfg) throw new Error('Unknown tool key: "' + key + '". Ensure the tool is registered via SyncEngine.registerTool().');
        return cfg;
    }

    /**
     * Returns all registered tools.
     */
    function getAllTools() {
        return registry;
    }

    function getToolKeys() {
        return Object.keys(registry);
    }

    function runAction(toolKey, actionName, args) {
        var cfg = getTool(toolKey);
        var action = cfg.ACTIONS && cfg.ACTIONS[actionName];
        if (typeof action !== 'function') {
            throw new Error("Action '" + actionName + "' not found on tool '" + toolKey + "'.");
        }

        var needsLock = (actionName === 'pull' || actionName === 'push');
        var execute = function() {
            if (actionName === 'pull' || actionName === 'push') {
                _App_ensureSheetExists(toolKey);
            }
            try {
                return action.apply(cfg, args || []);
            } catch (err) {
                var translated = _App_translateApiError(err);
                throw new Error(translated.message);
            }
        };

        if (typeof _App_resetExecutionTimer === 'function') {
            _App_resetExecutionTimer();
        }

        return Logger.run(toolKey, actionName, function () {
            if (needsLock) {
                var lockName = toolKey + "_" + actionName.toUpperCase();
                return _App_withDocumentLock(lockName, execute);
            } else {
                return execute();
            }
        });
    }

    return {
        registerTool: registerTool,
        getTool: getTool,
        getAllTools: getAllTools,
        getToolKeys: getToolKeys,
        runAction: runAction,
        Utils: {
            ok: _App_ok,
            fail: _App_fail,
            withDocumentLock: _App_withDocumentLock,
            include: _App_include,
            createTemplateFromFile: _App_createTemplateFromFile,
            formatStatus: _App_formatStatus,
            throttle: _App_throttle,
            callWithBackoff: _App_callWithBackoff,
            isExecutionLimitApproaching: _App_isExecutionLimitApproaching,
            translateApiError: _App_translateApiError,
            validateRowAgainstSchema: _App_validateRowAgainstSchema,
            BatchProcessor: _App_BatchProcessor,
            batchPatchResults: _App_batchPatchResults,
            extractIdFromUrl: _App_extractIdFromUrl,
            fetchParallel: _App_fetchParallel,
            formatDateTime: _App_formatDateTime,
            sendOrDraftEmail: _App_sendOrDraftEmail
        }
    };
})();

/**
 * Backward compatibility Proxy for legacy scripts still referencing APP_REGISTRY directly.
 */
var APP_REGISTRY = new Proxy({}, {
    get: function(target, prop) {
        return SyncEngine.getTool(prop);
    },
    ownKeys: function() {
        return Object.keys(SyncEngine.getAllTools());
    },
    getOwnPropertyDescriptor: function(target, prop) {
        return { enumerable: true, configurable: true };
    }
});

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

    function auditTool(key) {
        var cfg = getTool(key);
        return _validateToolConfig(key, cfg);
    }

    return {
        registerTool: registerTool,
        getTool: getTool,
        getAllTools: getAllTools,
        getToolKeys: getToolKeys,
        auditTool: auditTool
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

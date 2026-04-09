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

        // Post-process the config (generate HEADERS and totalCols from SCHEMA)
        if (config.FORMAT_CONFIG && config.FORMAT_CONFIG.COL_SCHEMA) {
            config.HEADERS = config.FORMAT_CONFIG.COL_SCHEMA.map(function(c) { return c.header; });
            config.FORMAT_CONFIG.totalCols = config.FORMAT_CONFIG.COL_SCHEMA.length;
        }

        var issues = _validateToolConfig(key, config);
        if (issues.length > 0) {
            throw new Error("Tool '" + key + "' is misconfigured: " + issues.join(' '));
        }

        registry[key] = config;
        // console.log("SyncEngine: Registered " + key);
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

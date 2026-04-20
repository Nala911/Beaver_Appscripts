/**
 * CORE ENGINE & DISPATCHER
 * ==========================================
 * Manages tool registration and provides a centralized execution dispatcher.
 */

Object.assign(App.Engine, (function(global) {
    var registry = {};

    function _validateToolConfig(key, config) {
        var issues = [];

        if (!config.SHEET_NAME) issues.push("Missing SHEET_NAME.");
        if (!config.TITLE) issues.push("Missing TITLE.");

        if (config.MENU_LABEL && !config.MENU_ENTRYPOINT) {
            issues.push("MENU_LABEL requires MENU_ENTRYPOINT.");
        }

        if (config.LAUNCH_MODE === App.Config.TOOL_LAUNCH_MODES.SIDEBAR && !config.SIDEBAR_HTML && config.MENU_ENTRYPOINT) {
            issues.push("Sidebar tools require SIDEBAR_HTML.");
        }

        if (config.LAUNCH_MODE === App.Config.TOOL_LAUNCH_MODES.MODAL && !(config.MODAL_HTML || config.SIDEBAR_HTML) && config.MENU_ENTRYPOINT) {
            issues.push("Modal tools require MODAL_HTML or SIDEBAR_HTML.");
        }

        if (config.FORMAT_CONFIG && config.FORMAT_CONFIG.COL_SCHEMA && !Array.isArray(config.FORMAT_CONFIG.COL_SCHEMA)) {
            issues.push("FORMAT_CONFIG.COL_SCHEMA must be an array.");
        }

        return issues;
    }

    /**
     * Registers a tool with the engine.
     */
    function registerTool(key, config) {
        config.TOOL_KEY = key;
        config.MENU_LABEL = config.MENU_LABEL || config.TITLE;
        config.MENU_ORDER = typeof config.MENU_ORDER === 'number' ? config.MENU_ORDER : 999;
        config.LAUNCH_MODE = config.LAUNCH_MODE || App.Config.TOOL_LAUNCH_MODES.SIDEBAR;

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
    }

    function getTool(key) {
        var cfg = registry[key];
        if (!cfg) throw new Error('Unknown tool key: "' + key + '".');
        return cfg;
    }

    function getAllTools() { return registry; }

    function getToolKeys() { return Object.keys(registry); }

    function auditTool(key) {
        var cfg = registry[key];
        if (!cfg) return ["Tool '" + key + "' is not registered."];
        return _validateToolConfig(key, cfg);
    }

    /**
     * Unified Server-Side Dispatcher
     * Routes requests from Sidebars to the correct tool service.
     */
    function exec(toolKey, action, params) {
        return Logger.run(toolKey, action, function() {
            var cfg = getTool(toolKey);
            
            // Check if the tool has a 'service' object defined (New Pattern)
            if (cfg.service && typeof cfg.service[action] === 'function') {
                return cfg.service[action](params);
            }

            // Fallback: Check for global functions (Old Pattern)
            // Note: This is for transition; eventually all logic should be in cfg.service.
            var legacyFunctionName = toolKey.charAt(0).toUpperCase() + toolKey.slice(1).toLowerCase().replace(/_([a-z])/g, function(g) { return g[1].toUpperCase(); }) + "_" + action;
            if (typeof global[legacyFunctionName] === 'function') {
                 return global[legacyFunctionName](params);
            }

            throw new Error("Action '" + action + "' not found for tool '" + toolKey + "'.");
        });
    }

    return {
        registerTool: registerTool,
        getTool: getTool,
        getAllTools: getAllTools,
        getToolKeys: getToolKeys,
        auditTool: auditTool,
        exec: exec,
        // Preferences helpers
        getPrefs: function(toolKey) {
            var key = "WorkspaceSync_Prefs_" + toolKey;
            var val = PropertiesService.getUserProperties().getProperty(key);
            try { return val ? JSON.parse(val) : {}; } catch (e) { return {}; }
        },
        setPrefs: function(toolKey, prefs) {
            var key = "WorkspaceSync_Prefs_" + toolKey;
            PropertiesService.getUserProperties().setProperty(key, JSON.stringify(prefs || {}));
        }
    };
})(this));

// Backward Compatibility Layer
var SyncEngine = App.Engine;
var APP_REGISTRY = new Proxy({}, {
    get: function(target, prop) { return App.Engine.getTool(prop); },
    ownKeys: function() { return App.Engine.getToolKeys(); },
    getOwnPropertyDescriptor: function() { return { enumerable: true, configurable: true }; }
});

/**
 * Universal Entry Point for Frontend Calls
 */
function App_exec(toolKey, action, params) {
    return App.Engine.exec(toolKey, action, params);
}

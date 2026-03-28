// Default theme definition
var DEFAULT_SHEET_THEME = {
    // Cell Backgrounds
    HEADER: '#424242',
    ACTION: '#2e5a70',
    EDITABLE: '#528dab',
    READ_ONLY: '#655356',

    // Status Colors (Used for conditional formatting rules)
    STATUS: {
        SUCCESS: '#10B981',    // Emerald Green
        PENDING: '#f59e0b',    // Amber/Yellow
        ERROR: '#EF4444',      // Red
        SYNCED: '#6366F1',     // Indigo
        WARNING: '#d59679'
    },

    // Text Colors
    TEXT: '#ffffff',         // Unified light text color for all backgrounds

    // Borders
    BORDER: '#ffffff',       // Soft gray borders instead of harsh black
    BORDER_STYLE: SpreadsheetApp.BorderStyle.SOLID, // Default border style

    // Typography
    FONTS: {
        PRIMARY: 'Roboto',     // Main font for all sheets
        MONOSPACE: 'Consolas'  // Used for IDs, Paths, and technical data
    },

    SIZES: {
        HEADER: 11,            // Header font size
        BODY: 10               // Data body font size
    },

    // Alignment & Layout
    LAYOUT: {
        HEADER_ALIGN_H: 'center',
        HEADER_ALIGN_V: 'middle',
        BODY_ALIGN_H: 'left',
        BODY_ALIGN_V: 'middle',
        BODY_WRAP: SpreadsheetApp.WrapStrategy.CLIP,
        HEADER_WEIGHT: 'bold',
        HEADER_FONT_STYLE: 'normal',
        HEADER_ROW_HEIGHT: 45,
        BODY_ROW_HEIGHT: 35
    }
};

/**
 * Helper to deep merge objects so we don't lose structure if keys are missing from saved theme.
 */
function deepMergeTheme_(target, source) {
    const output = Object.assign({}, target);
    if (isObject_(target) && isObject_(source)) {
        Object.keys(source).forEach(key => {
            if (isObject_(source[key])) {
                if (!(key in target))
                    Object.assign(output, { [key]: source[key] });
                else
                    output[key] = deepMergeTheme_(target[key], source[key]);
            } else {
                Object.assign(output, { [key]: source[key] });
            }
        });
    }
    return output;
}

function isObject_(item) {
    return (item && typeof item === 'object' && !Array.isArray(item));
}

function _UI_getTheme() {
    const savedTheme = _App_getProperty(APP_PROPS.THEME);
    if (savedTheme) {
        try {
            const merged = deepMergeTheme_(DEFAULT_SHEET_THEME, savedTheme);

            // Restore Enum properties that get converted to strings or empty objects during JSON serialization
            const bw = merged.LAYOUT.BODY_WRAP;
            if (typeof bw === 'string') {
                merged.LAYOUT.BODY_WRAP = SpreadsheetApp.WrapStrategy[bw] || SpreadsheetApp.WrapStrategy.CLIP;
            } else if (bw !== SpreadsheetApp.WrapStrategy.CLIP && bw !== SpreadsheetApp.WrapStrategy.WRAP && bw !== SpreadsheetApp.WrapStrategy.OVERFLOW) {
                merged.LAYOUT.BODY_WRAP = SpreadsheetApp.WrapStrategy.CLIP;
            }

            const bs = merged.BORDER_STYLE;
            if (typeof bs === 'string') {
                merged.BORDER_STYLE = SpreadsheetApp.BorderStyle[bs] || SpreadsheetApp.BorderStyle.SOLID;
            } else if (bs !== SpreadsheetApp.BorderStyle.SOLID && bs !== SpreadsheetApp.BorderStyle.SOLID_MEDIUM && bs !== SpreadsheetApp.BorderStyle.SOLID_THICK && bs !== SpreadsheetApp.BorderStyle.DASHED && bs !== SpreadsheetApp.BorderStyle.DOTTED && bs !== SpreadsheetApp.BorderStyle.DOUBLE) {
                merged.BORDER_STYLE = SpreadsheetApp.BorderStyle.SOLID;
            }

            return merged;
        } catch (e) {
            console.error('Failed to parse saved theme, falling back to defaults:', e);
        }
    }
    return DEFAULT_SHEET_THEME;
}
// Global Export! Scripts using SHEET_THEME will get the dynamic version lazily.
// We use a Proxy here so that PropertiesService (a slow API call) is only invoked
// when a script actively accesses the theme, preventing execution delays on all triggers.
var __ui_sheetThemeCache = null;
var SHEET_THEME = new Proxy({}, {
    get: function (target, prop) {
        if (!__ui_sheetThemeCache) {
            __ui_sheetThemeCache = _UI_getTheme(); // Load from PropertiesService only on access
        }
        return Reflect.get(__ui_sheetThemeCache, prop);
    },
    ownKeys: function () {
        if (!__ui_sheetThemeCache) __ui_sheetThemeCache = _UI_getTheme();
        return Reflect.ownKeys(__ui_sheetThemeCache);
    },
    getOwnPropertyDescriptor: function (target, prop) {
        if (!__ui_sheetThemeCache) __ui_sheetThemeCache = _UI_getTheme();
        return Reflect.getOwnPropertyDescriptor(__ui_sheetThemeCache, prop);
    }
});

// Default theme definition
var DEFAULT_SHEET_THEME = {
    // Cell Backgrounds
    HEADER: '#424242',
    FIRST_COLS_COLOR: '#2e5a70',
    MIDDLE_COLS_COLOR: '#528dab',
    LAST_COLS_COLOR: '#314974',

    // Status Colors (Used for conditional formatting rules)
    STATUS: {
        SUCCESS: '#10B981',    // Emerald Green
        PENDING: '#f59e0b',    // Amber/Yellow
        ERROR: '#EF4444',      // Red
        SYNCED: '#6366F1',     // Indigo
        WARNING: '#d59679'
    },

    // Standard Status Prefixes
    STATUS_PREFIXES: {
        SUCCESS: '✅ ',
        ERROR: '❌ ',
        WARNING: '⚠️ ',
        PENDING: '⏳ ',
        INFO: 'ℹ️ '
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

var SHEET_THEME = DEFAULT_SHEET_THEME;

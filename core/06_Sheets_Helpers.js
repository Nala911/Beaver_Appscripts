function _App_canScaffoldSheet(toolConfig) {
    return !!(toolConfig && toolConfig.HEADERS && toolConfig.HEADERS.length);
}
/**
 * Throws an error if the active sheet is not the expected one. Useful for direct action trigger functions.
 * @param {string} expectedSheetName The globally defined sheet name from SHEET_NAMES
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The active sheet object if valid
 */
function _App_assertActiveSheet(expectedSheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== expectedSheetName) {
        throw new Error("⚠️ Please run this action from the '" + expectedSheetName + "' sheet.");
    }
    return sheet;
}
/**
 * Returns a validation object. Useful for UI-triggered functions that need to return an error shape `{success: false, message: ...}` instead of failing ungracefully.
 * @param {string} expectedSheetName 
 * @returns {Object} `{ valid: boolean, sheet: Sheet, message?: string }`
 */
function _App_validateActiveSheet(expectedSheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== expectedSheetName) {
        return { valid: false, message: "⚠️ Please run this action from the '" + expectedSheetName + "' sheet." };
    }
    return { valid: true, sheet: sheet };
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🦫 Beaver Tools');

  _App_getMenuTools().forEach(function(cfg) {
    if (cfg.MENU_ENTRYPOINT) {
      menu.addItem(cfg.MENU_LABEL, cfg.MENU_ENTRYPOINT);
    }
  });

  menu
    .addSeparator()
    .addItem('⚙️ Theme Settings', 'UI_openThemeDialog')
    .addToUi();
}

// Global settings and Theme Config have been moved to 00_Config_Constants.js and 01_Config_Theme.js
// so that Apps Script evaluates them before the rest of the files.


// ==========================================
// Theme Dialog Server-Side Functions
// ==========================================

function Logger_showSidebar() {
  _App_launchTool('LOGS');
}

function UI_openThemeDialog() {
  const html = HtmlService.createTemplateFromFile('ThemeEditorSidebar').evaluate()
    .setTitle('🎨 Theme Studio')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '🎨 Theme Studio');
}

function UI_getThemeConfig() {
  // We send back both the current dynamic theme, and the default theme so the frontend knows what to compare
  return {
    current: _UI_themeToClient(_UI_getTheme()),
    defaults: _UI_themeToClient(DEFAULT_SHEET_THEME)
  };
}

/**
 * Converts a theme object into a client-safe format by replacing
 * SpreadsheetApp enum values with their string key names so they
 * serialize cleanly over google.script.run.
 */
function _UI_themeToClient(theme) {
  var safe = JSON.parse(JSON.stringify(theme));
  // BorderStyle enum → string
  var bs = theme.BORDER_STYLE;
  if (bs === SpreadsheetApp.BorderStyle.SOLID) safe.BORDER_STYLE = 'SOLID';
  else if (bs === SpreadsheetApp.BorderStyle.SOLID_MEDIUM) safe.BORDER_STYLE = 'SOLID_MEDIUM';
  else if (bs === SpreadsheetApp.BorderStyle.SOLID_THICK) safe.BORDER_STYLE = 'SOLID_THICK';
  else if (bs === SpreadsheetApp.BorderStyle.DASHED) safe.BORDER_STYLE = 'DASHED';
  else if (bs === SpreadsheetApp.BorderStyle.DOTTED) safe.BORDER_STYLE = 'DOTTED';
  else if (bs === SpreadsheetApp.BorderStyle.DOUBLE) safe.BORDER_STYLE = 'DOUBLE';
  else if (typeof bs === 'string') safe.BORDER_STYLE = bs;
  else safe.BORDER_STYLE = 'SOLID';
  // WrapStrategy enum → string
  if (theme.LAYOUT && theme.LAYOUT.BODY_WRAP !== undefined) {
    var bw = theme.LAYOUT.BODY_WRAP;
    if (bw === SpreadsheetApp.WrapStrategy.CLIP) safe.LAYOUT.BODY_WRAP = 'CLIP';
    else if (bw === SpreadsheetApp.WrapStrategy.WRAP) safe.LAYOUT.BODY_WRAP = 'WRAP';
    else if (bw === SpreadsheetApp.WrapStrategy.OVERFLOW) safe.LAYOUT.BODY_WRAP = 'OVERFLOW';
    else if (typeof bw === 'string') safe.LAYOUT.BODY_WRAP = bw;
    else safe.LAYOUT.BODY_WRAP = 'CLIP';
  }
  return safe;
}

function UI_saveThemeConfig(newThemeConfig) {
  try {
    // Validate that modifying basic properties works
    if (newThemeConfig) {
      _App_setProperty(APP_PROPS.THEME, newThemeConfig);
      Logger.info("Theme Settings", "Save Theme", "Theme saved successfully");
      Logger.flushLogs();
      return { success: true, message: 'Theme saved successfully!' };
    }
  } catch (e) {
    Logger.error("Theme Settings", "Save Theme", e);
    Logger.flushLogs();
    return { success: false, message: 'Failed to save theme: ' + e.toString() };
  }
}

function UI_resetThemeConfig() {
  try {
    _App_deleteProperty(APP_PROPS.THEME);
    Logger.info("Theme Settings", "Reset Theme", "Reset to default theme");
    Logger.flushLogs();
    return { success: true, message: 'Reset to default theme!' };
  } catch (e) {
    Logger.error("Theme Settings", "Reset Theme", e);
    Logger.flushLogs();
    return { success: false, message: 'Failed to reset theme: ' + e.toString() };
  }
}

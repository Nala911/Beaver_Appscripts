/**
 * Forms Sync Tool
 * Version: 5.0 (Plugin Architecture — registers with BeaverEngine)
 * Syncs questions and options between Google Sheets and Google Forms
 */

BeaverEngine.registerTool('FORMS_SYNC', {
    SHEET_NAME: SHEET_NAMES.FORMS_SYNC,
    TITLE: '📝 Forms Sync',
    MENU_LABEL: '📝 Google Forms',
    MENU_ENTRYPOINT: 'FormsSync_openSidebar',
    MENU_ORDER: 70,
    SIDEBAR_HTML: 'FormsSync_Sidebar',
    SIDEBAR_WIDTH: 300,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [100, 300, 150, 250, 250, 100, 120],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 1,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'REMOVE'] },
            { header: 'Question Title', type: 'TEXT' },
            { header: 'Type', type: 'DROPDOWN', options: ['MULTIPLE_CHOICE', 'CHECKBOX', 'LIST', 'TEXT', 'PARAGRAPH_TEXT', 'DATE', 'TIME', 'DATETIME', 'DURATION', 'SCALE', 'GRID', 'CHECKBOX_GRID', 'FILE_UPLOAD', 'PAGE_BREAK', 'SECTION_HEADER', 'IMAGE', 'VIDEO'] },
            { header: 'Options', type: 'TEXT' },
            { header: 'Help Text', type: 'TEXT' },
            { header: 'Required', type: 'CHECKBOX' },
            { header: 'Item ID', type: 'ID' }
        ]
    }
});

// --- CONFIGURATION ---
// Column-index aliases (1-based) — kept for backward compatibility.
// Tool metadata now lives in BeaverEngine.getTool('FORMS_SYNC').
var FORMSSYNC_CFG = {
    COLUMNS: {
        ACTION: 1, TITLE: 2, TYPE: 3, OPTIONS: 4, HELP_TEXT: 5, REQUIRED: 6, ID: 7
    },
    HEADER_ROW: 1
};

// --- SHEET SETUP LOGIC ---
/** @deprecated — Use _App_ensureSheetExists('FORMS_SYNC') instead. */
function _FormsSync_ensureSheetExistsAndActivate() {
    return _App_ensureSheetExists('FORMS_SYNC');
}

/** Opens the Forms Sync sidebar and ensures the sheet exists. */
function FormsSync_openSidebar() {
    return Logger.run('FORMS_SYNC', 'Open Sidebar', function () {
        _App_launchTool('FORMS_SYNC');
    });
}

// --- CORE HELPER LOGIC ---



function _FormsSync_extractFormId(inputUrlOrId) {
    if (!inputUrlOrId) return null;
    var match = inputUrlOrId.match(/\/d\/(.*?)(\/|$)/);
    return match ? match[1] : inputUrlOrId; // If match fails, assume it's already an ID
}

function _FormsSync_pullForm(formInput) {
    return Logger.run('FORMS_SYNC', 'Pull Form', function () {
        var formId = _FormsSync_extractFormId(formInput);
        if (!formId) return { success: false, message: "Invalid Form URL or ID" };

        try {
            var form = _App_callWithBackoff(function () { return FormApp.openById(formId); });
            var items = _App_callWithBackoff(function () { return form.getItems(); });

            var sheetData = [];

            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                var id = item.getId().toString();
                var title = item.getTitle() || "";
                var type = item.getType().toString();
                var options = "";
                var helpText = item.getHelpText() || "";
                var required = false;

                // Extract type-specific properties (options, required)
                try {
                    if (type === "MULTIPLE_CHOICE") {
                        var mcItem = item.asMultipleChoiceItem();
                        required = mcItem.isRequired();
                        options = mcItem.getChoices().map(function (c) { return c.getValue(); }).join("\n");
                    } else if (type === "CHECKBOX") {
                        var cbItem = item.asCheckboxItem();
                        required = cbItem.isRequired();
                        options = cbItem.getChoices().map(function (c) { return c.getValue(); }).join("\n");
                    } else if (type === "LIST") {
                        var liItem = item.asListItem();
                        required = liItem.isRequired();
                        options = liItem.getChoices().map(function (c) { return c.getValue(); }).join("\n");
                    } else if (type === "TEXT") {
                        required = item.asTextItem().isRequired();
                    } else if (type === "PARAGRAPH_TEXT") {
                        required = item.asParagraphTextItem().isRequired();
                    } else if (type === "DATE") {
                        required = item.asDateItem().isRequired();
                    } else if (type === "TIME") {
                        required = item.asTimeItem().isRequired();
                    } else if (type === "DATETIME") {
                        required = item.asDateTimeItem().isRequired();
                    } else if (type === "DURATION") {
                        required = item.asDurationItem().isRequired();
                    } else if (type === "SCALE") {
                        required = item.asScaleItem().isRequired();
                    } else if (type === "GRID") {
                        var gridItem = item.asGridItem();
                        required = gridItem.isRequired();
                        var gridRows = gridItem.getRows() || [];
                        var gridCols = gridItem.getColumns() || [];
                        options = gridRows.join("\n") + "\n||\n" + gridCols.join("\n");
                    } else if (type === "CHECKBOX_GRID") {
                        var cbGridItem = item.asCheckboxGridItem();
                        required = cbGridItem.isRequired();
                        var cbGridRows = cbGridItem.getRows() || [];
                        var cbGridCols = cbGridItem.getColumns() || [];
                        options = cbGridRows.join("\n") + "\n||\n" + cbGridCols.join("\n");
                    }
                } catch (propErr) {
                    Logger.warn(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'Property Error', "Error reading item properties for ID " + id + ": " + propErr);
                }

                sheetData.push(["", title, type, options, helpText, required, id]);
            }

            var sheet = _FormsSync_ensureSheetExistsAndActivate();

            // Clear old data
            var lastRow = sheet.getLastRow();
            if (lastRow > FORMSSYNC_CFG.HEADER_ROW) {
                sheet.getRange(FORMSSYNC_CFG.HEADER_ROW + 1, 1, lastRow - FORMSSYNC_CFG.HEADER_ROW, sheet.getLastColumn()).clearContent();
                sheet.getRange(FORMSSYNC_CFG.HEADER_ROW + 1, FORMSSYNC_CFG.COLUMNS.REQUIRED, lastRow - FORMSSYNC_CFG.HEADER_ROW, 1).removeCheckboxes();
            }

            // Set New Data
            if (sheetData.length > 0) {
                var targetRange = sheet.getRange(FORMSSYNC_CFG.HEADER_ROW + 1, 1, sheetData.length, sheetData[0].length);
                targetRange.setValues(sheetData);
            }

            // Apply body formatting via shared utility
            _App_applyBodyFormatting(sheet, sheetData.length, BeaverEngine.getTool('FORMS_SYNC').FORMAT_CONFIG);

            // Save Form ID to PropertiesService for syncing back
            _App_setProperty(APP_PROPS.FORMS_CURRENT_FORM, formId);
            // Save to UserProperties for sidebar auto-selection
            _App_setProperty(APP_PROPS.FORMS_SELECTED_FORM, formId);

            return { success: true, message: "Successfully pulled " + sheetData.length + " items." };
        } catch (e) {
            throw e;
        }
    });
}

function _FormsSync_syncToForm() {
    return Logger.run('FORMS_SYNC', 'Sync to Form', function () {
        var formId = _App_getProperty(APP_PROPS.FORMS_CURRENT_FORM);
        if (!formId) return { success: false, message: "No form connected. Please Pull data first." };

        try {
            var form = _App_callWithBackoff(function () { return FormApp.openById(formId); });
            var sheet = _FormsSync_ensureSheetExistsAndActivate();
            var dataRange = sheet.getDataRange();
            var data = dataRange.getValues();

            if (data.length <= FORMSSYNC_CFG.HEADER_ROW) {
                return { success: true, message: "No data to sync." };
            }

            var rowUpdates = [];

            // Loop through data starting after header
            for (var i = FORMSSYNC_CFG.HEADER_ROW; i < data.length; i++) {
                var row = data[i];
                var action = (row[FORMSSYNC_CFG.COLUMNS.ACTION - 1] || "").toString().trim().toUpperCase();
                var id = (row[FORMSSYNC_CFG.COLUMNS.ID - 1] || "").toString().trim();
                var title = (row[FORMSSYNC_CFG.COLUMNS.TITLE - 1] || "").toString();
                var type = (row[FORMSSYNC_CFG.COLUMNS.TYPE - 1] || "").toString();
                var optionsRaw = (row[FORMSSYNC_CFG.COLUMNS.OPTIONS - 1] || "").toString();
                var helpText = (row[FORMSSYNC_CFG.COLUMNS.HELP_TEXT - 1] || "").toString();
                var required = row[FORMSSYNC_CFG.COLUMNS.REQUIRED - 1] === true || row[FORMSSYNC_CFG.COLUMNS.REQUIRED - 1] === 'TRUE';

                var updateObj = {
                    action: action,
                    id: id
                };

                if (!action || (action !== "CREATE" && action !== "UPDATE" && action !== "REMOVE")) {
                    rowUpdates.push(updateObj);
                    continue;
                }

                var optionsArr = [];
                var gridRows = [];
                var gridCols = [];

                if (type === "GRID" || type === "CHECKBOX_GRID") {
                    // Parse "Row1\nRow2\n||\nCol1\nCol2" format
                    var gridParts = optionsRaw.split("||");
                    gridRows = (gridParts[0] || "").split("\n").map(function (s) { return s.trim(); }).filter(function (s) { return s.length > 0; });
                    gridCols = (gridParts[1] || "").split("\n").map(function (s) { return s.trim(); }).filter(function (s) { return s.length > 0; });
                } else {
                    optionsArr = optionsRaw ? optionsRaw.split("\n").map(function (o) { return o.trim(); }).filter(function (o) { return o.length > 0; }) : [];
                }

                try {
                    if (action === "CREATE") {
                        if (!title) throw new Error("Missing Title");
                        var targetItem = null;

                        _App_callWithBackoff(function () {
                            if (type === "MULTIPLE_CHOICE") targetItem = form.addMultipleChoiceItem();
                            else if (type === "CHECKBOX") targetItem = form.addCheckboxItem();
                            else if (type === "LIST") targetItem = form.addListItem();
                            else if (type === "TEXT") targetItem = form.addTextItem();
                            else if (type === "PARAGRAPH_TEXT") targetItem = form.addParagraphTextItem();
                            else if (type === "DATE") targetItem = form.addDateItem();
                            else if (type === "TIME") targetItem = form.addTimeItem();
                            else if (type === "DATETIME") targetItem = form.addDateTimeItem();
                            else if (type === "DURATION") targetItem = form.addDurationItem();
                            else if (type === "SCALE") targetItem = form.addScaleItem();
                            else if (type === "GRID") targetItem = form.addGridItem();
                            else if (type === "CHECKBOX_GRID") targetItem = form.addCheckboxGridItem();
                            else {
                                targetItem = form.addTextItem();
                                type = "TEXT";
                            }
                        });

                        _App_callWithBackoff(function () {
                            targetItem.setTitle(title);
                            targetItem.setHelpText(helpText);
                            _applyItemProperties(targetItem, type, required, optionsArr, gridRows, gridCols);
                        });

                        updateObj.id = targetItem.getId().toString();
                        Logger.info(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'Item: ' + title, '✅ Created');
                        updateObj.action = "";
                    }
                    else if (action === "UPDATE") {
                        if (!id) throw new Error("Missing ID");
                        var updItem = _App_callWithBackoff(function () { return form.getItemById(parseInt(id, 10)); });
                        if (!updItem) throw new Error("Item ID not found");

                        var currentType = updItem.getType().toString();

                        if (currentType === type) {
                            // Type matches, safely apply properties
                            _App_callWithBackoff(function () {
                                updItem.setTitle(title);
                                updItem.setHelpText(helpText);
                                _applyItemProperties(updItem, type, required, optionsArr, gridRows, gridCols);
                            });
                            Logger.info(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'Item: ' + title, '✅ Updated');
                        } else {
                            // Type changed! Google Forms API doesn't allow changing types of existing items.
                            // We must cache the index, delete the old, and create a new item of the target type.
                            var targetIndex = updItem.getIndex();
                            _App_callWithBackoff(function () { form.deleteItem(updItem); });

                            var newItem = null;
                            _App_callWithBackoff(function () {
                                if (type === "MULTIPLE_CHOICE") newItem = form.addMultipleChoiceItem();
                                else if (type === "CHECKBOX") newItem = form.addCheckboxItem();
                                else if (type === "LIST") newItem = form.addListItem();
                                else if (type === "TEXT") newItem = form.addTextItem();
                                else if (type === "PARAGRAPH_TEXT") newItem = form.addParagraphTextItem();
                                else if (type === "DATE") newItem = form.addDateItem();
                                else if (type === "TIME") newItem = form.addTimeItem();
                                else if (type === "DATETIME") newItem = form.addDateTimeItem();
                                else if (type === "DURATION") newItem = form.addDurationItem();
                                else if (type === "SCALE") newItem = form.addScaleItem();
                                else if (type === "GRID") newItem = form.addGridItem();
                                else if (type === "CHECKBOX_GRID") newItem = form.addCheckboxGridItem();
                                else {
                                    newItem = form.addTextItem();
                                    type = "TEXT";
                                }
                            });

                            _App_callWithBackoff(function () {
                                newItem.setTitle(title);
                                newItem.setHelpText(helpText);
                                _applyItemProperties(newItem, type, required, optionsArr, gridRows, gridCols);
                                // Move the newly created item to the old item's exact position
                                form.moveItem(newItem.getIndex(), targetIndex);
                            });

                            updateObj.id = newItem.getId().toString();
                            Logger.info(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'Item: ' + title, '✅ Updated (Type Recreated)');
                        }
                        updateObj.action = "";
                    }
                    else if (action === "REMOVE") {
                        if (!id) throw new Error("Missing ID");
                        var delItem = _App_callWithBackoff(function () { return form.getItemById(parseInt(id, 10)); });
                        if (delItem) {
                            _App_callWithBackoff(function () { form.deleteItem(delItem); });
                            Logger.info(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'ID Map: ' + id, '🗑️ Removed');
                        } else {
                            Logger.info(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'ID Map: ' + id, '⚠️ Already Deleted');
                        }
                        updateObj.action = "";
                    }
                } catch (err) {
                    Logger.error(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'Row ' + i, err);
                }

                rowUpdates.push(updateObj);
            }

            // Write updates back to sheet
            rowUpdates.forEach(function (upd, idx) {
                var rowNum = FORMSSYNC_CFG.HEADER_ROW + 1 + idx;
                sheet.getRange(rowNum, FORMSSYNC_CFG.COLUMNS.ACTION).setValue(upd.action);
                if (upd.id) sheet.getRange(rowNum, FORMSSYNC_CFG.COLUMNS.ID).setValue(upd.id);
            });

            return { success: true, message: "Sync Complete. Check the Developer Log for details." };
        } catch (e) {
            throw e;
        }
    });
}

function _applyItemProperties(targetItem, type, required, optionsArr, gridRows, gridCols) {
    try {
        if (type === "MULTIPLE_CHOICE") {
            var mcItem = targetItem.asMultipleChoiceItem();
            mcItem.setRequired(required);
            if (optionsArr.length > 0) _setChoicesSafe(mcItem, optionsArr, "MULTIPLE_CHOICE");
        } else if (type === "CHECKBOX") {
            var cbItem = targetItem.asCheckboxItem();
            cbItem.setRequired(required);
            if (optionsArr.length > 0) _setChoicesSafe(cbItem, optionsArr, "CHECKBOX");
        } else if (type === "LIST") {
            var liItem = targetItem.asListItem();
            liItem.setRequired(required);
            if (optionsArr.length > 0) _setChoicesSafe(liItem, optionsArr, "LIST");
        } else if (type === "TEXT") {
            targetItem.asTextItem().setRequired(required);
        } else if (type === "PARAGRAPH_TEXT") {
            targetItem.asParagraphTextItem().setRequired(required);
        } else if (type === "DATE") {
            targetItem.asDateItem().setRequired(required);
        } else if (type === "TIME") {
            targetItem.asTimeItem().setRequired(required);
        } else if (type === "DATETIME") {
            targetItem.asDateTimeItem().setRequired(required);
        } else if (type === "DURATION") {
            targetItem.asDurationItem().setRequired(required);
        } else if (type === "GRID") {
            var gridItem = targetItem.asGridItem();
            gridItem.setRequired(required);
            if (gridRows && gridRows.length > 0) gridItem.setRows(gridRows);
            if (gridCols && gridCols.length > 0) gridItem.setColumns(gridCols);
        } else if (type === "CHECKBOX_GRID") {
            var cbGridItem = targetItem.asCheckboxGridItem();
            cbGridItem.setRequired(required);
            if (gridRows && gridRows.length > 0) cbGridItem.setRows(gridRows);
            if (gridCols && gridCols.length > 0) cbGridItem.setColumns(gridCols);
        }
    } catch (e) {
        Logger.warn(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'Apply Properties', "Failed to apply properties", e);
    }
}

function _setChoicesSafe(item, optionsArr, type) {
    if (!optionsArr || optionsArr.length === 0) return;

    try {
        var hasOther = false;
        if (typeof item.hasOtherOption === 'function') {
            hasOther = item.hasOtherOption();
        }

        var existingChoices = [];
        if (typeof item.getChoices === 'function') {
            existingChoices = item.getChoices();
        }

        var choices = [];
        for (var i = 0; i < optionsArr.length; i++) {
            var optString = optionsArr[i];
            var added = false;

            if (i < existingChoices.length) {
                var ec = existingChoices[i];
                if (typeof ec.getPageNavigationType === 'function') {
                    var navType = ec.getPageNavigationType();
                    if (navType === FormApp.PageNavigationType.GO_TO_PAGE) {
                        var gotoPage = ec.getGotoPage();
                        if (gotoPage && typeof item.createChoice === 'function') {
                            try {
                                if (type === "MULTIPLE_CHOICE" || type === "LIST") {
                                    choices.push(item.createChoice(optString, gotoPage));
                                    added = true;
                                }
                            } catch (e) { }
                        }
                    } else if (navType) {
                        try {
                            if (type === "MULTIPLE_CHOICE" || type === "LIST") {
                                choices.push(item.createChoice(optString, navType));
                                added = true;
                            }
                        } catch (e) { }
                    }
                }
            }
            if (!added && typeof item.createChoice === 'function') {
                choices.push(item.createChoice(optString));
            }
        }

        if (choices.length > 0 && typeof item.setChoices === 'function') {
            item.setChoices(choices);
        }

        if (hasOther && typeof item.showOtherOption === 'function') {
            item.showOtherOption(true);
        }
    } catch (e) {
        Logger.warn(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'Set Choices', "Failed to set choices", e);
    }
}
// --- PUBLIC ENTRY POINTS ---



function FormsSync_getForms() {
    return Logger.run('FORMS_SYNC', 'Get Forms', function () {
        try {
            var files = DriveApp.searchFiles("mimeType='application/vnd.google-apps.form' and trashed=false");
            var forms = [];
            var count = 0;
            var MAX_FORMS = 10;

            while (files.hasNext() && count < MAX_FORMS) {
                var file = files.next();
                forms.push({
                    id: file.getId(),
                    title: file.getName() || "Untitled Form",
                    lastUpdated: file.getLastUpdated().getTime()
                });
                count++;
            }

            forms.sort(function (a, b) {
                return b.lastUpdated - a.lastUpdated;
            });

            var mappedForms = forms.map(function (f) {
                return { id: f.id, title: f.title };
            });

            var savedFormId = _App_getProperty(APP_PROPS.FORMS_SELECTED_FORM);

            return { forms: mappedForms, savedFormId: savedFormId };
        } catch (e) {
            Logger.error(BeaverEngine.getTool('FORMS_SYNC').TITLE, 'Get Forms', e);
            throw new Error("Failed to fetch forms: " + e.toString());
        }
    });
}

function FormsSync_pullForm(formInput) {
    return _FormsSync_pullForm(formInput);
}

function FormsSync_syncToForm() {
    return _FormsSync_syncToForm();
}

function FormsSync_getFormLinks() {
    return Logger.run('FORMS_SYNC', 'Get Form Links', function () {
        var formId = _App_getProperty(APP_PROPS.FORMS_CURRENT_FORM);
        if (!formId) return null;
        try {
            var form = FormApp.openById(formId);
            return {
                editUrl: form.getEditUrl(),
                responsesUrl: form.getSummaryUrl()
            };
        } catch (e) {
            return {
                editUrl: 'https://docs.google.com/forms/d/' + formId + '/edit',
                responsesUrl: 'https://docs.google.com/forms/d/' + formId + '/edit#responses'
            };
        }
    });
}

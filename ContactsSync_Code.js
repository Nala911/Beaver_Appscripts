/**
 * MASTER EDITION: Google Sheets <-> Google Contacts Sync
 * Version: 5.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CONTACTS_SYNC', {
    REQUIRED_SERVICES: [ { name: 'People API', test: function() { return typeof People !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CONTACTS_SYNC,
    TITLE: '📇 Contacts Sync Master',
    MENU_LABEL: '☎️ Google Contacts',
    MENU_ENTRYPOINT: 'ContactsSync_openSidebar',
    MENU_ORDER: 20,
    SIDEBAR_HTML: 'ContactsSync_Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [120, 130, 130, 180, 140, 140, 140, 100, 150, 120, 100, 80, 160, 250, 140],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 1,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'REMOVE'] },
            { header: 'First Name', type: 'TEXT' },
            { header: 'Last Name', type: 'TEXT' },
            { header: 'Email', type: 'TEXT' },
            { header: 'Phone', type: 'TEXT' },
            { header: 'Company', type: 'TEXT' },
            { header: 'Job Title', type: 'TEXT' },
            { header: 'Starred', type: 'DROPDOWN', options: ['Yes', 'No'] },
            { header: 'Street', type: 'TEXT' },
            { header: 'City', type: 'TEXT' },
            { header: 'State', type: 'TEXT' },
            { header: 'Zip', type: 'TEXT' },
            { header: 'Groups/Labels', type: 'TEXT' },
            { header: 'Notes', type: 'TEXT' },
            { header: 'Contact ID', type: 'ID', italic: true }
        ]
    }
});

// Column-index aliases — kept for backward compatibility within this file.
// Metadata (title, sidebar, headers, widths) now lives in SyncEngine.getTool('CONTACTS_SYNC').
var CONTACTS_SYNC_CFG = {
    COLUMNS: {
        ACTION: 0, FIRST_NAME: 1, LAST_NAME: 2,
        EMAIL: 3, PHONE: 4, COMPANY: 5, JOB_TITLE: 6, STARRED: 7,
        STREET: 8, CITY: 9, STATE: 10, ZIP: 11, GROUPS: 12, NOTES: 13, CONTACT_ID: 14
    },
    HEADER_ROW: 1
};

// Declarative format config now lives in SyncEngine.getTool('CONTACTS_SYNC').FORMAT_CONFIG

/** @deprecated — Use _App_ensureSheetExists('CONTACTS_SYNC') instead. */
function _ContactsSync_ensureSheetExistsAndActivate() {
    return _App_ensureSheetExists('CONTACTS_SYNC');
}

/** Opens the Contacts sidebar and ensures the sheet exists. */
function ContactsSync_openSidebar() {
  return Logger.run('CONTACTS_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CONTACTS_SYNC');
  });
}


function ContactsSync_getLoadData() {
    return Logger.run('CONTACTS_SYNC', 'Load Data', function () {
        if (typeof People === 'undefined') {
            throw new Error("⚠️ People API is not enabled. Go to Services -> Add 'People API'.");
        }

        try {
            var groupsResponse = _App_callWithBackoff(function () {
                return People.ContactGroups.list({ pageSize: 1000 });
            });
            var groups = groupsResponse.contactGroups || [];

            var excluded = ['Friends', 'Family', 'Coworkers', 'All Contacts', 'Chat contacts'];
            var formattedGroups = groups.map(function (g) {
                return { id: g.resourceName, name: g.formattedName || g.name };
            }).filter(function (g) {
                return g.id && g.name && !excluded.includes(g.name);
            });

            formattedGroups.unshift({ id: 'all', name: 'All Contacts' });

            var savedGroupIds = _App_getProperty(APP_PROPS.CONTACTS_SELECTED_GROUPS);
            if (!Array.isArray(savedGroupIds) || savedGroupIds.length === 0) {
                savedGroupIds = ['all'];
            }

            return _App_ok('Contacts load data ready.', {
                groups: formattedGroups,
                savedGroupIds: savedGroupIds
            });
        } catch (err) {
            throw new Error('Unable to load contact groups. ' + err.message);
        }
    });
}

function ContactsSync_savePreferences(groupIds) {
    if (groupIds) _App_setProperty(APP_PROPS.CONTACTS_SELECTED_GROUPS, groupIds);
}

function _ContactsSync_getPrimary(array) {
    if (!array || array.length === 0) return "";
    var primary = array.find(function (item) { return item.metadata && item.metadata.primary; });
    return primary ? primary.value : array[0].value;
}



function ContactsSync_pullContacts(request) {
    return Logger.run('CONTACTS_SYNC', 'Pull Contacts', function () {
        if (typeof People === 'undefined') {
            throw new Error("⚠️ People API is not enabled. Go to Services -> Add 'People API'.");
        }

        var sheet = SheetManager.ensureSheet('CONTACTS_SYNC');

        var outputData = [];
        var groupIds = request.groupIds || [];
        var pullAll = groupIds.includes('all');

        var groupsResponse = People.ContactGroups.list();
        var allGroups = groupsResponse.contactGroups || [];
        var groupMap = {};
        allGroups.forEach(function (g) {
            groupMap[g.resourceName] = g.formattedName || g.name;
        });

        var pageToken = null;
        var personFields = 'names,emailAddresses,phoneNumbers,organizations,memberships,biographies,addresses';

        do {
            var options = { pageSize: 1000, personFields: personFields };
            if (pageToken) options.pageToken = pageToken;

            var response;
            try {
                response = _App_callWithBackoff(function () {
                    return People.People.Connections.list('people/me', options);
                });
            } catch (err) {
                throw new Error("API Error: " + err.message);
            }

            var connections = response.connections || [];

            connections.forEach(function (person) {
                var pGroups = person.memberships ? person.memberships.map(function (m) {
                    return m.contactGroupMembership ? m.contactGroupMembership.contactGroupResourceName : null;
                }).filter(function (g) { return g; }) : [];

                var isInSelectedGroup = pullAll || pGroups.some(function (g) { return groupIds.includes(g); });

                if (isInSelectedGroup) {
                    var firstName = "";
                    var lastName = "";
                    if (person.names && person.names.length > 0) {
                        var primaryName = person.names.find(function (n) { return n.metadata && n.metadata.primary; }) || person.names[0];
                        firstName = primaryName.givenName || "";
                        lastName = primaryName.familyName || "";
                    }

                    var email = _ContactsSync_getPrimary(person.emailAddresses);
                    var phone = _ContactsSync_getPrimary(person.phoneNumbers);

                    var company = "";
                    var title = "";
                    if (person.organizations && person.organizations.length > 0) {
                        var primaryOrg = person.organizations.find(function (o) { return o.metadata && o.metadata.primary; }) || person.organizations[0];
                        company = primaryOrg.name || "";
                        title = primaryOrg.title || "";
                    }

                    var notes = person.biographies && person.biographies.length > 0 ? (person.biographies[0].value || "") : "";

                    var isStarred = pGroups.includes('contactGroups/starred') ? "Yes" : "No";

                    var street = "", city = "", state = "", zip = "";
                    if (person.addresses && person.addresses.length > 0) {
                        var primaryAddress = person.addresses.find(function (a) { return a.metadata && a.metadata.primary; }) || person.addresses[0];
                        street = primaryAddress.streetAddress || "";
                        city = primaryAddress.city || "";
                        state = primaryAddress.region || "";
                        zip = primaryAddress.postalCode || "";
                    }

                    var groupNames = pGroups.map(function (gId) { return groupMap[gId] || "Unknown Group"; }).join(", ");

                    outputData.push([
                        "", // Action
                        firstName,
                        lastName,
                        email,
                        phone,
                        company,
                        title,
                        isStarred,
                        street,
                        city,
                        state,
                        zip,
                        groupNames,
                        notes,
                        person.resourceName // Contact ID
                    ]);
                }
            });
            pageToken = response.nextPageToken;
        } while (pageToken);

        // Apply body formatting with duplicate highlighting
        var formatConfig = JSON.parse(JSON.stringify(SyncEngine.getTool('CONTACTS_SYNC').FORMAT_CONFIG));
        formatConfig.conditionalRules = formatConfig.conditionalRules.concat([
            { type: 'custom', formula: '=AND($F2<>"", COUNTIF($F:$F, $F2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 6 },
            { type: 'custom', formula: '=AND($G2<>"", COUNTIF($G:$G, $G2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 7 }
        ]);
        SheetManager.overwriteRows('CONTACTS_SYNC', outputData, {
            totalCols: SyncEngine.getTool('CONTACTS_SYNC').HEADERS.length,
            formatConfig: formatConfig
        });

        ContactsSync_savePreferences(groupIds);
        return _App_ok('Successfully imported ' + outputData.length + " contacts.");
    });
}

// Kept for backward compatibility — now delegates to shared utility
function _ContactsSync_highlightDuplicates(sheet) {
    var lastRow = sheet.getLastRow();
    var numDataRows = lastRow > 1 ? lastRow - 1 : 0;
    var formatConfig = JSON.parse(JSON.stringify(SyncEngine.getTool('CONTACTS_SYNC').FORMAT_CONFIG));
    formatConfig.conditionalRules = formatConfig.conditionalRules.concat([
        { type: 'custom', formula: '=AND($F2<>"", COUNTIF($F:$F, $F2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 6 },
        { type: 'custom', formula: '=AND($G2<>"", COUNTIF($G:$G, $G2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 7 }
    ]);
    _App_applyBodyFormatting(sheet, numDataRows, formatConfig);
}

function ContactsSync_checkForUnsavedChanges() {
    return SheetManager.hasPendingActions('CONTACTS_SYNC');
}

function ContactsSync_pushChanges() {
    return Logger.run('CONTACTS_SYNC', 'Push Changes', function () {
        if (typeof People === 'undefined') {
            throw new Error("⚠️ People API is not enabled. Go to Services -> Add 'People API'.");
        }

        var sheet = _App_assertActiveSheet(SHEET_NAMES.CONTACTS_SYNC);

        var dataRange = sheet.getDataRange();
        var data = dataRange.getValues();
        var headers = data[0];
        var allRows = data.slice(1);

        var groupsResponse = People.ContactGroups.list();
        var allGroups = groupsResponse.contactGroups || [];
        var groupNameToId = {};
        allGroups.forEach(function (g) {
            groupNameToId[g.formattedName || g.name] = g.resourceName;
        });

        var pendingRows = [];
        allRows.forEach(function (row, idx) {
            if (row[CONTACTS_SYNC_CFG.COLUMNS.ACTION]) {
                pendingRows.push({ data: row, originalIndex: idx });
            }
        });

        if (pendingRows.length === 0) return "No pending actions found.";

        var stats = _App_BatchProcessor('CONTACTS_SYNC', pendingRows, function (item) {
            var row = item.data;
            var originalIdx = item.originalIndex;
            var rowUpdates = {
                action: row[CONTACTS_SYNC_CFG.COLUMNS.ACTION],
                contactId: row[CONTACTS_SYNC_CFG.COLUMNS.CONTACT_ID],
                status: "",
                originalIndex: originalIdx
            };

            try {
                var action = rowUpdates.action.toString().toUpperCase();
                var contactData = {
                    firstName: row[CONTACTS_SYNC_CFG.COLUMNS.FIRST_NAME] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.FIRST_NAME]) : "",
                    lastName: row[CONTACTS_SYNC_CFG.COLUMNS.LAST_NAME] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.LAST_NAME]) : "",
                    email: row[CONTACTS_SYNC_CFG.COLUMNS.EMAIL] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.EMAIL]) : "",
                    phone: row[CONTACTS_SYNC_CFG.COLUMNS.PHONE] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.PHONE]) : "",
                    company: row[CONTACTS_SYNC_CFG.COLUMNS.COMPANY] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.COMPANY]) : "",
                    title: row[CONTACTS_SYNC_CFG.COLUMNS.JOB_TITLE] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.JOB_TITLE]) : "",
                    starred: row[CONTACTS_SYNC_CFG.COLUMNS.STARRED],
                    street: row[CONTACTS_SYNC_CFG.COLUMNS.STREET] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.STREET]) : "",
                    city: row[CONTACTS_SYNC_CFG.COLUMNS.CITY] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.CITY]) : "",
                    state: row[CONTACTS_SYNC_CFG.COLUMNS.STATE] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.STATE]) : "",
                    zip: row[CONTACTS_SYNC_CFG.COLUMNS.ZIP] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.ZIP]) : "",
                    groupsStr: row[CONTACTS_SYNC_CFG.COLUMNS.GROUPS] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.GROUPS]) : "",
                    notes: row[CONTACTS_SYNC_CFG.COLUMNS.NOTES] !== "" ? String(row[CONTACTS_SYNC_CFG.COLUMNS.NOTES]) : ""
                };

                if (contactData.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contactData.email)) {
                    throw new Error("⚠️ Invalid email format");
                }

                var person = { names: [], emailAddresses: [], phoneNumbers: [], organizations: [], biographies: [], addresses: [] };

                if (contactData.firstName || contactData.lastName) {
                    person.names.push({ givenName: contactData.firstName || "", familyName: contactData.lastName || "" });
                } else {
                    throw new Error("⚠️ Name is required to push.");
                }

                if (contactData.email) person.emailAddresses.push({ value: contactData.email });
                if (contactData.phone) person.phoneNumbers.push({ value: contactData.phone });
                if (contactData.company || contactData.title) person.organizations.push({ name: contactData.company || "", title: contactData.title || "" });
                if (contactData.notes) person.biographies.push({ value: contactData.notes });
                if (contactData.street || contactData.city || contactData.state || contactData.zip) {
                    person.addresses.push({
                        streetAddress: contactData.street || "",
                        city: contactData.city || "",
                        region: contactData.state || "",
                        postalCode: contactData.zip || ""
                    });
                }

                switch (action) {
                    case "CREATE":
                        var createdPerson = People.People.createContact(person);
                        rowUpdates.contactId = createdPerson.resourceName;
                        if (contactData.groupsStr) _ContactsSync_applyGroups(createdPerson.resourceName, contactData.groupsStr, groupNameToId);
                        if (contactData.starred === "Yes") {
                            People.ContactGroups.Members.modify({ resourceNamesToAdd: [createdPerson.resourceName] }, 'contactGroups/starred');
                        }
                        rowUpdates.status = "✅ Created";
                        rowUpdates.action = "";
                        break;

                    case "UPDATE":
                        if (!rowUpdates.contactId) throw new Error("⚠️ Missing Contact ID");
                        var existingPerson = People.People.get(rowUpdates.contactId, { personFields: 'names,emailAddresses,phoneNumbers,organizations,biographies,addresses' });
                        person.etag = existingPerson.etag;

                        if (existingPerson.emailAddresses && existingPerson.emailAddresses.length > 0) {
                            var primaryMailIndex = existingPerson.emailAddresses.findIndex(function (e) { return e.metadata && e.metadata.primary; });
                            if (primaryMailIndex === -1) primaryMailIndex = 0;
                            var existingMails = existingPerson.emailAddresses;
                            if (contactData.email) existingMails[primaryMailIndex].value = contactData.email;
                            person.emailAddresses = existingMails;
                        }

                        if (existingPerson.phoneNumbers && existingPerson.phoneNumbers.length > 0) {
                            var primaryPhoneIndex = existingPerson.phoneNumbers.findIndex(function (p) { return p.metadata && p.metadata.primary; });
                            if (primaryPhoneIndex === -1) primaryPhoneIndex = 0;
                            var existingPhones = existingPerson.phoneNumbers;
                            if (contactData.phone) existingPhones[primaryPhoneIndex].value = contactData.phone;
                            person.phoneNumbers = existingPhones;
                        }

                        if (existingPerson.addresses && existingPerson.addresses.length > 0) {
                            var primaryAddressIndex = existingPerson.addresses.findIndex(function (a) { return a.metadata && a.metadata.primary; });
                            if (primaryAddressIndex === -1) primaryAddressIndex = 0;
                            var existingAddresses = existingPerson.addresses;
                            if (contactData.street || contactData.city || contactData.state || contactData.zip) {
                                var newAddr = { streetAddress: contactData.street || "", city: contactData.city || "", region: contactData.state || "", postalCode: contactData.zip || "" };
                                if (primaryAddressIndex > -1) { newAddr.metadata = existingAddresses[primaryAddressIndex].metadata; existingAddresses[primaryAddressIndex] = newAddr; }
                                else { newAddr.metadata = { primary: true }; existingAddresses.push(newAddr); }
                            }
                            person.addresses = existingAddresses;
                        }

                        People.People.updateContact(person, rowUpdates.contactId, { updatePersonFields: 'names,emailAddresses,phoneNumbers,organizations,biographies,addresses' });
                        if (contactData.groupsStr) _ContactsSync_applyGroups(rowUpdates.contactId, contactData.groupsStr, groupNameToId);
                        if (contactData.starred === "Yes") {
                            try { People.ContactGroups.Members.modify({ resourceNamesToAdd: [rowUpdates.contactId] }, 'contactGroups/starred'); } catch (e) { }
                        }
                        rowUpdates.status = "✅ Updated";
                        rowUpdates.action = "";
                        break;

                    case "REMOVE":
                        if (!rowUpdates.contactId) throw new Error("⚠️ Missing Contact ID");
                        try { People.People.deleteContact(rowUpdates.contactId); } catch (e) { }
                        rowUpdates.status = "🗑️ Removed";
                        rowUpdates.action = "";
                        break;

                    default:
                        rowUpdates.status = "❓ Unknown Action '" + action + "'";
                }
                
                Logger.info(SyncEngine.getTool('CONTACTS_SYNC').TITLE, 'Row ' + (originalIdx + 2), rowUpdates.status);
                return rowUpdates;

            } catch (e) {
                rowUpdates.status = "⚠️ " + e.message;
                Logger.error(SyncEngine.getTool('CONTACTS_SYNC').TITLE, 'Row ' + (originalIdx + 2), e);
                return rowUpdates;
            }
        }, {
            onBatchComplete: function (batchResults) {
                var rowNumbers = [];
                var updatesArr = [];
                batchResults.forEach(function (res) {
                    if (res && res.originalIndex !== undefined) {
                        rowNumbers.push(res.originalIndex + 2);
                        updatesArr.push({ 'Action': res.action, 'Contact ID': res.contactId });
                    }
                });
                if (rowNumbers.length > 0) {
                    SheetManager.batchPatchRows('CONTACTS_SYNC', rowNumbers, updatesArr);
                }
            }
        });

        return _App_ok("Sync Complete. Processed: " + stats.processedCount);
    });
}

function _ContactsSync_applyGroups(resourceName, groupsStr, groupNameToId) {
    if (!groupsStr) return;
    var requestedGroups = groupsStr.split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });

    requestedGroups.forEach(function (gName) {
        var id = groupNameToId[gName];

        // Auto-create dynamically if group doesn't exist
        if (!id) {
            try {
                var newGroup = _App_callWithBackoff(function () {
                    return People.ContactGroups.create({
                        contactGroup: { name: gName }
                    });
                });
                id = newGroup.resourceName;
                groupNameToId[gName] = id; // Cache it for the rest of the run
            } catch (e) {
                console.error("Failed to auto-create group: " + gName);
            }
        }

        if (id) {
            try {
                _App_callWithBackoff(function () {
                    People.ContactGroups.Members.modify({ resourceNamesToAdd: [resourceName] }, id);
                });
            } catch (e) { }
        }
    });
}

// Stage 1: Skeleton — headers, column widths, freeze, data validations only
function _ContactsSync_setupSheetStructure(sheet) {
    var headers = SyncEngine.getTool('CONTACTS_SYNC').HEADERS;

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers])
        .setFontWeight(SHEET_THEME.LAYOUT.HEADER_WEIGHT)
        .setFontSize(SHEET_THEME.SIZES.HEADER)
        .setFontFamily(SHEET_THEME.FONTS.PRIMARY)
        .setBackground(SHEET_THEME.HEADER)
        .setFontColor(SHEET_THEME.TEXT)
        .setBorder(true, true, true, true, true, true, SHEET_THEME.BORDER, SHEET_THEME.BORDER_STYLE)
        .setVerticalAlignment(SHEET_THEME.LAYOUT.HEADER_ALIGN_V)
        .setHorizontalAlignment(SHEET_THEME.LAYOUT.HEADER_ALIGN_H);

    sheet.setFrozenRows(1);
    
    // widths typically managed by APP_REGISTRY, here are fallbacks if needed
    _ContactsSync_applyDataValidationsInternal(sheet);
}

function _ContactsSync_applyDataValidationsInternal(sheet) {
    var maxRows = sheet.getMaxRows();
    if (maxRows < 2) return;

    var ruleAction = SpreadsheetApp.newDataValidation().requireValueInList(["CREATE", "UPDATE", "REMOVE"], true).build();
    sheet.getRange(2, CONTACTS_SYNC_CFG.COLUMNS.ACTION + 1, maxRows - 1).setDataValidation(ruleAction);

    var ruleStarred = SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build();
    sheet.getRange(2, CONTACTS_SYNC_CFG.COLUMNS.STARRED + 1, maxRows - 1).setDataValidation(ruleStarred);
}

function ContactsSync_modifyGroupInActiveRow(groupName, action) {
    return Logger.run('CONTACTS_SYNC', 'Modify Group Row', function () {
        var validation = _App_validateActiveSheet(SHEET_NAMES.CONTACTS_SYNC);
        if (!validation.valid) return { success: false, message: validation.message };
        var sheet = validation.sheet;

        var cell = sheet.getActiveCell();
        var row = cell.getRow();

        // Ensure user is on a valid data row
        if (row < 2) return { success: false, message: "Please select a contact row." };

        var groupCell = sheet.getRange(row, CONTACTS_SYNC_CFG.COLUMNS.GROUPS + 1);
        var currentVal = groupCell.getValue().toString().trim();
        var existingGroups = currentVal ? currentVal.split(',').map(function (s) { return s.trim(); }) : [];

        var changed = false;

        if (action === 'add') {
            if (!existingGroups.includes(groupName)) {
                existingGroups.push(groupName);
                changed = true;
            } else {
                return { success: true, message: "Group '" + groupName + "' is already on this row." };
            }
        } else if (action === 'remove') {
            var index = existingGroups.indexOf(groupName);
            if (index > -1) {
                existingGroups.splice(index, 1);
                changed = true;
            } else {
                return { success: true, message: "Group '" + groupName + "' is not on this row." };
            }
        }

        if (changed) {
            groupCell.setValue(existingGroups.join(', '));
            // Auto-set Action to UPDATE if not CREATE
            var actionRange = sheet.getRange(row, CONTACTS_SYNC_CFG.COLUMNS.ACTION + 1);
            var currentAction = actionRange.getValue().toString().trim().toUpperCase();
            if (currentAction !== "CREATE") {
                actionRange.setValue("UPDATE");
            }
            return { success: true, message: (action === 'add' ? "Added" : "Removed") + " '" + groupName + "'." };
        }
        return { success: true, message: "No changes made." };
    });
}

function ContactsSync_createContactGroup(groupName) {
    return Logger.run('CONTACTS_SYNC', 'Create Group', function () {
        if (typeof People === 'undefined') throw new Error("People API not enabled");
        var newGroup = _App_callWithBackoff(function () {
            return People.ContactGroups.create({
                contactGroup: { name: groupName }
            });
        });
        return _App_ok('Contact group created.', {
            id: newGroup.resourceName,
            name: newGroup.formattedName || newGroup.name
        });
    });
}

function ContactsSync_deleteContactGroup(resourceName) {
    return Logger.run('CONTACTS_SYNC', 'Delete Group', function () {
        if (typeof People === 'undefined') throw new Error("People API not enabled");
        _App_callWithBackoff(function () {
            People.ContactGroups.remove(resourceName, { deleteContacts: false });
        });
        return true;
    });
}

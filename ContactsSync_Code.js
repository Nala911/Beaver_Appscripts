/**
 * Google Contacts
 * Version: 6.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CONTACTS_SYNC', {
    REQUIRED_SERVICES: [{ name: 'People API', test: function () { return typeof People !== 'undefined'; } }],
    SHEET_NAME: SHEET_NAMES.CONTACTS_SYNC,
    TITLE: SHEET_NAMES.CONTACTS_SYNC,
    MENU_LABEL: SHEET_NAMES.CONTACTS_SYNC,
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
            { header: 'Status', type: 'STATUS' },
            { header: 'First Name', type: 'TEXT' },
            { header: 'Last Name', type: 'TEXT' },
            { header: 'Email', type: 'TEXT' },
            { header: 'Phone', type: 'TEXT' },
            { header: 'Company', type: 'TEXT' },
            { header: 'Job Title', type: 'TEXT' },
            { header: 'Starred', type: 'CHECKBOX' },
            { header: 'Street', type: 'TEXT' },
            { header: 'City', type: 'TEXT' },
            { header: 'State', type: 'TEXT' },
            { header: 'Zip', type: 'TEXT' },
            {
                header: 'Groups/Labels', type: 'DROPDOWN', allowInvalid: true, options: function () {
                    var groups = [];
                    try {
                        var response = _App_callWithBackoff(function () {
                            return People.ContactGroups.list({ pageSize: 1000 });
                        });
                        var excluded = ['Friends', 'Family', 'Coworkers', 'All Contacts', 'Starred'];
                        (response.contactGroups || []).forEach(function (g) {
                            var name = g.formattedName || g.name;
                            if (name && !excluded.includes(name)) {
                                groups.push(name);
                            }
                        });
                        groups.sort();
                    } catch (e) { }
                    return groups.length ? groups.slice(0, 499) : ['None'];
                }
            },
            { header: 'Notes', type: 'TEXT' },
            { header: 'Contact ID', type: 'ID', italic: true }
        ]
    }
});



// Declarative format config now lives in SyncEngine.getTool('CONTACTS_SYNC').FORMAT_CONFIG

/** Opens the Contacts sidebar and ensures the sheet exists. */
function ContactsSync_openSidebar() {
    return Logger.run('CONTACTS_SYNC', 'Open Sidebar', function () {
        _App_launchTool('CONTACTS_SYNC');
    });
}


function ContactsSync_savePreferences(groupIds) {
    return Logger.run('CONTACTS_SYNC', 'Save Preferences', function () {
        if (groupIds) _App_setProperty(APP_PROPS.CONTACTS_SELECTED_GROUPS, groupIds);
        return _App_ok('Preferences saved.');
    });
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

                    var isStarred = pGroups.includes('contactGroups/starred');

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
                        "", // Status
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
    return Logger.run('CONTACTS_SYNC', 'Check Unsaved', function () {
        return _App_ok('Check complete.', SheetManager.hasPendingActions('CONTACTS_SYNC'));
    });
}

function ContactsSync_pushChanges() {
    return Logger.run('CONTACTS_SYNC', 'Push Changes', function () {
        if (typeof People === 'undefined') {
            throw new Error("⚠️ People API is not enabled. Go to Services -> Add 'People API'.");
        }

        var groupsResponse = People.ContactGroups.list();
        var allGroups = groupsResponse.contactGroups || [];
        var groupNameToId = {};
        allGroups.forEach(function (g) {
            groupNameToId[g.formattedName || g.name] = g.resourceName;
        });

        var pendingRows = SheetManager.readPendingObjects('CONTACTS_SYNC');

        if (pendingRows.length === 0) return _App_ok("No pending actions found.");

        var stats = _App_BatchProcessor('CONTACTS_SYNC', pendingRows, function (item) {
            var rowUpdates = {
                action: item['Action'],
                contactId: item['Contact ID'],
                status: "",
                _rowNumber: item._rowNumber
            };

            var action = rowUpdates.action.toString().toUpperCase();
            var contactData = {
                    firstName: item['First Name'] !== "" ? String(item['First Name']) : "",
                    lastName: item['Last Name'] !== "" ? String(item['Last Name']) : "",
                    email: item['Email'] !== "" ? String(item['Email']) : "",
                    phone: item['Phone'] !== "" ? String(item['Phone']) : "",
                    company: item['Company'] !== "" ? String(item['Company']) : "",
                    title: item['Job Title'] !== "" ? String(item['Job Title']) : "",
                    starred: item['Starred'],
                    street: item['Street'] !== "" ? String(item['Street']) : "",
                    city: item['City'] !== "" ? String(item['City']) : "",
                    state: item['State'] !== "" ? String(item['State']) : "",
                    zip: item['Zip'] !== "" ? String(item['Zip']) : "",
                    groupsStr: item['Groups/Labels'] !== "" ? String(item['Groups/Labels']) : "",
                    notes: item['Notes'] !== "" ? String(item['Notes']) : ""
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
                        if (contactData.starred === true || contactData.starred === 'TRUE') {
                            People.ContactGroups.Members.modify({ resourceNamesToAdd: [createdPerson.resourceName] }, 'contactGroups/starred');
                        }
                        rowUpdates.status = SHEET_THEME.STATUS_PREFIXES.SUCCESS + "Created";
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
                        if (contactData.starred === true || contactData.starred === 'TRUE') {
                            try { People.ContactGroups.Members.modify({ resourceNamesToAdd: [rowUpdates.contactId] }, 'contactGroups/starred'); } catch (e) { }
                        }
                        rowUpdates.status = SHEET_THEME.STATUS_PREFIXES.SUCCESS + "Updated";
                        rowUpdates.action = "";
                        break;

                    case "REMOVE":
                        if (!rowUpdates.contactId) throw new Error("⚠️ Missing Contact ID");
                        try { People.People.deleteContact(rowUpdates.contactId); } catch (e) { }
                        rowUpdates.status = SHEET_THEME.STATUS_PREFIXES.SUCCESS + "Removed";
                        rowUpdates.action = "";
                        break;

                    default:
                        rowUpdates.status = "❓ Unknown Action '" + action + "'";
                }

                return rowUpdates;

        }, {
            onBatchComplete: function (batchResults) {
                var rowNumbers = [];
                var updatesArr = [];
                var prefixes = SHEET_THEME.STATUS_PREFIXES;

                batchResults.forEach(function (res) {
                    if (res && res._rowNumber !== undefined) {
                        rowNumbers.push(res._rowNumber);
                        if (res.isError) {
                            updatesArr.push({ 'Status': prefixes.ERROR + res.error });
                        } else {
                            updatesArr.push({ 'Action': res.action, 'Status': res.status, 'Contact ID': res.contactId });
                        }
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

/**
 * Checks for groups in the pending rows that don't exist in Google Contacts.
 * Used for pre-push confirmation in the sidebar.
 */
function ContactsSync_getMissingGroups() {
    return Logger.run('CONTACTS_SYNC', 'Check Missing Groups', function () {
        var pendingItems = SheetManager.readPendingObjects('CONTACTS_SYNC');
        if (pendingItems.length === 0) return _App_ok('No pending actions.', []);

        var groupsInSheet = [];
        pendingItems.forEach(function (item) {
            var action = (item['Action'] || '').toString().toUpperCase();
            if (action === 'CREATE' || action === 'UPDATE') {
                var groupsStr = item['Groups/Labels'] ? String(item['Groups/Labels']) : '';
                if (groupsStr) {
                    var split = groupsStr.split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
                    split.forEach(function (g) {
                        if (groupsInSheet.indexOf(g) === -1) groupsInSheet.push(g);
                    });
                }
            }
        });

        if (groupsInSheet.length === 0) return _App_ok('No groups to check.', []);

        var groupsResponse = People.ContactGroups.list();
        var allGroups = groupsResponse.contactGroups || [];
        var existingNames = allGroups.map(function (g) { return g.formattedName || g.name; });

        var missing = groupsInSheet.filter(function (name) {
            return existingNames.indexOf(name) === -1;
        });

        return _App_ok('Missing groups identified.', missing);
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
                // Silently skip if auto-creation fails
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
    var tool = SyncEngine.getTool('CONTACTS_SYNC');
    _App_applyBodyFormatting(sheet, 0, tool.FORMAT_CONFIG);
}

function _ContactsSync_applyDataValidationsInternal(sheet) {
    var maxRows = sheet.getMaxRows();
    if (maxRows < 2) return;
    var headers = SheetManager.getHeaders('CONTACTS_SYNC');
    var actionColIndex = headers.indexOf('Action') + 1;
    var ruleAction = SpreadsheetApp.newDataValidation().requireValueInList(["CREATE", "UPDATE", "REMOVE"], true).build();
    if (actionColIndex > 0) sheet.getRange(2, actionColIndex, maxRows - 1).setDataValidation(ruleAction);
}


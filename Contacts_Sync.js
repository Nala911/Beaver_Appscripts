/**
 * MASTER EDITION: Google Sheets <-> Google Contacts Sync
 * Version: 5.0 (Plugin Architecture — registers with App.Engine)
 */

App.Engine.registerTool('CONTACTS_SYNC', {
    REQUIRED_SERVICES: [ { name: 'People API', test: function() { return typeof People !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CONTACTS_SYNC,
    TITLE: '📇 Contacts Sync Master',
    MENU_LABEL: '☎️ Google Contacts',
    MENU_ENTRYPOINT: 'Contacts_showSidebar',
    MENU_ORDER: 20,
    SIDEBAR_HTML: 'Contacts_Sidebar',
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
    },

    /**
     * CONTACTS SYNC SERVICE ACTIONS
     */
    service: {
        getLoadData: function() {
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

                var prefs = App.Engine.getPrefs('CONTACTS_SYNC');

                return _App_ok('Contacts load data ready.', {
                    groups: formattedGroups,
                    savedGroupIds: prefs.selectedGroupIds || ['all']
                });
            } catch (err) {
                throw new Error('Unable to load contact groups. ' + err.message);
            }
        },

        savePreferences: function(groupIds) {
            var prefs = App.Engine.getPrefs('CONTACTS_SYNC');
            if (groupIds) prefs.selectedGroupIds = groupIds;
            App.Engine.setPrefs('CONTACTS_SYNC', prefs);
            return _App_ok('Preferences saved.');
        },

        pullContacts: function(request) {
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
                        var firstName = "", lastName = "";
                        if (person.names && person.names.length > 0) {
                            var primaryName = person.names.find(function (n) { return n.metadata && n.metadata.primary; }) || person.names[0];
                            firstName = primaryName.givenName || "";
                            lastName = primaryName.familyName || "";
                        }

                        var email = _ContactsSync_getPrimary(person.emailAddresses);
                        var phone = _ContactsSync_getPrimary(person.phoneNumbers);

                        var company = "", title = "";
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
            var formatConfig = JSON.parse(JSON.stringify(App.Engine.getTool('CONTACTS_SYNC').FORMAT_CONFIG));
            formatConfig.conditionalRules = formatConfig.conditionalRules.concat([
                { type: 'custom', formula: '=AND($F2<>"", COUNTIF($F:$F, $F2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 6 },
                { type: 'custom', formula: '=AND($G2<>"", COUNTIF($G:$G, $G2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 7 }
            ]);
            SheetManager.overwriteRows('CONTACTS_SYNC', outputData, {
                totalCols: App.Engine.getTool('CONTACTS_SYNC').HEADERS.length,
                formatConfig: formatConfig
            });

            this.savePreferences(groupIds);
            return _App_ok('Successfully imported ' + outputData.length + " contacts.");
        },

        checkForUnsavedChanges: function() {
            return SheetManager.hasPendingActions('CONTACTS_SYNC');
        },

        pushChanges: function() {
            if (typeof People === 'undefined') {
                throw new Error("⚠️ People API is not enabled. Go to Services -> Add 'People API'.");
            }

            var groupsResponse = People.ContactGroups.list();
            var allGroups = groupsResponse.contactGroups || [];
            var groupNameToId = {};
            allGroups.forEach(function (g) {
                groupNameToId[g.formattedName || g.name] = g.resourceName;
            });

            var stats = ExecutionService.processPendingRows('CONTACTS_SYNC', function(rowObj) {
                var action = String(rowObj['Action'] || '').toUpperCase();
                var contactId = rowObj['Contact ID'] ? String(rowObj['Contact ID']) : null;

                var updates = { 'Action': '', 'Log': '' };
                var contactData = {
                    firstName: String(rowObj['First Name'] || ""),
                    lastName: String(rowObj['Last Name'] || ""),
                    email: String(rowObj['Email'] || ""),
                    phone: String(rowObj['Phone'] || ""),
                    company: String(rowObj['Company'] || ""),
                    title: String(rowObj['Job Title'] || ""),
                    starred: rowObj['Starred'],
                    street: String(rowObj['Street'] || ""),
                    city: String(rowObj['City'] || ""),
                    state: String(rowObj['State'] || ""),
                    zip: String(rowObj['Zip'] || ""),
                    groupsStr: String(rowObj['Groups/Labels'] || ""),
                    notes: String(rowObj['Notes'] || "")
                };

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
                        updates['Contact ID'] = createdPerson.resourceName;
                        if (contactData.groupsStr) _ContactsSync_applyGroups(createdPerson.resourceName, contactData.groupsStr, groupNameToId);
                        if (contactData.starred === "Yes") {
                            People.ContactGroups.Members.modify({ resourceNamesToAdd: [createdPerson.resourceName] }, 'contactGroups/starred');
                        }
                        updates['Log'] = "✅ Created";
                        break;

                    case "UPDATE":
                        if (!contactId) throw new Error("⚠️ Missing Contact ID");
                        var existingPerson = People.People.get(contactId, {
                            personFields: 'names,emailAddresses,phoneNumbers,organizations,biographies,addresses'
                        });
                        person.etag = existingPerson.etag;

                        // MERGE LOGIC (Prevent Data Loss for Secondary Items)
                        if (existingPerson.emailAddresses && existingPerson.emailAddresses.length > 0) {
                            var primaryMailIndex = existingPerson.emailAddresses.findIndex(function (e) { return e.metadata && e.metadata.primary; });
                            if (primaryMailIndex === -1) primaryMailIndex = 0; 
                            var existingMails = existingPerson.emailAddresses;
                            if (contactData.email) {
                                if (primaryMailIndex > -1) existingMails[primaryMailIndex].value = contactData.email;
                                else existingMails.push({ value: contactData.email, metadata: { primary: true } });
                            }
                            person.emailAddresses = existingMails;
                        }

                        if (existingPerson.phoneNumbers && existingPerson.phoneNumbers.length > 0) {
                            var primaryPhoneIndex = existingPerson.phoneNumbers.findIndex(function (p) { return p.metadata && p.metadata.primary; });
                            if (primaryPhoneIndex === -1) primaryPhoneIndex = 0;
                            var existingPhones = existingPerson.phoneNumbers;
                            if (contactData.phone) {
                                if (primaryPhoneIndex > -1) existingPhones[primaryPhoneIndex].value = contactData.phone;
                                else existingPhones.push({ value: contactData.phone, metadata: { primary: true } });
                            }
                            person.phoneNumbers = existingPhones;
                        }

                        if (existingPerson.addresses && existingPerson.addresses.length > 0) {
                            var primaryAddressIndex = existingPerson.addresses.findIndex(function (a) { return a.metadata && a.metadata.primary; });
                            if (primaryAddressIndex === -1) primaryAddressIndex = 0;
                            var existingAddresses = existingPerson.addresses;
                            var hasNewAddressData = contactData.street || contactData.city || contactData.state || contactData.zip;
                            if (hasNewAddressData) {
                                var newAddr = { streetAddress: contactData.street || "", city: contactData.city || "", region: contactData.state || "", postalCode: contactData.zip || "" };
                                if (primaryAddressIndex > -1) { newAddr.metadata = existingAddresses[primaryAddressIndex].metadata; existingAddresses[primaryAddressIndex] = newAddr; }
                                else { newAddr.metadata = { primary: true }; existingAddresses.push(newAddr); }
                            }
                            person.addresses = existingAddresses;
                        }

                        People.People.updateContact(person, contactId, { updatePersonFields: 'names,emailAddresses,phoneNumbers,organizations,biographies,addresses' });
                        if (contactData.groupsStr) _ContactsSync_applyGroups(contactId, contactData.groupsStr, groupNameToId);
                        if (contactData.starred === "Yes") { try { People.ContactGroups.Members.modify({ resourceNamesToAdd: [contactId] }, 'contactGroups/starred'); } catch (e) { } }
                        else { try { People.ContactGroups.Members.modify({ resourceNamesToRemove: [contactId] }, 'contactGroups/starred'); } catch (e) { } }
                        updates['Log'] = "✅ Updated";
                        break;

                    case "REMOVE":
                        if (!contactId) throw new Error("⚠️ Missing Contact ID");
                        try { People.People.deleteContact(contactId); updates['Log'] = "🗑️ Removed"; }
                        catch (delErr) { updates['Log'] = "⚠️ Resource missing"; }
                        break;

                    default:
                        updates['Log'] = "❓ Unknown Action '" + action + "'";
                        updates['Action'] = rowObj['Action'];
                }
                SheetManager.patchRow('CONTACTS_SYNC', rowObj._rowNumber, updates);
            });

            if (stats.processed === 0 && stats.errors === 0) {
                return _App_ok("No data to sync.");
            }

            return _App_ok("Sync Complete. Success: " + stats.processed + ", Errors: " + stats.errors);
        },

        modifyGroupInActiveRow: function(params) {
            var groupName = params.groupName;
            var action = params.action;
            var validation = _App_validateActiveSheet(SHEET_NAMES.CONTACTS_SYNC);
            if (!validation.valid) throw new Error(validation.message);
            var sheet = validation.sheet;
            var cell = sheet.getActiveCell();
            var row = cell.getRow();
            if (row < 2) throw new Error("Please select a contact row.");

            var groupCell = sheet.getRange(row, App.Engine.getTool('CONTACTS_SYNC').HEADERS.indexOf('Groups/Labels') + 1);
            var currentVal = groupCell.getValue().toString().trim();
            var existingGroups = currentVal ? currentVal.split(',').map(function (s) { return s.trim(); }) : [];
            var changed = false;

            if (action === 'add') {
                if (!existingGroups.includes(groupName)) { existingGroups.push(groupName); changed = true; }
                else return _App_ok("Group '" + groupName + "' is already on this row.");
            } else if (action === 'remove') {
                var index = existingGroups.indexOf(groupName);
                if (index > -1) { existingGroups.splice(index, 1); changed = true; }
                else return _App_ok("Group '" + groupName + "' is not on this row.");
            }

            if (changed) {
                groupCell.setValue(existingGroups.join(', '));
                var actionRange = sheet.getRange(row, App.Engine.getTool('CONTACTS_SYNC').HEADERS.indexOf('Action') + 1);
                var currentAction = actionRange.getValue().toString().trim().toUpperCase();
                if (currentAction !== "CREATE") actionRange.setValue("UPDATE");
                return _App_ok((action === 'add' ? "Added" : "Removed") + " '" + groupName + "'.");
            }
            return _App_ok("No changes made.");
        },

        createContactGroup: function(groupName) {
            if (typeof People === 'undefined') throw new Error("People API not enabled");
            var newGroup = _App_callWithBackoff(function () {
                return People.ContactGroups.create({ contactGroup: { name: groupName } });
            });
            return _App_ok('Contact group created.', { id: newGroup.resourceName, name: newGroup.formattedName || newGroup.name });
        },

        deleteContactGroup: function(resourceName) {
            if (typeof People === 'undefined') throw new Error("People API not enabled");
            _App_callWithBackoff(function () {
                People.ContactGroups.remove(resourceName, { deleteContacts: false });
            });
            return _App_ok("Group deleted.");
        }
    }
});

// Column-index aliases — kept for backward compatibility within this file.
// Metadata (title, sidebar, headers, widths) now lives in App.Engine.getTool('CONTACTS_SYNC').

// Declarative format config now lives in App.Engine.getTool('CONTACTS_SYNC').FORMAT_CONFIG

/** Opens the Contacts sidebar and ensures the sheet exists. */
function Contacts_showSidebar() {
  return Logger.run('CONTACTS_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CONTACTS_SYNC');
  });
}


function _ContactsSync_getPrimary(array) {
    if (!array || array.length === 0) return "";
    var primary = array.find(function (item) { return item.metadata && item.metadata.primary; });
    return primary ? primary.value : array[0].value;
}



// Kept for backward compatibility — now delegates to shared utility
function _ContactsSync_highlightDuplicates(sheet) {
    var lastRow = sheet.getLastRow();
    var numDataRows = lastRow > 1 ? lastRow - 1 : 0;
    var formatConfig = JSON.parse(JSON.stringify(App.Engine.getTool('CONTACTS_SYNC').FORMAT_CONFIG));
    formatConfig.conditionalRules = formatConfig.conditionalRules.concat([
        { type: 'custom', formula: '=AND($F2<>"", COUNTIF($F:$F, $F2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 6 },
        { type: 'custom', formula: '=AND($G2<>"", COUNTIF($G:$G, $G2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 7 }
    ]);
    _App_applyBodyFormatting(sheet, numDataRows, formatConfig);
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

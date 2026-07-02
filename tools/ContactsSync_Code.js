/**
 * Google Contacts
 * Version: 6.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CONTACTS_SYNC', {
    REQUIRED_SERVICES: [{ name: 'People API', test: function () { return typeof People !== 'undefined'; } }],
    SHEET_NAME: SHEET_NAMES.CONTACTS_SYNC,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.CONTACTS_SYNC,
    MENU_LABEL: SHEET_NAMES.CONTACTS_SYNC,
    MENU_ENTRYPOINT: 'ContactsSync_openSidebar',
    MENU_ORDER: 20,
    SIDEBAR_HTML: 'tools/ContactsSync_Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [
            { type: 'pending', actionCol: 'A', scope: 'actionOnly' },
            { type: 'custom', formula: '=AND($E2<>\'\', COUNTIF($E:$E, $E2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 5 },
            { type: 'custom', formula: '=AND($F2<>\'\', COUNTIF($F:$F, $F2)>1)', color: SHEET_THEME.STATUS.WARNING, scope: 'custom_col', col: 6 }
        ],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'DELETE'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'First Name', type: 'TEXT' },
            { header: 'Last Name', type: 'TEXT' },
            { header: 'Email', type: 'EMAIL' },
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
    },
    ACTIONS: {
        getMissingGroups: function () {
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
        },
        savePreferences: function (groupIds) {
            if (groupIds) _App_setProperty(APP_PROPS.CONTACTS_SELECTED_GROUPS, groupIds);
            return _App_ok('Preferences saved.');
        },
        pull: function (request) {
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

            // Apply body formatting using the registered tool config directly
            SheetManager.overwriteRows('CONTACTS_SYNC', outputData, {
                totalCols: SyncEngine.getTool('CONTACTS_SYNC').HEADERS.length,
                formatConfig: SyncEngine.getTool('CONTACTS_SYNC').FORMAT_CONFIG
            });

            SyncEngine.getTool('CONTACTS_SYNC').ACTIONS.savePreferences(groupIds);
            return _App_ok('Successfully imported ' + outputData.length + " contacts.");
        },
        push: function () {
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

            var groupAdditions = {};

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
                        if (contactData.groupsStr) {
                            contactData.groupsStr.split(',').map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (gName) {
                                var gId = _ContactsSync_getOrCreateGroup(gName, groupNameToId);
                                if (gId) {
                                    if (!groupAdditions[gId]) groupAdditions[gId] = [];
                                    groupAdditions[gId].push(createdPerson.resourceName);
                                }
                            });
                        }
                        if (contactData.starred === true || contactData.starred === 'TRUE') {
                            if (!groupAdditions['contactGroups/starred']) groupAdditions['contactGroups/starred'] = [];
                            groupAdditions['contactGroups/starred'].push(createdPerson.resourceName);
                        }
                        rowUpdates.status = _App_formatStatus('SUCCESS', "Created");
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
                        if (contactData.groupsStr) {
                            contactData.groupsStr.split(',').map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (gName) {
                                var gId = _ContactsSync_getOrCreateGroup(gName, groupNameToId);
                                if (gId) {
                                    if (!groupAdditions[gId]) groupAdditions[gId] = [];
                                    groupAdditions[gId].push(rowUpdates.contactId);
                                }
                            });
                        }
                        if (contactData.starred === true || contactData.starred === 'TRUE') {
                            if (!groupAdditions['contactGroups/starred']) groupAdditions['contactGroups/starred'] = [];
                            groupAdditions['contactGroups/starred'].push(rowUpdates.contactId);
                        }
                        rowUpdates.status = _App_formatStatus('SUCCESS', "Updated");
                        rowUpdates.action = "";
                        break;

                    case "DELETE":
                        if (!rowUpdates.contactId) throw new Error("⚠️ Missing Contact ID");
                        try { People.People.deleteContact(rowUpdates.contactId); } catch (e) { }
                        rowUpdates.status = _App_formatStatus('SUCCESS', "Deleted");
                        rowUpdates.action = "";
                        break;

                    default:
                        rowUpdates.status = _App_formatStatus('WARNING', "Unknown Action '" + action + "'");
                }

                return rowUpdates;

            }, {
                onBatchComplete: function (batchResults) {
                    Object.keys(groupAdditions).forEach(function (gId) {
                        var members = groupAdditions[gId];
                        if (members && members.length > 0) {
                            try {
                                _App_callWithBackoff(function () {
                                    People.ContactGroups.Members.modify({ resourceNamesToAdd: members }, gId);
                                });
                            } catch (e) {
                                Logger.warn('CONTACTS_SYNC', 'Batch Group Modify', 'Failed to add members to ' + gId + ': ' + e.message);
                            }
                        }
                    });
                    groupAdditions = {};

                    _App_batchPatchResults('CONTACTS_SYNC', batchResults, function (res) {
                        return {
                            'Contact ID': res.contactId
                        };
                    });
                }
            });

            return _App_ok("Sync Complete. Processed: " + stats.processedCount);
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Contacts sidebar and ensures the sheet exists. */
function ContactsSync_openSidebar() {
    return Logger.run('CONTACTS_SYNC', 'Open Sidebar', function () {
        _App_launchTool('CONTACTS_SYNC');
    });
}


function _ContactsSync_getOrCreateGroup(gName, groupNameToId) {
    var id = groupNameToId[gName];
    if (!id) {
        try {
            var newGroup = _App_callWithBackoff(function () {
                return People.ContactGroups.create({
                    contactGroup: { name: gName }
                });
            });
            id = newGroup.resourceName;
            groupNameToId[gName] = id;
        } catch (e) {
            // Silently return null if creation fails
        }
    }
    return id;
}

/**
 * Global wrapper to allow Contacts pull to be run from triggers or manually from the IDE.
 * @param {Object} [request] Optional request parameters.
 */
function ContactsSync_pullContacts(request) {
    if (!request || typeof request !== 'object' || !request.groupIds) {
        var savedGroups = _App_getProperty(APP_PROPS.CONTACTS_SELECTED_GROUPS);
        if (!Array.isArray(savedGroups)) {
            savedGroups = savedGroups ? [savedGroups] : ['all'];
        }
        request = {
            groupIds: savedGroups
        };
    }
    return SyncEngine.runAction('CONTACTS_SYNC', 'pull', [request]);
}





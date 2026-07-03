/**
 * Google Chat Space Sync Tool
 * Version: 1.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CHAT_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Chat API', test: function() { return typeof Chat !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CHAT_SYNC,
    FROZEN_COLS: 2,
    TITLE: SHEET_NAMES.CHAT_SYNC,
    MENU_LABEL: SHEET_NAMES.CHAT_SYNC,
    MENU_ENTRYPOINT: 'ChatSpaceSync_openSidebar',
    MENU_ORDER: 15,
    SIDEBAR_HTML: 'tools/ChatSpaceSync/Sidebar',
    SIDEBAR_WIDTH: 400,
    FORMAT_CONFIG: {
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['ADD_MEMBER', 'REMOVE_MEMBER'] },
            { header: 'Status', type: 'STATUS' },
            { header: 'Space Name', type: 'TEXT' },
            { header: 'Member Email', type: 'TEXT' },
            { header: 'Role', type: 'DROPDOWN', options: ['ROLE_MEMBER', 'ROLE_MANAGER'] },
            { header: 'Type', type: 'TEXT' }, // User or Group
            { header: 'Space ID', type: 'ID' },
            { header: 'Membership ID', type: 'ID' }
        ]
    },
    HELP_ITEMS: {
        gettingStarted: {
            title: "Getting Started",
            content: "<p><strong>Getting Started</strong></p><p>Follow these steps to sync Chat Space members:</p><ol><li><strong>Select Spaces:</strong> Check target spaces in the sidebar.</li><li><strong>Pull:</strong> Click <strong>Pull Members</strong> to import current members.</li><li><strong>Modify:</strong> Set action to <code>ADD_MEMBER</code> or <code>REMOVE_MEMBER</code>.</li><li><strong>Push:</strong> Click <strong>Push Changes</strong>.</li></ol>"
        },
        items: [
            {
                icon: "help-circle",
                color: "var(--primary)",
                label: "Column Guide",
                shortDesc: "Action, Role, and Member Email.",
                tooltipId: "help-columns-guide",
                tooltipContent: "<p><strong>Core Columns Guide</strong></p><ul><li><strong>Action:</strong> Set to <code>ADD_MEMBER</code> to invite a user or <code>REMOVE_MEMBER</code> to evict.</li><li><strong>Role:</strong> Choose <code>ROLE_MEMBER</code> or <code>ROLE_MANAGER</code>.</li><li><strong>Type:</strong> Read-only user category (User, Group, Bot).</li><li><strong>IDs:</strong> System-generated IDs. Do not manually edit.</li></ul>"
            }
        ]
    },
    ACTIONS: {
        getLoadData: function () {
            try {
                var spacesList = [];
                var pageToken = null;
                do {
                    var response = _App_callWithBackoff(function() {
                        return Chat.Spaces.list({ pageToken: pageToken });
                    });
                    if (response.spaces) {
                        spacesList = spacesList.concat(response.spaces);
                    }
                    pageToken = response.nextPageToken;
                } while (pageToken);

                var uniqueSpaces = spacesList.map(function (s) {
                    return {
                        id: s.name,
                        name: s.displayName || s.name
                    };
                });

                var savedSpaceIds = _App_getProperty(APP_PROPS.CHAT_SELECTED_SPACES);
                if (!Array.isArray(savedSpaceIds)) savedSpaceIds = [];

                return _App_ok('Chat spaces load data ready.', {
                    spaces: uniqueSpaces,
                    savedSpaceIds: savedSpaceIds
                });
            } catch (err) {
                throw new Error('Unable to load chat spaces: ' + err.message);
            }
        },
        savePreferences: function (spaceIds) {
            if (spaceIds) _App_setProperty(APP_PROPS.CHAT_SELECTED_SPACES, spaceIds);
            return _App_ok('Preferences saved.');
        },
        pull: function () {
            return _ChatSpaceSync_pullMembers();
        },
        push: function () {
            return _ChatSpaceSync_pushChanges();
        }
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Chat Sync sidebar and ensures the sheet exists. */
function ChatSpaceSync_openSidebar() {
  return Logger.run('CHAT_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CHAT_SYNC');
  });
}

// --- PREFERENCES & STATE ---

function _ChatSpaceSync_pullMembers() {
  return Logger.run('CHAT_SYNC', 'Pull Members', function () {
    return _App_withDocumentLock('CHAT_SYNC_PULL', function () {
      var TARGET_SHEET_NAME = SHEET_NAMES.CHAT_SYNC;
      var sheet = _App_ensureSheetExists('CHAT_SYNC');

    var outputObjects = [];
    var spacesList = [];
    var pageToken = null;

    var savedSpaceIds = _App_getProperty(APP_PROPS.CHAT_SELECTED_SPACES);
    if (!Array.isArray(savedSpaceIds)) savedSpaceIds = [];

    if (savedSpaceIds.length > 0) {
      // Pull only selected spaces
      savedSpaceIds.forEach(function (spaceId) {
        try {
          var space = _App_callWithBackoff(function () {
            return Chat.Spaces.get(spaceId);
          });
          if (space) spacesList.push(space);
        } catch (err) {
          Logger.warn('CHAT_SYNC', 'Fetch Space Error', 'Space ' + spaceId + ': ' + err.message);
        }
      });
    } else {
      // Fetch all spaces the user is a member of
      do {
        var response = _App_callWithBackoff(function() {
            return Chat.Spaces.list({
              pageToken: pageToken
            });
        });
        
        if (response.spaces) {
          spacesList = spacesList.concat(response.spaces);
        }
        pageToken = response.nextPageToken;
      } while (pageToken);
    }

    spacesList.forEach(function (space) {
      try {
        var spaceNameId = space.name;
        var spaceDisplayName = space.displayName || space.name;
        var memberPageToken = null;
        var members = [];

        do {
            var memberResponse = _App_callWithBackoff(function() {
                return Chat.Spaces.Members.list(spaceNameId, {
                    pageToken: memberPageToken
                });
            });
            if (memberResponse.memberships) {
                members = members.concat(memberResponse.memberships);
            }
            memberPageToken = memberResponse.nextPageToken;
        } while (memberPageToken);

        members.forEach(function (m) {
          var memberEmail = "";
          var memberType = "Unknown";
          
          if (m.member && m.member.type === "HUMAN") {
              memberEmail = m.member.displayName || m.member.name;
              memberType = "User";
          } else if (m.groupMember) {
              memberEmail = m.groupMember.id;
              memberType = "Group";
          } else if (m.member && m.member.type === "BOT") {
              memberEmail = m.member.displayName || "Bot";
              memberType = "Bot";
          }

          outputObjects.push({
            'Action': "",
            'Status': "",
            'Space Name': spaceDisplayName,
            'Member Email': memberEmail,
            'Role': m.role === 'ROLE_MANAGER' ? 'ROLE_MANAGER' : 'ROLE_MEMBER',
            'Type': memberType,
            'Space ID': spaceNameId,
            'Membership ID': m.name
          });
        });
      } catch (err) {
        throw new Error('Pull Members failed for ' + space.name + ': ' + err.message);
      }
    });

    // Sort by Space Name alphabetically
    outputObjects.sort(function(a, b) {
        return a['Space Name'].localeCompare(b['Space Name']);
    });

      SheetManager.overwriteObjects('CHAT_SYNC', outputObjects);
      
      var summary = 'Successfully imported ' + outputObjects.length + " members into '" + TARGET_SHEET_NAME + "'.";
      return _App_ok(summary);
    });
  });
}

// --- THE "PUSH" WORKFLOW ---

function _ChatSpaceSync_pushChanges() {
  return Logger.run('CHAT_SYNC', 'Push Changes', function () {
    return _App_withDocumentLock('CHAT_SYNC_PUSH', function () {
      var pendingItems = SheetManager.readPendingObjects('CHAT_SYNC');

      if (pendingItems.length === 0) return _App_ok("No pending actions found.");

    var stats = _App_BatchProcessor('CHAT_SYNC', pendingItems, function (item) {
      var rowUpdates = {
        action: item['Action'],
        membershipId: item['Membership ID'] ? String(item['Membership ID']) : null,
        status: "",
        _rowNumber: item._rowNumber
      };

      var action = rowUpdates.action.toString().toUpperCase();
        var targetEmail = item['Member Email'];
        var targetRole = item['Role'] || 'ROLE_MEMBER';
        var spaceId = item['Space ID'];

        if (!spaceId) throw new Error("⚠️ Data Error: Missing Space ID");

        switch (action) {
          case "ADD_MEMBER":
            if (!targetEmail) throw new Error("⚠️ Data Error: Missing Member Email");
            
            var membership = {
              member: {
                name: "users/" + targetEmail,
                type: "HUMAN"
              },
              role: targetRole
            };

            var newMembership = _App_callWithBackoff(function () {
              return Chat.Spaces.Members.create(membership, spaceId);
            });

            rowUpdates.membershipId = newMembership.name;
            rowUpdates.status = _App_formatStatus('SUCCESS', "Added");
            rowUpdates.action = "";
            break;

          case "REMOVE_MEMBER":
            if (!rowUpdates.membershipId) throw new Error("⚠️ Data Error: Missing Membership ID for REMOVE");
            
            _App_callWithBackoff(function () {
               Chat.Spaces.Members.remove(rowUpdates.membershipId);
            });
            
            rowUpdates.status = _App_formatStatus('SUCCESS', "Removed");
            rowUpdates.action = "";
            break;

          default:
            throw new Error("❓ Unknown Action '" + action + "'");
        }

        return rowUpdates;

    }, {
      onBatchComplete: function (batchResults) {
        _App_batchPatchResults('CHAT_SYNC', batchResults, function (res) {
          return {
            'Membership ID': res.membershipId
          };
        });
      }
    });

      return _App_ok("Sync Complete. Processed: " + stats.processedCount);
    });
  });
}

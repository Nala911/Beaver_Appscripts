/**
 * Google Chat Space Sync Tool
 * Version: 1.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CHAT_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Chat API', test: function() { return typeof Chat !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CHAT_SPACE_SYNC,
    TITLE: SHEET_NAMES.CHAT_SPACE_SYNC,
    MENU_LABEL: SHEET_NAMES.CHAT_SPACE_SYNC,
    MENU_ENTRYPOINT: 'ChatSpaceSync_showSidebar',
    MENU_ORDER: 15,
    SIDEBAR_HTML: 'ChatSpaceSync_Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [120, 250, 250, 150, 150, 200, 250],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 2,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['ADD_MEMBER', 'REMOVE_MEMBER'] },
            { header: 'Space Name', type: 'TEXT' },
            { header: 'Member Email', type: 'TEXT' },
            { header: 'Role', type: 'DROPDOWN', options: ['ROLE_MEMBER', 'ROLE_MANAGER'] },
            { header: 'Type', type: 'TEXT' }, // User or Group
            { header: 'Space ID', type: 'ID' },
            { header: 'Membership ID', type: 'ID' }
        ]
    }
});

// --- MENU & UI HANDLERS ---

/** Opens the Chat Sync sidebar and ensures the sheet exists. */
function ChatSpaceSync_showSidebar() {
  return Logger.run('CHAT_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CHAT_SYNC');
  });
}

// --- API FOR SIDEBAR ---

function ChatSpaceSync_getLoadData() {
  return Logger.run('CHAT_SYNC', 'Load Data', function () {
    try {
      var spacesList = [];
      var pageToken = null;
      
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

      var uniqueSpaces = spacesList.map(function(s) {
        return {
          id: s.name, // Space names are their unique IDs in Chat API (e.g. "spaces/12345")
          name: s.displayName || s.name
        };
      });

      var savedSpaceIds = _App_getProperty(APP_PROPS.CHAT_SELECTED_SPACES);
      if (!Array.isArray(savedSpaceIds)) savedSpaceIds = [];

      return _App_ok('Spaces loaded.', {
        spaces: uniqueSpaces,
        savedSpaceIds: savedSpaceIds
      });
    } catch (err) {
      throw new Error('Unable to load spaces. ' + err.message);
    }
  });
}

function ChatSpaceSync_savePreferences(spaceIds) {
  return Logger.run('CHAT_SYNC', 'Save Preferences', function () {
    if (spaceIds) _App_setProperty(APP_PROPS.CHAT_SELECTED_SPACES, spaceIds);
    return _App_ok('Preferences saved.');
  });
}

// --- THE "PULL" WORKFLOW ---

function ChatSpaceSync_pullMembers(request) {
  return Logger.run('CHAT_SYNC', 'Pull Members', function () {
    var TARGET_SHEET_NAME = SHEET_NAMES.CHAT_SPACE_SYNC;
    var sheet = _App_ensureSheetExists('CHAT_SYNC');

    Logger.info(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Pull Members', 'Pull started — spaces: [' + request.spaceIds.join(', ') + ']');

    var outputObjects = [];

    request.spaceIds.forEach(function (spaceNameId) {
      try {
        var space = _App_callWithBackoff(function () { return Chat.Spaces.get(spaceNameId); });
        if (!space) return;

        var pageToken = null;
        var members = [];
        do {
            var response = _App_callWithBackoff(function() {
                return Chat.Spaces.Members.list(spaceNameId, {
                    pageToken: pageToken
                });
            });
            if (response.memberships) {
                members = members.concat(response.memberships);
            }
            pageToken = response.nextPageToken;
        } while (pageToken);

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
            'Space Name': space.displayName || space.name,
            'Member Email': memberEmail,
            'Role': m.role === 'ROLE_MANAGER' ? 'ROLE_MANAGER' : 'ROLE_MEMBER',
            'Type': memberType,
            'Space ID': space.name,
            'Membership ID': m.name
          });
        });
        Logger.info(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Pull Members', 'Pulled ' + members.length + ' member(s) from space: ' + (space.displayName || space.name));
      } catch (err) {
        Logger.error(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Pull Members — ' + spaceNameId, err);
      }
    });

    SheetManager.overwriteObjects('CHAT_SYNC', outputObjects);
    ChatSpaceSync_savePreferences(request.spaceIds);
    
    var summary = 'Successfully imported ' + outputObjects.length + " members into '" + TARGET_SHEET_NAME + "'.";
    Logger.info(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Pull Members', summary);
    return _App_ok(summary);
  });
}

function ChatSpaceSync_checkForUnsavedChanges() {
  return Logger.run('CHAT_SYNC', 'Check Unsaved', function () {
    return _App_ok('Check complete.', SheetManager.hasPendingActions('CHAT_SYNC'));
  });
}

// --- THE "PUSH" WORKFLOW ---

function ChatSpaceSync_pushChanges() {
  return Logger.run('CHAT_SYNC', 'Push Changes', function () {
    var pendingItems = SheetManager.readPendingObjects('CHAT_SYNC');

    if (pendingItems.length === 0) return _App_ok("No pending actions found.");

    Logger.info(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Global', 'Push started — processing ' + pendingItems.length + ' pending row(s)');

    var stats = _App_BatchProcessor('CHAT_SYNC', pendingItems, function (item) {
      var rowUpdates = {
        action: item['Action'],
        membershipId: item['Membership ID'] ? String(item['Membership ID']) : null,
        status: "",
        _rowNumber: item._rowNumber
      };

      try {
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
            rowUpdates.status = "✅ Added";
            rowUpdates.action = "";
            break;

          case "REMOVE_MEMBER":
            if (!rowUpdates.membershipId) throw new Error("⚠️ Data Error: Missing Membership ID for REMOVE");
            
            _App_callWithBackoff(function () {
               Chat.Spaces.Members.remove(rowUpdates.membershipId);
            });
            
            rowUpdates.status = "🗑️ Removed";
            rowUpdates.action = "";
            break;

          default:
            rowUpdates.status = "❓ Unknown Action '" + action + "'";
        }

        Logger.info(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Row ' + item._rowNumber, rowUpdates.status);
        return rowUpdates;

      } catch (e) {
        rowUpdates.status = e.message;
        Logger.error(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Row ' + item._rowNumber, e);
        return rowUpdates;
      }
    }, {
      onBatchComplete: function (batchResults) {
        var rowNumbers = [];
        var patchData = [];
        batchResults.forEach(function (res) {
          if (res && res._rowNumber !== undefined) {
            rowNumbers.push(res._rowNumber);
            patchData.push({
              'Action': res.action,
              'Membership ID': res.membershipId
            });
          }
        });
        if (rowNumbers.length > 0) {
          SheetManager.batchPatchRows('CHAT_SYNC', rowNumbers, patchData);
        }
      }
    });

    return _App_ok("Sync Complete. Processed: " + stats.processedCount);
  });
}

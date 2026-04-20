/**
 * Google Chat Space Sync Tool
 * Version: 1.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CHAT_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Chat API', test: function() { return typeof Chat !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CHAT_SPACE_SYNC,
    TITLE: '💬 Chat Space Manager',
    MENU_LABEL: '💬 Google Chat',
    MENU_ENTRYPOINT: 'ChatSync_showSidebar',
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
function ChatSync_showSidebar() {
  return Logger.run('CHAT_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CHAT_SYNC');
  });
}

// --- API FOR SIDEBAR ---

function ChatSync_getLoadData() {
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

      var prefs = SyncEngine.getPrefs('CHAT_SYNC');

      return _App_ok('Spaces loaded.', {
        spaces: uniqueSpaces,
        savedSpaceIds: prefs.selectedSpaceIds || []
      });
    } catch (err) {
      throw new Error('Unable to load spaces. ' + err.message);
    }
  });
}

function ChatSync_savePreferences(spaceIds) {
  var prefs = SyncEngine.getPrefs('CHAT_SYNC');
  if (spaceIds) prefs.selectedSpaceIds = spaceIds;
  SyncEngine.setPrefs('CHAT_SYNC', prefs);
}

// --- THE "PULL" WORKFLOW ---

function ChatSync_pullMembers(request) {
  return Logger.run('CHAT_SYNC', 'Pull Members', function () {
    var TARGET_SHEET_NAME = SHEET_NAMES.CHAT_SPACE_SYNC;
    var sheet = _App_ensureSheetExists('CHAT_SYNC');

    Logger.info(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Pull Members', 'Pull started — spaces: [' + request.spaceIds.join(', ') + ']');

    SheetManager.clearData('CHAT_SYNC');

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
    ChatSync_savePreferences(request.spaceIds);
    
    var summary = 'Successfully imported ' + outputObjects.length + " members into '" + TARGET_SHEET_NAME + "'.";
    Logger.info(SyncEngine.getTool('CHAT_SYNC').TITLE, 'Pull Members', summary);
    return _App_ok(summary);
  });
}

function ChatSync_checkForUnsavedChanges() {
  return SheetManager.hasPendingActions('CHAT_SYNC');
}

// --- THE "PUSH" WORKFLOW ---

function ChatSync_pushChanges() {
  return Logger.run('CHAT_SYNC', 'Push Changes', function () {
    
    var stats = ExecutionService.processPendingRows('CHAT_SYNC', function(rowObj) {
        var action = String(rowObj['Action'] || '').toUpperCase();
        var targetEmail = rowObj['Member Email'];
        var targetRole = rowObj['Role'] || 'ROLE_MEMBER';
        var spaceId = rowObj['Space ID'];
        var membershipId = rowObj['Membership ID'];

        if (!spaceId) throw new Error("⚠️ Data Error: Missing Space ID");

        var updates = {
            'Action': '',
            'Log': ''
        };

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

            var newMembership = Chat.Spaces.Members.create(membership, spaceId);

            updates['Membership ID'] = newMembership.name;
            updates['Log'] = "✅ Added";
            break;

          case "REMOVE_MEMBER":
            if (!membershipId) throw new Error("⚠️ Data Error: Missing Membership ID for REMOVE");
            
            Chat.Spaces.Members.remove(membershipId);
            updates['Log'] = "🗑️ Removed";
            break;

          default:
            updates['Log'] = "❓ Unknown Action '" + action + "'";
            updates['Action'] = rowObj['Action']; // Keep action if unknown
        }

        SheetManager.patchRow('CHAT_SYNC', rowObj._rowNumber, updates);
    });

    if (stats.processed === 0 && stats.errors === 0) {
        return _App_ok("No data to sync.");
    }

    return _App_ok("Sync Complete. Success: " + stats.processed + ", Errors: " + stats.errors);
  });
}

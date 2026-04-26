/**
 * Google Chat Space Sync Tool
 * Version: 1.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CHAT_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Chat API', test: function() { return typeof Chat !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CHAT_SPACE_SYNC,
    TITLE: SHEET_NAMES.CHAT_SPACE_SYNC,
    MENU_LABEL: SHEET_NAMES.CHAT_SPACE_SYNC,
    MENU_ENTRYPOINT: 'ChatSpaceSync_openSidebar',
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
            { header: 'Status', type: 'STATUS' },
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
function ChatSpaceSync_openSidebar() {
  return Logger.run('CHAT_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CHAT_SYNC');
  });
}

// --- THE "PULL" WORKFLOW ---

function ChatSpaceSync_pullMembers() {
  return Logger.run('CHAT_SYNC', 'Pull Members', function () {
    var TARGET_SHEET_NAME = SHEET_NAMES.CHAT_SPACE_SYNC;
    var sheet = _App_ensureSheetExists('CHAT_SYNC');

    var outputObjects = [];
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
            rowUpdates.status = SHEET_THEME.STATUS_PREFIXES.SUCCESS + "Added";
            rowUpdates.action = "";
            break;

          case "REMOVE_MEMBER":
            if (!rowUpdates.membershipId) throw new Error("⚠️ Data Error: Missing Membership ID for REMOVE");
            
            _App_callWithBackoff(function () {
               Chat.Spaces.Members.remove(rowUpdates.membershipId);
            });
            
            rowUpdates.status = SHEET_THEME.STATUS_PREFIXES.SUCCESS + "Removed";
            rowUpdates.action = "";
            break;

          default:
            throw new Error("❓ Unknown Action '" + action + "'");
        }

        return rowUpdates;

    }, {
      onBatchComplete: function (batchResults) {
        var rowNumbers = [];
        var patchData = [];
        var prefixes = SHEET_THEME.STATUS_PREFIXES;
        batchResults.forEach(function (res) {
          if (res && res._rowNumber !== undefined) {
            rowNumbers.push(res._rowNumber);
            if (res.isError) {
              patchData.push({ 'Status': prefixes.ERROR + res.error });
            } else {
              patchData.push({
                'Action': res.action,
                'Status': res.status,
                'Membership ID': res.membershipId
              });
            }
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

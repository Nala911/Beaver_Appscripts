/**
 * Google Calendar
 * Version: 6.0 (Plugin Architecture — registers with SyncEngine)
 */

SyncEngine.registerTool('CALENDAR_SYNC', {
    REQUIRED_SERVICES: [ { name: 'Calendar API', test: function() { return typeof Calendar !== 'undefined'; } } ],
    SHEET_NAME: SHEET_NAMES.CALENDAR_SYNC,
    TITLE: SHEET_NAMES.CALENDAR_SYNC,
    MENU_LABEL: SHEET_NAMES.CALENDAR_SYNC,
    MENU_ENTRYPOINT: 'CalendarSync_openSidebar',
    MENU_ORDER: 10,
    SIDEBAR_HTML: 'CalendarSync_Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FROZEN_COLS: 0,
    COL_WIDTHS: [120, 180, 200, null, null, 250, null, null, null, null, null, null, null],
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 2,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'MOVE', 'REMOVE'] },
            { header: 'Target Calendar Name', type: 'DROPDOWN', options: function() { try { return CalendarApp.getAllCalendars().map(function(c){return c.getName()}); } catch(e) { return []; } } },
            { header: 'Event Title', type: 'TEXT' },
            { header: 'Start Time', type: 'TEXT' },
            { header: 'End Time', type: 'TEXT' },
            { header: 'Description', type: 'TEXT' },
            { header: 'Location', type: 'TEXT' },
            { header: 'Add Meet?', type: 'DROPDOWN', options: ['Yes', 'No'] },
            { header: 'Guests', type: 'EMAIL_LIST' },
            { header: 'Color', type: 'DROPDOWN', options: function() { try { return ['Default'].concat(Object.keys(CalendarApp.EventColor)); } catch(e){ return ['Default']; } } },
            { header: 'Visibility', type: 'DROPDOWN', options: ['Default', 'Public', 'Private'] },
            { header: 'Event ID', type: 'ID' },
            { header: 'Original Calendar ID', type: 'ID' }
        ]
    }
});

// --- CONFIGURATION & CONSTANTS ---
// Metadata (title, sidebar, headers, widths) lives in SyncEngine.getTool('CALENDAR_SYNC').
var CALENDAR_SYNC_CFG = {
  COLORS: ["Default"].concat(Object.keys(CalendarApp.EventColor)),
  VISIBILITY: ["Default", "Public", "Private"]
};

// --- MENU & UI HANDLERS ---

/** @deprecated — Use _App_ensureSheetExists('CALENDAR_SYNC') instead. */
function _CalendarSync_ensureSheetExistsAndActivate() {
  return _App_ensureSheetExists('CALENDAR_SYNC');
}

/** Opens the Calendar sidebar and ensures the sheet exists. */
function CalendarSync_openSidebar() {
  return Logger.run('CALENDAR_SYNC', 'Open Sidebar', function () {
    _App_launchTool('CALENDAR_SYNC');
  });
}


// --- API FOR SIDEBAR ---

function CalendarSync_getLoadData() {
  return Logger.run('CALENDAR_SYNC', 'Load Data', function () {
    try {
      var allCalendars = _App_callWithBackoff(function () {
        return CalendarApp.getAllCalendars();
      });
      var seen = {};
      var uniqueCals = [];

      (allCalendars || []).forEach(function (c) {
        var calId = c && c.getId ? c.getId() : '';
        if (!calId || seen[calId]) return;
        seen[calId] = true;
        uniqueCals.push({
          id: calId,
          name: c.getName(),
          color: c.getColor()
        });
      });

      var savedCalIds = _App_getProperty(APP_PROPS.CAL_SELECTED_IDS);
      if (!Array.isArray(savedCalIds)) savedCalIds = [];

      return _App_ok('Calendar load data ready.', {
        calendars: uniqueCals,
        savedCalIds: savedCalIds,
        savedStartDate: _App_getProperty(APP_PROPS.CAL_START_DATE),
        savedEndDate: _App_getProperty(APP_PROPS.CAL_END_DATE)
      });
    } catch (err) {
      throw new Error('Unable to load calendars. ' + err.message);
    }
  });
}

function CalendarSync_savePreferences(calIds, startStr, endStr) {
  if (calIds) _App_setProperty(APP_PROPS.CAL_SELECTED_IDS, calIds);
  if (startStr !== undefined) _App_setProperty(APP_PROPS.CAL_START_DATE, startStr);
  if (endStr !== undefined) _App_setProperty(APP_PROPS.CAL_END_DATE, endStr);
}

// --- PART 2: THE "PULL" WORKFLOW ---

function CalendarSync_pullEvents(request) {
  return Logger.run('CALENDAR_SYNC', 'Pull Events', function () {
    var TARGET_SHEET_NAME = SHEET_NAMES.CALENDAR_SYNC;
    var sheet = _CalendarSync_ensureSheetExistsAndActivate();

    Logger.info(SyncEngine.getTool('CALENDAR_SYNC').TITLE, 'Pull Events', 'Pull started — calendars: [' + request.calIds.join(', ') + '] from ' + request.startDate + ' to ' + request.endDate);

    SheetManager.clearData('CALENDAR_SYNC');

    // Fetch Events
    var start = new Date(request.startDate);
    var end = new Date(request.endDate);
    var outputObjects = [];

    request.calIds.forEach(function (calId) {
      try {
        var cal = _App_callWithBackoff(function () { return CalendarApp.getCalendarById(calId); });
        if (!cal) return;

        var events = _App_callWithBackoff(function () { return cal.getEvents(start, end); });
        events.forEach(function (e) {
          outputObjects.push({
            'Action': "",
            'Target Calendar Name': cal.getName(),
            'Event Title': e.getTitle(),
            'Start Time': Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss"),
            'End Time': Utilities.formatDate(e.getEndTime(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss"),
            'Description': e.getDescription(),
            'Location': e.getLocation(),
            'Add Meet?': "",
            'Guests': e.getGuestList().map(function (g) { return g.getEmail(); }).join(","),
            'Color': "Default",
            'Visibility': "Default",
            'Event ID': e.getId(),
            'Original Calendar ID': cal.getId()
          });
        });
        Logger.info(SyncEngine.getTool('CALENDAR_SYNC').TITLE, 'Pull Events', 'Pulled ' + events.length + ' event(s) from calendar: ' + cal.getName());
      } catch (err) {
        Logger.error(SyncEngine.getTool('CALENDAR_SYNC').TITLE, 'Pull Events — ' + calId, err);
      }
    });

    // Populate Sheet via DAO
    SheetManager.overwriteObjects('CALENDAR_SYNC', outputObjects);

    CalendarSync_savePreferences(request.calIds, request.startDate, request.endDate);
    var summary = 'Successfully imported ' + outputObjects.length + " events into '" + TARGET_SHEET_NAME + "'.";
    Logger.info(SyncEngine.getTool('CALENDAR_SYNC').TITLE, 'Pull Events', summary);
    return _App_ok(summary);
  });
}

function CalendarSync_checkForUnsavedChanges() {
  return SheetManager.hasPendingActions('CALENDAR_SYNC');
}

// --- PART 2: THE "PUSH" WORKFLOW ---

function CalendarSync_pushChanges() {
  return Logger.run('CALENDAR_SYNC', 'Push Changes', function () {
    var pendingItems = SheetManager.readPendingObjects('CALENDAR_SYNC');

    if (pendingItems.length === 0) return "No pending actions found.";

    Logger.info(SyncEngine.getTool('CALENDAR_SYNC').TITLE, 'Global', 'Push started — processing ' + pendingItems.length + ' pending row(s)');

    var allCals = CalendarApp.getAllCalendars();
    var calMap = new Map();
    var calObjMap = new Map();

    allCals.forEach(function (c) {
      calMap.set(c.getName(), c.getId());
      calObjMap.set(c.getId(), c);
    });

    var stats = _App_BatchProcessor('CALENDAR_SYNC', pendingItems, function (item) {
      var rowUpdates = {
        action: item['Action'],
        eventId: item['Event ID'] ? String(item['Event ID']) : null,
        calId: item['Original Calendar ID'] ? String(item['Original Calendar ID']) : null,
        status: "",
        _rowNumber: item._rowNumber
      };

      try {
        var action = rowUpdates.action.toString().toUpperCase();
        var targetCalName = item['Target Calendar Name'];
        var targetCalId = calMap.get(targetCalName);

        var eventData = {
          title: item['Event Title'],
          start: item['Start Time'],
          end: item['End Time'],
          desc: item['Description'],
          loc: item['Location'],
          meet: item['Add Meet?'],
          guests: item['Guests'],
          color: item['Color'],
          visibility: item['Visibility']
        };

        if (!(eventData.start instanceof Date)) eventData.start = new Date(eventData.start);
        if (!(eventData.end instanceof Date)) eventData.end = new Date(eventData.end);

        if (isNaN(eventData.start.getTime())) throw new Error("⚠️ Data Error: Invalid Start Time format");
        if (isNaN(eventData.end.getTime())) throw new Error("⚠️ Data Error: Invalid End Time format");
        if (eventData.end <= eventData.start) throw new Error("⚠️ Data Error: End Time cannot be before or equal to Start Time");
        if (!eventData.title) throw new Error("⚠️ Data Error: Missing Event Title");

        if (eventData.guests) {
          var invalidEmails = eventData.guests.split(',').map(function (g) { return g.trim() }).filter(function (e) { return e && !_CalendarSync_validateEmail(e) });
          if (invalidEmails.length > 0) throw new Error("⚠️ Data Error: Invalid guest email(s): " + invalidEmails.join(', '));
        }

        switch (action) {
          case "CREATE":
            if (!targetCalName) throw new Error("⚠️ Data Error: Missing Target Calendar Name");
            if (!targetCalId) throw new Error("⚠️ Data Error: Calendar '" + targetCalName + "' not found");

            var createCal = calObjMap.get(targetCalId);
            if (!createCal) throw new Error("❌ API Error: Target calendar object is null");

            var newEvent = _App_callWithBackoff(function () {
              return createCal.createEvent(eventData.title, eventData.start, eventData.end, {
                description: eventData.desc,
                location: eventData.loc,
                guests: eventData.guests ? eventData.guests.split(',').map(function (g) { return g.trim(); }).join(',') : ""
              });
            });

            var optionErr = _CalendarSync_applyEventOptions(newEvent, eventData);
            if (eventData.meet === 'Yes') {
              try { _CalendarSync_addMeetLinkToEvent(targetCalId, newEvent.getId()); }
              catch (meetErr) { optionErr = optionErr ? optionErr + ", Meet: " + meetErr.message : "Meet: " + meetErr.message; }
            }

            rowUpdates.eventId = newEvent.getId();
            rowUpdates.calId = targetCalId;
            rowUpdates.status = "✅ Created" + (optionErr ? " (⚠️ " + optionErr + ")" : "");
            rowUpdates.action = "";
            break;

          case "UPDATE":
            if (!rowUpdates.eventId) throw new Error("⚠️ Data Error: Missing Event ID for UPDATE");

            if (targetCalId && rowUpdates.calId && targetCalId !== rowUpdates.calId) {
              rowUpdates = _CalendarSync_processMove(rowUpdates, calObjMap, targetCalId, eventData);
              break;
            }

            var updateCal = calObjMap.get(rowUpdates.calId);
            var eventToUpdate = _CalendarSync_findEvent(updateCal, true, rowUpdates.eventId, eventData);

            if (!eventToUpdate) throw new Error("⚠️ Data Error: Event ID not found on calendar");

            _App_callWithBackoff(function () {
              eventToUpdate.setTitle(eventData.title);
              eventToUpdate.setTime(eventData.start, eventData.end);
              eventToUpdate.setDescription(eventData.desc);
              eventToUpdate.setLocation(eventData.loc);
            });

            var updateOptionErr = _CalendarSync_applyEventOptions(eventToUpdate, eventData);
            if (eventData.meet === 'Yes') {
              try { _CalendarSync_addMeetLinkToEvent(rowUpdates.calId || eventToUpdate.getOriginalCalendarId(), rowUpdates.eventId); }
              catch (meetErr) { updateOptionErr = updateOptionErr ? updateOptionErr + ", Meet: " + meetErr.message : "Meet: " + meetErr.message; }
            }

            var currentGuests = eventToUpdate.getGuestList();
            var targetGuests = eventData.guests ? eventData.guests.split(',').map(function (g) { return g.trim(); }).filter(function (g) { return g !== ""; }) : [];

            currentGuests.forEach(function (guestObj) {
              var email = guestObj.getEmail();
              if (targetGuests.indexOf(email) === -1) eventToUpdate.removeGuest(email);
            });
            targetGuests.forEach(function (email) { eventToUpdate.addGuest(email); });

            rowUpdates.status = "✅ Updated" + (updateOptionErr ? " (⚠️ " + updateOptionErr + ")" : "");
            rowUpdates.action = "";
            break;

          case "MOVE":
            if (!targetCalId) throw new Error("⚠️ Data Error: Target Calendar '" + targetCalName + "' not found");
            rowUpdates = _CalendarSync_processMove(rowUpdates, calObjMap, targetCalId, eventData);
            break;

          case "REMOVE":
            if (!rowUpdates.eventId) throw new Error("⚠️ Data Error: Missing Event ID for REMOVE");
            var delCal = calObjMap.get(rowUpdates.calId);
            if (!delCal) throw new Error("⚠️ Data Error: Original Calendar inaccessible");

            var eventToDel = _CalendarSync_findEvent(delCal, false, rowUpdates.eventId, eventData);
            if (eventToDel) {
              _App_callWithBackoff(function () { eventToDel.deleteEvent(); });
              rowUpdates.status = "🗑️ Removed";
              rowUpdates.action = "";
            } else {
              rowUpdates.status = "⚠️ Already Deleted (Event not found)";
              rowUpdates.action = "";
            }
            break;

          default:
            rowUpdates.status = "❓ Unknown Action '" + action + "'";
        }

        var isError = rowUpdates.status && (rowUpdates.status.indexOf('❌') > -1 || rowUpdates.status.indexOf('⚠️') > -1);
        if (isError) Logger.error(SyncEngine.getTool('CALENDAR_SYNC').TITLE, 'Row ' + item._rowNumber, rowUpdates.status);
        else if (rowUpdates.status) Logger.info(SyncEngine.getTool('CALENDAR_SYNC').TITLE, 'Row ' + item._rowNumber, rowUpdates.status);

        return rowUpdates;

      } catch (e) {
        rowUpdates.status = e.message;
        Logger.error(SyncEngine.getTool('CALENDAR_SYNC').TITLE, 'Row ' + item._rowNumber, e);
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
              'Event ID': res.eventId,
              'Original Calendar ID': res.calId
            });
          }
        });
        if (rowNumbers.length > 0) {
          SheetManager.batchPatchRows('CALENDAR_SYNC', rowNumbers, patchData);
        }
      }
    });

    return _App_ok("Sync Complete. Processed: " + stats.processedCount);
  });
}

// --- HELPER VALIDATORS ---

function _CalendarSync_validateEmail(email) {
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

function _CalendarSync_processMove(rowUpdates, calObjMap, targetCalId, eventData) {
  var oldCal = calObjMap.get(rowUpdates.calId);
  var newCal = calObjMap.get(targetCalId);

  if (!newCal) throw new Error("⚠️ Data Error: Target calendar not accessible");

  var newEvent = newCal.createEvent(eventData.title, eventData.start, eventData.end, {
    description: eventData.desc,
    location: eventData.loc,
    guests: eventData.guests ? eventData.guests.split(',').map(function (g) { return g.trim(); }).join(',') : ""
  });
  _CalendarSync_applyEventOptions(newEvent, eventData);

  var meetWarning = "";
  if (eventData.meet === 'Yes') {
    try { _CalendarSync_addMeetLinkToEvent(targetCalId, newEvent.getId()); }
    catch (meetErr) { meetWarning = " (⚠️ Meet: " + meetErr.message + ")"; }
  }

  var deleteWarning = "";
  if (oldCal && rowUpdates.eventId) {
    try {
      var oldEvent = _CalendarSync_findEvent(oldCal, false, rowUpdates.eventId, eventData);
      if (oldEvent) oldEvent.deleteEvent();
      else deleteWarning = " (⚠️ Old event not found)";
    } catch (delErr) {
      deleteWarning = " (⚠️ Could not delete old event: " + delErr.message + ")";
    }
  } else {
    deleteWarning = " (⚠️ Old calendar inaccessible)";
  }

  rowUpdates.eventId = newEvent.getId();
  rowUpdates.calId = targetCalId;
  rowUpdates.status = "✅ Moved" + deleteWarning + meetWarning;
  rowUpdates.action = "";
  return rowUpdates;
}

function _CalendarSync_applyEventOptions(event, data) {
  var warning = null;
  if (data.color && data.color !== 'Default') {
    if (CalendarApp.EventColor[data.color]) {
      try { event.setColor(CalendarApp.EventColor[data.color]); } catch (e) { warning = "Color set failed"; }
    } else {
      warning = "Invalid Color";
    }
  }
  if (data.visibility) {
    try {
      if (data.visibility === 'Public') event.setVisibility(CalendarApp.Visibility.PUBLIC);
      else if (data.visibility === 'Private') event.setVisibility(CalendarApp.Visibility.PRIVATE);
    } catch (e) {
      warning = warning ? warning + ", Visibility set failed" : "Visibility set failed";
    }
  }
  return warning;
}

// --- UTILITIES ---

function _CalendarSync_findEvent(cal, allowGlobal, eventId, eventData) {
  var event = null;
  eventId = String(eventId).trim();

  // 1. Direct fetch
  if (cal) {
    try { event = _App_callWithBackoff(function() { return cal.getEventById(eventId); }); } catch(e){}
  }
  if (!event && allowGlobal) {
    try { event = _App_callWithBackoff(function() { return CalendarApp.getEventById(eventId); }); } catch(e){}
  }

  // 2. Fallback 1: Append @google.com (Common for CSV imports)
  if (!event && eventId.indexOf('@google.com') === -1) {
    var suffixed = eventId + '@google.com';
    if (cal) {
      try { event = _App_callWithBackoff(function() { return cal.getEventById(suffixed); }); } catch(e){}
    }
    if (!event && allowGlobal) {
      try { event = _App_callWithBackoff(function() { return CalendarApp.getEventById(suffixed); }); } catch(e){}
    }
  }

  // 3. Fallback 2: Strip @google.com (Just in case)
  if (!event && eventId.indexOf('@google.com') > -1) {
    var stripped = eventId.split('@')[0];
    if (cal) {
      try { event = _App_callWithBackoff(function() { return cal.getEventById(stripped); }); } catch(e){}
    }
    if (!event && allowGlobal) {
      try { event = _App_callWithBackoff(function() { return CalendarApp.getEventById(stripped); }); } catch(e){}
    }
  }

  // 4. Fallback 3: Time-window Search (Foolproof for CSV imports missing proper IDs)
  if (!event && cal && eventData && eventData.start && eventData.end) {
    try {
      // Create a +/- 24h search window to account for timezone drift
      var sTime = new Date(eventData.start).getTime();
      var eTime = new Date(eventData.end).getTime();
      if (!isNaN(sTime) && !isNaN(eTime)) {
        var searchStart = new Date(sTime - 86400000);
        var searchEnd = new Date(eTime + 86400000);
        
        var eventsInRange = _App_callWithBackoff(function() { return cal.getEvents(searchStart, searchEnd); });
        
        for (var i = 0; i < eventsInRange.length; i++) {
          var eId = eventsInRange[i].getId();
          // Match if it's identical, or if one is a base prefix of the other
          if (eId === eventId || eId.indexOf(eventId) === 0 || eventId.indexOf(eId) === 0) {
            event = eventsInRange[i];
            break;
          }
        }
      }
    } catch(err) {
      // Ignore inner search failures (let the caller handle the missing event)
    }
  }

  return event;
}

function _CalendarSync_addMeetLinkToEvent(calendarId, eventId) {
  if (typeof Calendar === 'undefined') {
    throw new Error("Enable 'Google Calendar API' in Services");
  }
  Calendar.Events.patch({
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: "hangoutsMeet" }
      }
    }
  }, calendarId, eventId, { conferenceDataVersion: 1 });
}

function _CalendarSync_setupSheetStructure(sheet) {
  // Logic mostly shifted to APP_REGISTRY, kept for backward compatibility if ever called directly
  var headers = SyncEngine.getTool('CALENDAR_SYNC').HEADERS;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

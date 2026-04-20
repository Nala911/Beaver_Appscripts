/**
 * CALENDAR SYNC SERVICE
 * ==========================================
 * Handles synchronization between Google Sheets and Google Calendar.
 */

App.Engine.registerTool('CALENDAR_SYNC', {
    SHEET_NAME: App.Config.SHEET_NAMES.CALENDAR_SYNC,
    TITLE: '📅 Calendar Sync Master',
    MENU_LABEL: '🗓️ Google Calendar',
    MENU_ENTRYPOINT: 'Calendar_showSidebar',
    MENU_ORDER: 10,
    SIDEBAR_HTML: 'Calendar_Sidebar',
    SIDEBAR_WIDTH: 400,
    FROZEN_ROWS: 1,
    FORMAT_CONFIG: {
        numReadOnlyColsAtEnd: 2,
        conditionalRules: [{ type: 'pending', actionCol: 'A', scope: 'actionOnly' }],
        COL_SCHEMA: [
            { header: 'Action', type: 'ACTION', options: ['CREATE', 'UPDATE', 'MOVE', 'REMOVE'] },
            { header: 'Target Calendar Name', type: 'DROPDOWN', options: function() { try { return CalendarApp.getAllCalendars().map(function(c){return c.getName()}); } catch(e) { return []; } } },
            { header: 'Event Title', type: 'TEXT' },
            { header: 'Start Time', type: 'DATETIME' },
            { header: 'End Time', type: 'DATETIME' },
            { header: 'Description', type: 'TEXT' },
            { header: 'Location', type: 'TEXT' },
            { header: 'Add Meet?', type: 'DROPDOWN', options: ['Yes', 'No'] },
            { header: 'Guests', type: 'EMAIL_LIST' },
            { header: 'Color', type: 'DROPDOWN', options: function() { try { return ['Default'].concat(Object.keys(CalendarApp.EventColor)); } catch(e){ return ['Default']; } } },
            { header: 'Visibility', type: 'DROPDOWN', options: ['Default', 'Public', 'Private'] },
            { header: 'Event ID', type: 'ID' },
            { header: 'Original Calendar ID', type: 'ID' }
        ]
    },

    /**
     * Tool Service Logic
     */
    service: {
        getLoadData: function() {
            var cals = CalendarApp.getAllCalendars().map(function(c) {
                return { id: c.getId(), name: c.getName(), color: c.getColor() };
            });
            var prefs = App.Engine.getPrefs('CALENDAR_SYNC');
            return _App_ok('Data loaded', {
                calendars: cals,
                savedCalIds: prefs.selectedCalIds || [],
                savedStartDate: prefs.startDate || '',
                savedEndDate: prefs.endDate || ''
            });
        },

        pullEvents: function(req) {
            App.Data.clearData('CALENDAR_SYNC');
            var start = new Date(req.startDate);
            var end = new Date(req.endDate);
            var results = [];

            req.calIds.forEach(function(id) {
                var cal = CalendarApp.getCalendarById(id);
                if (!cal) return;
                cal.getEvents(start, end).forEach(function(e) {
                    results.push({
                        'Target Calendar Name': cal.getName(),
                        'Event Title': e.getTitle(),
                        'Start Time': e.getStartTime(),
                        'End Time': e.getEndTime(),
                        'Description': e.getDescription(),
                        'Location': e.getLocation(),
                        'Guests': e.getGuestList().map(function(g) { return g.getEmail(); }),
                        'Event ID': e.getId(),
                        'Original Calendar ID': id
                    });
                });
            });

            App.Data.writeObjects('CALENDAR_SYNC', results);
            App.Engine.setPrefs('CALENDAR_SYNC', { selectedCalIds: req.calIds, startDate: req.startDate, endDate: req.endDate });
            return _App_ok("Pulled " + results.length + " events.");
        },

        pushChanges: function() {
            var allCals = CalendarApp.getAllCalendars();
            var calMap = {};
            allCals.forEach(function(c) { calMap[c.getName()] = c; calMap[c.getId()] = c; });

            var pending = App.Data.readPendingActions('CALENDAR_SYNC');
            var processed = 0, errors = 0;

            pending.forEach(function(row) {
                try {
                    var action = String(row['Action']).toUpperCase();
                    var targetCal = calMap[row['Target Calendar Name']];
                    var originalCal = calMap[row['Original Calendar ID']];
                    var eventId = row['Event ID'];

                    var updates = { 'Action': '', 'Log': '✅ Success' };

                    switch (action) {
                        case 'CREATE':
                            if (!targetCal) throw new Error("Target calendar not found");
                            var ev = targetCal.createEvent(row['Event Title'], row['Start Time'], row['End Time'], {
                                description: row['Description'],
                                location: row['Location'],
                                guests: (row['Guests'] || []).join(',')
                            });
                            updates['Event ID'] = ev.getId();
                            updates['Original Calendar ID'] = targetCal.getId();
                            break;

                        case 'UPDATE':
                            var evToUpdate = _CalendarSync_findEvent(originalCal, true, eventId, row);
                            if (!evToUpdate) throw new Error("Event not found");
                            evToUpdate.setTitle(row['Event Title']);
                            evToUpdate.setTime(row['Start Time'], row['End Time']);
                            evToUpdate.setDescription(row['Description']);
                            evToUpdate.setLocation(row['Location']);
                            break;

                        case 'REMOVE':
                            var evToDel = _CalendarSync_findEvent(originalCal, false, eventId, row);
                            if (evToDel) evToDel.deleteEvent();
                            updates['Log'] = '🗑️ Removed';
                            break;
                    }
                    App.Data.patchRow('CALENDAR_SYNC', row._rowNumber, updates);
                    processed++;
                } catch (e) {
                    App.Data.patchRow('CALENDAR_SYNC', row._rowNumber, { 'Log': '❌ ' + e.message });
                    errors++;
                }
            });

            return _App_ok("Sync complete. Success: " + processed + ", Errors: " + errors);
        }
    }
});

/**
 * Entry point for menu (idempotent setup)
 */
function Calendar_showSidebar() {
    App.UI.launchTool('CALENDAR_SYNC');
}

/**
 * Internal helper for event finding (Shared logic)
 */
function _CalendarSync_findEvent(cal, allowGlobal, eventId, eventData) {
  var event = null;
  if (cal) { try { event = cal.getEventById(eventId); } catch(e){} }
  if (!event && allowGlobal) { try { event = CalendarApp.getEventById(eventId); } catch(e){} }
  
  // Minimal fallback if ID fails
  if (!event && cal && eventData.start && eventData.end) {
      var events = cal.getEvents(new Date(new Date(eventData.start).getTime() - 1000), new Date(new Date(eventData.end).getTime() + 1000));
      event = events.find(function(e) { return e.getTitle() === eventData.title; });
  }
  return event;
}

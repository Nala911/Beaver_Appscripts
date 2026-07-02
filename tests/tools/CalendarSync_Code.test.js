// Tests for Google Calendar Sync tool

describe('CalendarSync Tool', () => {
  const SHEET_NAME = '🗓️ Google Calendar';
  let sheet;

  beforeEach(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.insertSheet(SHEET_NAME);
  });

  test('should pull calendar events and populate sheet', () => {
    // 1. Setup mock calendar and events
    const mockEvent1 = {
      getTitle: () => 'Design Review',
      getStartTime: () => new Date('2026-10-12T10:00:00Z'),
      getEndTime: () => new Date('2026-10-12T11:00:00Z'),
      getDescription: () => 'Go over architecture designs',
      getLocation: () => 'Room 404',
      getGuestList: () => [
        { getEmail: () => 'john@example.com' },
        { getEmail: () => 'alice@example.com' }
      ],
      getId: () => 'event-123'
    };

    const mockCal = {
      getName: () => 'Personal Calendar',
      getId: () => 'personal@example.com',
      getEvents: jest.fn(() => [mockEvent1])
    };

    CalendarApp.getAllCalendars.mockReturnValue([mockCal]);

    // 2. Act
    const result = SyncEngine.runAction('CALENDAR_SYNC', 'pull', [{
      startDate: '2026-10-01',
      endDate: '2026-10-31'
    }]);

    // 3. Assert
    expect(result.success).toBe(true);
    expect(result.message).toContain('imported 1 events');
    expect(mockCal.getEvents).toHaveBeenCalledWith(
      new Date('2026-10-01'),
      new Date('2026-10-31')
    );

    const readObjects = SheetManager.readObjects('CALENDAR_SYNC');
    expect(readObjects.length).toBe(1);
    expect(readObjects[0]['Calendar Name']).toBe('Personal Calendar');
    expect(readObjects[0]['Event Title']).toBe('Design Review');
    expect(readObjects[0]['Start Time']).toBe('10/12/2026 10:00:00');
    expect(readObjects[0]['End Time']).toBe('10/12/2026 11:00:00');
    expect(readObjects[0]['Description']).toBe('Go over architecture designs');
    expect(readObjects[0]['Location']).toBe('Room 404');
    expect(readObjects[0]['Guests']).toBe('john@example.com,alice@example.com');
    expect(readObjects[0]['Event ID']).toBe('event-123');
    expect(readObjects[0]['Calendar ID']).toBe('personal@example.com');
  });

  test('should pull calendar events using global wrapper and fallbacks', () => {
    // 1. Setup mock calendar and events
    const mockEvent = {
      getTitle: () => 'Team Standup',
      getStartTime: () => new Date('2026-10-15T09:00:00Z'),
      getEndTime: () => new Date('2026-10-15T09:30:00Z'),
      getDescription: () => 'Daily sync',
      getLocation: () => 'Virtual',
      getGuestList: () => [],
      getId: () => 'event-456'
    };

    const mockCal = {
      getName: () => 'Personal Calendar',
      getId: () => 'personal@example.com',
      getEvents: jest.fn(() => [mockEvent])
    };

    CalendarApp.getAllCalendars.mockReturnValue([mockCal]);

    // Set properties using our mock APP_PROPS/properties store
    _App_setProperty(APP_PROPS.CAL_START_DATE, '2026-10-10');
    _App_setProperty(APP_PROPS.CAL_END_DATE, '2026-10-20');

    // 2. Act - call global wrapper function with no arguments
    const result = CalendarSync_pullEvents();

    // 3. Assert
    expect(result.success).toBe(true);
    expect(mockCal.getEvents).toHaveBeenCalledWith(
      new Date('2026-10-10'),
      new Date('2026-10-20')
    );

    const readObjects = SheetManager.readObjects('CALENDAR_SYNC');
    expect(readObjects.length).toBe(1);
    expect(readObjects[0]['Event Title']).toBe('Team Standup');
  });
});


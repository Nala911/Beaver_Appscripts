// Tests for central SheetManager (DAO Pattern)

describe('SheetManager', () => {
  const TOOL_KEY = 'TASKS_SYNC';
  const SHEET_NAME = '📋 Google Tasks';
  let sheet;

  beforeEach(() => {
    // SpreadsheetApp is globally mocked
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.insertSheet(SHEET_NAME);
  });

  test('should read empty sheet as empty array', () => {
    const data = SheetManager.readObjects(TOOL_KEY);
    expect(data).toEqual([]);
  });

  test('should write objects and read them back', () => {
    const headers = SheetManager.getHeaders(TOOL_KEY);
    
    // Set headers on the first row of the sheet
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const itemsToWrite = [
      {
        'Action': 'CREATE',
        'Status': '',
        'Task List Name': 'Personal',
        'Task Title': 'Buy groceries',
        'Description': 'Milk and eggs',
        'Due Date': '10/12/2026',
        'Completed?': 'Not Completed',
        'Task ID': '',
        'Task List ID': ''
      },
      {
        'Action': 'UPDATE',
        'Status': '',
        'Task List Name': 'Work',
        'Task Title': 'Finish reports',
        'Description': 'Q3 report completion',
        'Due Date': '10/15/2026',
        'Completed?': 'Not Completed',
        'Task ID': 't-1234',
        'Task List ID': 'l-5678'
      }
    ];

    SheetManager.writeObjects(TOOL_KEY, itemsToWrite);

    const readItems = SheetManager.readObjects(TOOL_KEY);
    
    expect(readItems.length).toBe(2);
    expect(readItems[0]['Task Title']).toBe('Buy groceries');
    expect(readItems[0]['Action']).toBe('CREATE');
    expect(readItems[1]['Task ID']).toBe('t-1234');
  });

  test('should read only pending items with actions set', () => {
    const headers = SheetManager.getHeaders(TOOL_KEY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const items = [
      { 'Action': '', 'Status': 'Synced', 'Task Title': 'Task 1' },
      { 'Action': 'CREATE', 'Status': '', 'Task Title': 'Task 2' },
      { 'Action': '', 'Status': 'Synced', 'Task Title': 'Task 3' },
      { 'Action': 'DELETE', 'Status': '', 'Task Title': 'Task 4' }
    ];
    SheetManager.writeObjects(TOOL_KEY, items);

    const pending = SheetManager.readPendingObjects(TOOL_KEY);
    expect(pending.length).toBe(2);
    expect(pending[0]['Task Title']).toBe('Task 2');
    expect(pending[0]._rowNumber).toBe(3); // row 1 headers, row 2 is task 1, row 3 is task 2
    expect(pending[1]['Task Title']).toBe('Task 4');
    expect(pending[1]._rowNumber).toBe(5);
  });

  test('should clear sheet data starting from row 2', () => {
    const headers = SheetManager.getHeaders(TOOL_KEY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const items = [
      { 'Action': 'CREATE', 'Task Title': 'Task 1' }
    ];
    SheetManager.writeObjects(TOOL_KEY, items);
    expect(sheet.getLastRow()).toBe(2);

    SheetManager.clearData(TOOL_KEY);
    expect(sheet.getLastRow()).toBe(1); // Only headers left
  });

  test('should patch a specific row', () => {
    const headers = SheetManager.getHeaders(TOOL_KEY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const items = [
      { 'Action': 'CREATE', 'Task Title': 'Original Title', 'Status': '' }
    ];
    SheetManager.writeObjects(TOOL_KEY, items);

    SheetManager.patchRow(TOOL_KEY, 2, {
      'Task Title': 'Patched Title',
      'Status': '✅ Synced'
    });

    const readBack = SheetManager.readObjects(TOOL_KEY);
    expect(readBack[0]['Task Title']).toBe('Patched Title');
    expect(readBack[0]['Status']).toBe('✅ Synced');
  });

  test('should optimize read and batch patch for highly fragmented pending actions', () => {
    const headers = SheetManager.getHeaders(TOOL_KEY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Create 30 rows where only indices 0, 7, 14, 21, 28 are pending
    const items = [];
    for (let r = 0; r < 30; r++) {
      if (r === 0 || r === 7 || r === 14 || r === 21 || r === 28) {
        items.push({ 'Action': 'CREATE', 'Status': '', 'Task Title': `Task ${r}` });
      } else {
        items.push({ 'Action': '', 'Status': 'Synced', 'Task Title': `Task ${r}` });
      }
    }
    SheetManager.writeObjects(TOOL_KEY, items);

    // 1. Verify readPendingObjects parses them correctly using full-sheet optimization
    const pending = SheetManager.readPendingObjects(TOOL_KEY);
    expect(pending.length).toBe(5);
    expect(pending[0]['Task Title']).toBe('Task 0');
    expect(pending[0]._rowNumber).toBe(2);
    expect(pending[1]['Task Title']).toBe('Task 7');
    expect(pending[1]._rowNumber).toBe(9);
    expect(pending[4]['Task Title']).toBe('Task 28');
    expect(pending[4]._rowNumber).toBe(30);

    // 2. Verify batchPatchRows works correctly using full-sheet write optimization
    const rowNumbers = pending.map(p => p._rowNumber);
    const updates = pending.map((p, idx) => ({
      'Action': '',
      'Status': `✅ Done ${idx}`
    }));

    SheetManager.batchPatchRows(TOOL_KEY, rowNumbers, updates);

    const allObjects = SheetManager.readObjects(TOOL_KEY);
    expect(allObjects[0]['Action']).toBe('');
    expect(allObjects[0]['Status']).toBe('✅ Done 0');
    expect(allObjects[7]['Action']).toBe('');
    expect(allObjects[7]['Status']).toBe('✅ Done 1');
    expect(allObjects[28]['Action']).toBe('');
    expect(allObjects[28]['Status']).toBe('✅ Done 4');

    // Non-pending should remain untouched
    expect(allObjects[1]['Action']).toBe('');
    expect(allObjects[1]['Status']).toBe('Synced');
  });
});

// Tests for Google Tasks Sync tool

describe('TasksSync Tool', () => {
  const SHEET_NAME = SHEET_NAMES.TASKS_SYNC;
  let sheet;

  beforeEach(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.insertSheet(SHEET_NAME);
  });

  test('should pull tasks and populate sheet', () => {
    // 1. Arrange mock task lists and tasks
    Tasks.Tasklists.list.mockReturnValue({
      items: [
        { id: 'list-1', title: 'Work Tasks' }
      ]
    });
    
    Tasks.Tasks.list.mockReturnValue({
      items: [
        { id: 'task-1', title: 'Task A', notes: 'Desc A', due: '2026-10-12T00:00:00.000Z', status: 'needsAction' },
        { id: 'task-2', title: 'Task B', notes: 'Desc B', due: '', status: 'completed' }
      ]
    });

    // 2. Act
    const result = SyncEngine.runAction('TASKS_SYNC', 'pull');

    // 3. Assert
    expect(result.success).toBe(true);
    expect(result.message).toContain('imported 2 tasks');

    const readObjects = SheetManager.readObjects('TASKS_SYNC');
    expect(readObjects.length).toBe(2);
    
    // Check Task A details
    expect(readObjects[0]['Task Title']).toBe('Task A');
    expect(readObjects[0]['Task List Name']).toBe('Work Tasks');
    expect(readObjects[0]['Due Date']).toBe('10/12/2026');
    expect(readObjects[0]['Completed?']).toBe('Not Completed');
    expect(readObjects[0]['Task ID']).toBe('task-1');

    // Check Task B details
    expect(readObjects[1]['Task Title']).toBe('Task B');
    expect(readObjects[1]['Completed?']).toBe('Completed');
  });

  test('should push CREATE changes successfully', () => {
    const headers = SheetManager.getHeaders('TASKS_SYNC');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Pre-populate mock task lists
    Tasks.Tasklists.list.mockReturnValue({
      items: [{ id: 'list-1', title: 'Work Tasks' }]
    });

    // Write a pending CREATE row
    const pendingItem = {
      'Action': 'CREATE',
      'Status': '',
      'Task List Name': 'Work Tasks',
      'Task Title': 'New Groceries',
      'Description': 'Buy milk',
      'Due Date': '10/20/2026',
      'Completed?': 'Not Completed'
    };
    SheetManager.writeObjects('TASKS_SYNC', [pendingItem]);

    // Setup mock response for Tasks.Tasks.insert
    Tasks.Tasks.insert.mockReturnValue({ id: 'inserted-task-999' });

    // Act
    const pushResult = SyncEngine.runAction('TASKS_SYNC', 'push');

    // Assert
    expect(pushResult.success).toBe(true);
    expect(Tasks.Tasks.insert).toHaveBeenCalledWith(
      expect.objectContaining({
        title: 'New Groceries',
        notes: 'Buy milk',
        status: 'needsAction'
      }),
      'list-1'
    );

    const updatedRows = SheetManager.readObjects('TASKS_SYNC');
    expect(updatedRows[0]['Action']).toBe(''); // Cleared
    expect(updatedRows[0]['Status']).toContain('Created'); // Success status
    expect(updatedRows[0]['Task ID']).toBe('inserted-task-999');
    expect(updatedRows[0]['Task List ID']).toBe('list-1');
  });

  test('should push UPDATE changes successfully', () => {
    const headers = SheetManager.getHeaders('TASKS_SYNC');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Pre-populate mock task lists
    Tasks.Tasklists.list.mockReturnValue({
      items: [{ id: 'list-1', title: 'Work Tasks' }]
    });

    const pendingItem = {
      'Action': 'UPDATE',
      'Status': '',
      'Task List Name': 'Work Tasks',
      'Task Title': 'Updated Task Name',
      'Description': 'Updated Desc',
      'Completed?': 'Completed',
      'Task ID': 'task-777',
      'Task List ID': 'list-1'
    };
    SheetManager.writeObjects('TASKS_SYNC', [pendingItem]);

    // Act
    const pushResult = SyncEngine.runAction('TASKS_SYNC', 'push');

    // Assert
    expect(pushResult.success).toBe(true);
    expect(Tasks.Tasks.patch).toHaveBeenCalledWith(
      expect.objectContaining({
        title: 'Updated Task Name',
        notes: 'Updated Desc',
        status: 'completed'
      }),
      'list-1',
      'task-777'
    );

    const updatedRows = SheetManager.readObjects('TASKS_SYNC');
    expect(updatedRows[0]['Action']).toBe('');
    expect(updatedRows[0]['Status']).toContain('Updated');
  });

  test('should push DELETE changes successfully', () => {
    const headers = SheetManager.getHeaders('TASKS_SYNC');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const pendingItem = {
      'Action': 'DELETE',
      'Status': '',
      'Task ID': 'task-delete-888',
      'Task List ID': 'list-delete-888'
    };
    SheetManager.writeObjects('TASKS_SYNC', [pendingItem]);

    // Act
    const pushResult = SyncEngine.runAction('TASKS_SYNC', 'push');

    // Assert
    expect(pushResult.success).toBe(true);
    expect(Tasks.Tasks.remove).toHaveBeenCalledWith('list-delete-888', 'task-delete-888');

    const updatedRows = SheetManager.readObjects('TASKS_SYNC');
    expect(updatedRows[0]['Action']).toBe('');
    expect(updatedRows[0]['Status']).toContain('Deleted');
  });
});

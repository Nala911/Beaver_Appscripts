// Tests for Google Drive File Details Sync tool
describe('DriveFileDetails Tool', () => {
  const SHEET_NAME = '💾 Google Drive';
  let sheet;

  beforeEach(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.insertSheet(SHEET_NAME);
  });

  test('should push UPDATE changes successfully for PDF file metadata rename', () => {
    const headers = SheetManager.getHeaders('DRIVE_SYNC');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const pendingItem = {
      'Action': 'UPDATE',
      'Status': '',
      'Item Name': 'new_name.pdf',
      'Description': 'My PDF Description',
      'Starred': false,
      'Type': 'PDF',
      'Editors': '',
      'Viewers': '',
      'Is Public?': false,
      'Parent Path': 'My Drive',
      'Item ID': 'file-pdf-123'
    };
    SheetManager.writeObjects('DRIVE_SYNC', [pendingItem]);

    // Mock Drive.Files.get
    Drive.Files.get.mockReturnValue({
      name: 'old_name.pdf',
      description: 'Old Desc',
      starred: false,
      permissions: []
    });

    // Mock Drive.Files.update
    Drive.Files.update.mockReturnValue({});

    // Act
    const pushResult = SyncEngine.runAction('DRIVE_SYNC', 'push');

    // Assert
    expect(pushResult).toContainEqual(expect.stringContaining('Sequence Complete.'));
    expect(Drive.Files.get).toHaveBeenCalledWith('file-pdf-123', expect.any(Object));
    expect(Drive.Files.update).toHaveBeenCalledWith(
      expect.objectContaining({ name: 'new_name.pdf' }),
      'file-pdf-123',
      null,
      expect.any(Object)
    );
  });
});

// Tests for Google Contacts Sync tool

describe('ContactsSync Tool', () => {
  const SHEET_NAME = '☎️ Google Contacts';
  let sheet;

  beforeEach(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.insertSheet(SHEET_NAME);
  });

  test('should pull contacts and populate sheet', () => {
    // 1. Arrange mock groups and connections
    People.ContactGroups = {
      list: jest.fn(() => ({
        contactGroups: [
          { resourceName: 'contactGroups/starred', formattedName: 'Starred' },
          { resourceName: 'contactGroups/friends', formattedName: 'Friends' }
        ]
      }))
    };

    People.People = {
      Connections: {
        list: jest.fn(() => ({
          connections: [
            {
              resourceName: 'people/c111',
              names: [{ givenName: 'John', familyName: 'Doe', metadata: { primary: true } }],
              emailAddresses: [{ value: 'john@example.com', metadata: { primary: true } }],
              phoneNumbers: [{ value: '123456', metadata: { primary: true } }],
              organizations: [{ name: 'Acme Corp', title: 'Developer', metadata: { primary: true } }],
              memberships: [{ contactGroupMembership: { contactGroupResourceName: 'contactGroups/friends' } }],
              biographies: [{ value: 'Test Notes' }],
              addresses: [{ streetAddress: '123 St', city: 'City', region: 'State', postalCode: '54321', metadata: { primary: true } }]
            }
          ]
        }))
      }
    };

    // 2. Act
    const result = SyncEngine.runAction('CONTACTS_SYNC', 'pull', [{ groupIds: ['all'] }]);

    // 3. Assert
    expect(result.success).toBe(true);
    expect(result.message).toContain('Successfully imported 1 contacts');

    const readObjects = SheetManager.readObjects('CONTACTS_SYNC');
    expect(readObjects.length).toBe(1);

    expect(readObjects[0]['First Name']).toBe('John');
    expect(readObjects[0]['Last Name']).toBe('Doe');
    expect(readObjects[0]['Email']).toBe('john@example.com');
    expect(readObjects[0]['Phone']).toBe('123456');
    expect(readObjects[0]['Company']).toBe('Acme Corp');
    expect(readObjects[0]['Job Title']).toBe('Developer');
    expect(readObjects[0]['Starred']).toBe(false);
    expect(readObjects[0]['Street']).toBe('123 St');
    expect(readObjects[0]['City']).toBe('City');
    expect(readObjects[0]['State']).toBe('State');
    expect(readObjects[0]['Zip']).toBe('54321');
    expect(readObjects[0]['Groups/Labels']).toBe('Friends');
    expect(readObjects[0]['Notes']).toBe('Test Notes');
    expect(readObjects[0]['Contact ID']).toBe('people/c111');
  });

  test('should push CREATE changes successfully', () => {
    const headers = SheetManager.getHeaders('CONTACTS_SYNC');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Mock ContactGroups list
    People.ContactGroups = {
      list: jest.fn(() => ({
        contactGroups: [
          { resourceName: 'contactGroups/friends', formattedName: 'Friends' }
        ]
      })),
      Members: {
        modify: jest.fn()
      }
    };

    // Mock createContact
    People.People = {
      createContact: jest.fn(() => ({
        resourceName: 'people/c222'
      }))
    };

    // Write a pending CREATE row
    const pendingItem = {
      'Action': 'CREATE',
      'Status': '',
      'First Name': 'Alice',
      'Last Name': 'Smith',
      'Email': 'alice@example.com',
      'Phone': '987654',
      'Company': 'Gizmos Inc',
      'Job Title': 'Designer',
      'Starred': true,
      'Street': '456 Way',
      'City': 'Town',
      'State': 'Region',
      'Zip': '12345',
      'Groups/Labels': 'Friends',
      'Notes': 'Some design notes'
    };
    SheetManager.writeObjects('CONTACTS_SYNC', [pendingItem]);

    // Act
    const pushResult = SyncEngine.runAction('CONTACTS_SYNC', 'push');

    // Assert
    expect(pushResult.success).toBe(true);
    expect(People.People.createContact).toHaveBeenCalledWith(
      expect.objectContaining({
        names: [{ givenName: 'Alice', familyName: 'Smith' }],
        emailAddresses: [{ value: 'alice@example.com' }],
        phoneNumbers: [{ value: '987654' }]
      })
    );

    expect(People.ContactGroups.Members.modify).toHaveBeenCalledWith(
      { resourceNamesToAdd: ['people/c222'] },
      'contactGroups/friends'
    );

    const updatedRows = SheetManager.readObjects('CONTACTS_SYNC');
    expect(updatedRows[0]['Action']).toBe('');
    expect(updatedRows[0]['Status']).toContain('Created');
    expect(updatedRows[0]['Contact ID']).toBe('people/c222');
  });

  test('should push UPDATE changes successfully', () => {
    const headers = SheetManager.getHeaders('CONTACTS_SYNC');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    People.ContactGroups = {
      list: jest.fn(() => ({
        contactGroups: [
          { resourceName: 'contactGroups/friends', formattedName: 'Friends' }
        ]
      })),
      Members: {
        modify: jest.fn()
      }
    };

    // Mock People get and update
    People.People = {
      get: jest.fn(() => ({
        etag: 'mock-etag',
        names: [{ givenName: 'Alice', familyName: 'Smith' }],
        emailAddresses: [{ value: 'alice-old@example.com', metadata: { primary: true } }],
        phoneNumbers: [{ value: '987654', metadata: { primary: true } }],
        addresses: [{ streetAddress: '456 Way', city: 'Town', region: 'Region', postalCode: '12345', metadata: { primary: true } }]
      })),
      updateContact: jest.fn()
    };

    const pendingItem = {
      'Action': 'UPDATE',
      'Status': '',
      'First Name': 'Alice',
      'Last Name': 'Smith',
      'Email': 'alice-new@example.com',
      'Phone': '987654',
      'Company': 'Gizmos Inc',
      'Job Title': 'Designer',
      'Starred': false,
      'Street': '456 Way',
      'City': 'Town',
      'State': 'Region',
      'Zip': '12345',
      'Groups/Labels': 'Friends',
      'Notes': 'Some design notes',
      'Contact ID': 'people/c222'
    };
    SheetManager.writeObjects('CONTACTS_SYNC', [pendingItem]);

    // Act
    const pushResult = SyncEngine.runAction('CONTACTS_SYNC', 'push');

    // Assert
    expect(pushResult.success).toBe(true);
    expect(People.People.updateContact).toHaveBeenCalledWith(
      expect.objectContaining({
        etag: 'mock-etag',
        emailAddresses: [{ value: 'alice-new@example.com', metadata: { primary: true } }]
      }),
      'people/c222',
      expect.any(Object)
    );

    const updatedRows = SheetManager.readObjects('CONTACTS_SYNC');
    expect(updatedRows[0]['Action']).toBe('');
    expect(updatedRows[0]['Status']).toContain('Updated');
  });

  test('should push DELETE changes successfully', () => {
    const headers = SheetManager.getHeaders('CONTACTS_SYNC');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    People.ContactGroups = {
      list: jest.fn(() => ({
        contactGroups: []
      }))
    };

    People.People = {
      deleteContact: jest.fn()
    };

    const pendingItem = {
      'Action': 'DELETE',
      'Status': '',
      'Contact ID': 'people/c222'
    };
    SheetManager.writeObjects('CONTACTS_SYNC', [pendingItem]);

    // Act
    const pushResult = SyncEngine.runAction('CONTACTS_SYNC', 'push');

    // Assert
    expect(pushResult.success).toBe(true);
    expect(People.People.deleteContact).toHaveBeenCalledWith('people/c222');

    const updatedRows = SheetManager.readObjects('CONTACTS_SYNC');
    expect(updatedRows[0]['Action']).toBe('');
    expect(updatedRows[0]['Status']).toContain('Deleted');
  });
});

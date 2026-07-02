// Tests for Dynamic Proxy Mocking & Real-time Tracing

describe('Dynamic Proxy Mocking', () => {
  test('should successfully trace existing mocked API calls in real-time', () => {
    // Act
    const active = SpreadsheetApp.getActiveSpreadsheet();
    
    // Assert
    expect(active).toBeDefined();
    expect(global._mockCallsTrace.length).toBeGreaterThan(0);
    
    const trace = global._mockCallsTrace[0];
    expect(trace.service).toBe('SpreadsheetApp');
    expect(trace.method).toBe('getActiveSpreadsheet');
    expect(trace.status).toBe('SUCCESS');
  });

  test('should throw a descriptive Missing Mock Error when calling an unmocked method', () => {
    // Act & Assert
    expect(() => {
      SpreadsheetApp.someUnmockedMethod();
    }).toThrow("Missing Mock Error: 'SpreadsheetApp.someUnmockedMethod' was called but is not mocked");
    
    // Verify it was logged in the trace as a failure
    const trace = global._mockCallsTrace.find(t => t.method === 'someUnmockedMethod');
    expect(trace).toBeDefined();
    expect(trace.status).toBe('FAILED');
    expect(trace.error).toContain('someUnmockedMethod');
  });

  test('should auto-load unmocked Advanced services like People and provide clear error messages', () => {
    // Act & Assert
    expect(() => {
      People.ContactGroups.list();
    }).toThrow("Missing Mock Error: 'People.ContactGroups.list' was called but is not mocked");
    
    // Verify it was logged in the trace as a failure
    const trace = global._mockCallsTrace.find(t => t.service === 'People.ContactGroups' && t.method === 'list');
    expect(trace).toBeDefined();
    expect(trace.status).toBe('FAILED');
    expect(trace.error).toContain('People.ContactGroups.list');
  });
});

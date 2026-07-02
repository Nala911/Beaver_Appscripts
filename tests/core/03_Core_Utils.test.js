// Tests for unified core utilities in core/03_Core_Utils.js

describe('Core Utilities', () => {
  describe('_App_extractIdFromUrl', () => {
    test('should extract ID from standard Google Doc URL', () => {
      const url = 'https://docs.google.com/document/d/1A-Bx9_z6X1xY74LhW22H4K-Ww_o1234567890abcdef/edit';
      const id = _App_extractIdFromUrl(url);
      expect(id).toBe('1A-Bx9_z6X1xY74LhW22H4K-Ww_o1234567890abcdef');
    });

    test('should extract ID from Google Form viewform URL', () => {
      const url = 'https://docs.google.com/forms/d/e/1FAIpQLSf2aX1xY74LhW22H4K-Ww_o1234567890abcdef/viewform';
      const id = _App_extractIdFromUrl(url);
      expect(id).toBe('1FAIpQLSf2aX1xY74LhW22H4K-Ww_o1234567890abcdef');
    });

    test('should extract ID from Google Form edit URL', () => {
      const url = 'https://docs.google.com/forms/d/1FAIpQLSf2aX1xY74LhW22H4K-Ww_o1234567890abcdef/edit';
      const id = _App_extractIdFromUrl(url);
      expect(id).toBe('1FAIpQLSf2aX1xY74LhW22H4K-Ww_o1234567890abcdef');
    });

    test('should return raw ID if input is already a valid ID', () => {
      const rawId = '1FAIpQLSf2aX1xY74LhW22H4K-Ww_o1234567890abcdef';
      const id = _App_extractIdFromUrl(rawId);
      expect(id).toBe(rawId);
    });

    test('should return null for empty or invalid input', () => {
      expect(_App_extractIdFromUrl('')).toBeNull();
      expect(_App_extractIdFromUrl(null)).toBeNull();
      expect(_App_extractIdFromUrl(undefined)).toBeNull();
      expect(_App_extractIdFromUrl('https://google.com')).toBeNull();
    });
  });

  describe('_App_fetchParallel', () => {
    test('should execute fetchAll mock successfully and return results', () => {
      const requests = [
        { url: 'https://api.example.com/1', method: 'get' },
        { url: 'https://api.example.com/2', method: 'post', payload: 'data' }
      ];

      const responses = _App_fetchParallel(requests);
      
      expect(UrlFetchApp.fetchAll).toHaveBeenCalledWith(requests);
      expect(responses.length).toBe(2);
      expect(responses[0].getResponseCode()).toBe(200);
      expect(responses[0].getContentText()).toBe('{}');
    });

    test('should return empty array for empty requests', () => {
      expect(_App_fetchParallel([])).toEqual([]);
      expect(_App_fetchParallel(null)).toEqual([]);
    });
  });

  describe('SyncEngine.Utils Namespace Integration', () => {
    test('should correctly expose all core utilities via SyncEngine.Utils', () => {
      expect(SyncEngine.Utils).toBeDefined();
      expect(SyncEngine.Utils.ok).toBe(_App_ok);
      expect(SyncEngine.Utils.fail).toBe(_App_fail);
      expect(SyncEngine.Utils.withDocumentLock).toBe(_App_withDocumentLock);
      expect(SyncEngine.Utils.include).toBe(_App_include);
      expect(SyncEngine.Utils.createTemplateFromFile).toBe(_App_createTemplateFromFile);
      expect(SyncEngine.Utils.formatStatus).toBe(_App_formatStatus);
      expect(SyncEngine.Utils.throttle).toBe(_App_throttle);
      expect(SyncEngine.Utils.callWithBackoff).toBe(_App_callWithBackoff);
      expect(SyncEngine.Utils.isExecutionLimitApproaching).toBe(_App_isExecutionLimitApproaching);
      expect(SyncEngine.Utils.translateApiError).toBe(_App_translateApiError);
      expect(SyncEngine.Utils.validateRowAgainstSchema).toBe(_App_validateRowAgainstSchema);
      expect(SyncEngine.Utils.BatchProcessor).toBe(_App_BatchProcessor);
      expect(SyncEngine.Utils.batchPatchResults).toBe(_App_batchPatchResults);
      expect(SyncEngine.Utils.extractIdFromUrl).toBe(_App_extractIdFromUrl);
      expect(SyncEngine.Utils.fetchParallel).toBe(_App_fetchParallel);
      expect(SyncEngine.Utils.formatDateTime).toBe(_App_formatDateTime);
    });
  });
});

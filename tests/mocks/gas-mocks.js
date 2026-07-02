// Google Apps Script Global API Mocks for Local Testing

function _createGasMockProxy(serviceName, target) {
  if (target && target._isMockFunction) {
    return target;
  }
  
  var isDummy = false;
  var baseObject;
  if (typeof target === 'function') {
    baseObject = target;
  } else {
    baseObject = function() {
      throw new Error("Missing Mock Error: '" + serviceName + "' was called but is not mocked in tests/mocks/gas-mocks.js. Please define its return value in your test.");
    };
    isDummy = true;
  }
  
  if (target && typeof target === 'object') {
    Object.assign(baseObject, target);
  }

  return new Proxy(baseObject, {
    apply: function(targetFn, thisArg, argumentsList) {
      var traceEntry = {
        timestamp: new Date().toISOString(),
        service: serviceName.includes('.') ? serviceName.substring(0, serviceName.lastIndexOf('.')) : serviceName,
        method: serviceName.includes('.') ? serviceName.substring(serviceName.lastIndexOf('.') + 1) : '',
        arguments: argumentsList,
        status: 'SUCCESS',
        returnValue: undefined
      };

      if (global._mockCallsTrace) global._mockCallsTrace.push(traceEntry);

      try {
        var ret = targetFn.apply(thisArg, argumentsList);
        traceEntry.returnValue = ret;
        return ret;
      } catch (e) {
        traceEntry.status = 'FAILED';
        traceEntry.error = e.message || String(e);
        throw e;
      }
    },

    get: function(obj, prop) {
      if (typeof prop === 'symbol' || prop === 'then' || prop === 'asymmetricMatch' || prop === 'calls' || prop === 'mock' || prop === '_isMockFunction' || prop === 'prototype' || prop === 'mockClear' || prop === 'mockReset' || prop === 'mockRestore') {
        return obj[prop];
      }
      var propStr = String(prop);
      if (propStr.startsWith('_') || propStr === 'toJSON') {
        return obj[prop];
      }

      if (obj.hasOwnProperty(prop)) {
        var val = obj[prop];
        return _createGasMockProxy(serviceName + '.' + propStr, val);
      }

      var methodPath = serviceName + '.' + propStr;
      return _createGasMockProxy(methodPath, undefined);
    }
  });
}


// 1. Properties & Cache Services Mocks
class MockStore {
  constructor() {
    this.store = {};
  }
  getProperty(key) {
    return this.store.hasOwnProperty(key) ? this.store[key] : null;
  }
  setProperty(key, val) {
    this.store[key] = String(val);
  }
  deleteProperty(key) {
    delete this.store[key];
  }
}

class MockCache {
  constructor() {
    this.cache = {};
  }
  get(key) {
    return this.cache.hasOwnProperty(key) ? this.cache[key] : null;
  }
  put(key, value, ttl) {
    this.cache[key] = String(value);
  }
  remove(key) {
    delete this.cache[key];
  }
}

const documentProperties = new MockStore();
const userProperties = new MockStore();
const scriptProperties = new MockStore();

const documentCache = new MockCache();
const userCache = new MockCache();
const scriptCache = new MockCache();

global.PropertiesService = {
  getDocumentProperties: () => documentProperties,
  getUserProperties: () => userProperties,
  getScriptProperties: () => scriptProperties,
};

global.CacheService = {
  getDocumentCache: () => documentCache,
  getUserCache: () => userCache,
  getScriptCache: () => scriptCache,
};

// 2. Lock & Utilities Services Mocks
const mockLock = {
  tryLock: (timeout) => true,
  releaseLock: () => {},
};

global.LockService = {
  getDocumentLock: () => mockLock,
  getUserLock: () => mockLock,
  getScriptLock: () => mockLock,
};

global.Utilities = {
  formatDate: (date, tz, format) => {
    // Simple mock formatting for MM/dd/yyyy
    const pad = (n) => String(n).padStart(2, '0');
    const m = pad(date.getMonth() + 1);
    const d = pad(date.getDate());
    const y = date.getFullYear();
    if (format === 'MM/dd/yyyy') return `${m}/${d}/${y}`;
    if (format === 'MM/dd/yyyy HH:mm:ss') {
      const mU = pad(date.getUTCMonth() + 1);
      const dU = pad(date.getUTCDate());
      const yU = date.getUTCFullYear();
      const hh = pad(date.getUTCHours());
      const mm = pad(date.getUTCMinutes());
      const ss = pad(date.getUTCSeconds());
      return `${mU}/${dU}/${yU} ${hh}:${mm}:${ss}`;
    }
    if (format === 'yyyy-MM-dd HH:mm') {
      const hh = pad(date.getHours());
      const mm = pad(date.getMinutes());
      return `${y}-${m}-${d} ${hh}:${mm}`;
    }
    return date.toISOString();
  },
  sleep: (ms) => {},
};

// 3. HTML Service Mock
global.HtmlService = {
  createHtmlOutputFromFile: (filename) => ({
    getContent: () => `<!-- content of ${filename} -->`,
  }),
};

// 4. SpreadsheetApp Mock
class MockSheet {
  constructor(name) {
    this.name = name;
    this.data = []; // 2D grid representation
    this.frozenRows = 0;
    this.frozenCols = 0;
    this.widths = {};
  }
  
  getLastRow() {
    let lastRow = 0;
    for (let r = 0; r < this.data.length; r++) {
      const row = this.data[r];
      if (row && row.some(val => val !== "" && val !== null && val !== undefined)) {
        lastRow = r + 1;
      }
    }
    return lastRow;
  }
  
  getLastColumn() {
    let lastCol = 0;
    for (let r = 0; r < this.data.length; r++) {
      const row = this.data[r] || [];
      for (let c = 0; c < row.length; c++) {
        if (row[c] !== "" && row[c] !== null && row[c] !== undefined) {
          lastCol = Math.max(lastCol, c + 1);
        }
      }
    }
    return lastCol;
  }

  getName() {
    return this.name;
  }

  setColumnWidth(col, width) {
    this.widths[col] = width;
    return this;
  }

  setRowHeight(row, height) {
    return this;
  }

  setRowHeights(startRow, numRows, height) {
    return this;
  }

  deleteColumns(start, howMany) {
    return this;
  }

  insertColumnsAfter(afterPosition, howMany) {
    return this;
  }

  activate() {
    return this;
  }

  getMaxRows() {
    return Math.max(1000, this.data.length);
  }

  getMaxColumns() {
    return Math.max(26, this.data.length > 0 ? this.data[0].length : 0);
  }
  
  getRange(row, col, numRows, numCols) {
    const sheetInstance = this;
    const rStart = row - 1;
    const cStart = col - 1;
    const rCount = numRows || 1;
    const cCount = numCols || 1;

    return {
      getValues: function() {
        const values = [];
        for (let r = rStart; r < rStart + rCount; r++) {
          const rowVals = [];
          const sheetRow = sheetInstance.data[r] || [];
          for (let c = cStart; c < cStart + cCount; c++) {
            rowVals.push(sheetRow[c] !== undefined ? sheetRow[c] : "");
          }
          values.push(rowVals);
        }
        return values;
      },
      getDisplayValues: function() {
        return this.getValues();
      },
      setValues: function(values) {
        for (let r = 0; r < values.length; r++) {
          const targetRowIdx = rStart + r;
          if (!sheetInstance.data[targetRowIdx]) {
            sheetInstance.data[targetRowIdx] = [];
          }
          for (let c = 0; c < values[r].length; c++) {
            sheetInstance.data[targetRowIdx][cStart + c] = values[r][c];
          }
        }
        return this;
      },
      clearContent: function() {
        for (let r = rStart; r < rStart + rCount; r++) {
          if (sheetInstance.data[r]) {
            for (let c = cStart; c < cStart + cCount; c++) {
              sheetInstance.data[r][c] = "";
            }
          }
        }
        return this;
      },
      setDataValidation: jest.fn().mockReturnThis(),
      setColumnWidth: jest.fn().mockReturnThis(),
      setFontColor: jest.fn().mockReturnThis(),
      setBackground: jest.fn().mockReturnThis(),
      setHorizontalAlignment: jest.fn().mockReturnThis(),
      setVerticalAlignment: jest.fn().mockReturnThis(),
      setWrapStrategy: jest.fn().mockReturnThis(),
      setFontWeight: jest.fn().mockReturnThis(),
      setFontStyle: jest.fn().mockReturnThis(),
      setFontSize: jest.fn().mockReturnThis(),
      setFontFamily: jest.fn().mockReturnThis(),
      setBorder: jest.fn().mockReturnThis(),
      setRowHeight: jest.fn().mockReturnThis(),
      setFontFamilies: jest.fn().mockReturnThis(),
      setFontStyles: jest.fn().mockReturnThis(),
      setBackgrounds: jest.fn().mockReturnThis(),
      setNumberFormats: jest.fn().mockReturnThis(),
      setDataValidations: jest.fn().mockReturnThis(),
    };
  }
  
  setFrozenRows(num) {
    this.frozenRows = num;
  }
  
  setFrozenColumns(num) {
    this.frozenCols = num;
  }
  
  setRowHeight(row, height) {}
  
  clearConditionalFormatRules() {}
  
  setConditionalFormatRules(rules) {}
  
  getConditionalFormatRules() { return []; }
}

class MockSpreadsheet {
  constructor() {
    this.sheets = {};
    this.timezone = "GMT";
    this.id = "mock-spreadsheet-id";
  }
  
  getSheetByName(name) {
    return this.sheets[name] || null;
  }
  
  getActiveSheet() {
    const keys = Object.keys(this.sheets);
    return keys.length > 0 ? this.sheets[keys[0]] : null;
  }
  
  insertSheet(name) {
    const newSheet = new MockSheet(name);
    this.sheets[name] = newSheet;
    return newSheet;
  }
  
  getSpreadsheetTimeZone() {
    return this.timezone;
  }
  
  getId() {
    return this.id;
  }
}

const activeSpreadsheet = new MockSpreadsheet();
const mockUi = {
  createMenu: jest.fn(() => ({
    addItem: jest.fn().mockReturnThis(),
    addToUi: jest.fn(),
  })),
  showSidebar: jest.fn(),
  showModalDialog: jest.fn(),
};

global.SpreadsheetApp = {
  BorderStyle: { SOLID: 'SOLID' },
  WrapStrategy: { CLIP: 'CLIP' },
  getActiveSpreadsheet: () => activeSpreadsheet,
  getUi: () => mockUi,
  newDataValidation: () => ({
    requireValueInList: jest.fn().mockReturnThis(),
    requireCheckbox: jest.fn().mockReturnThis(),
    requireFormulaSatisfied: jest.fn().mockReturnThis(),
    requireDate: jest.fn().mockReturnThis(),
    setAllowInvalid: jest.fn().mockReturnThis(),
    setHelpText: jest.fn().mockReturnThis(),
    build: jest.fn(() => ({ _isMockValidationRule: true })),
  }),
  newConditionalFormatRule: () => ({
    whenTextEqualTo: jest.fn().mockReturnThis(),
    whenFormulaSatisfied: jest.fn().mockReturnThis(),
    setBackground: jest.fn().mockReturnThis(),
    setFontColor: jest.fn().mockReturnThis(),
    setRanges: jest.fn().mockReturnThis(),
    build: jest.fn(() => ({ _isMockConditionalFormatRule: true })),
  }),
};

// Export helper to reset state between tests
global._resetMockSpreadsheet = () => {
  activeSpreadsheet.sheets = {};
  documentProperties.store = {};
  userProperties.store = {};
  scriptProperties.store = {};
  documentCache.cache = {};
  userCache.cache = {};
  scriptCache.cache = {};
  global._mockCallsTrace = [];
};

// 5. Google Workspace Service Mocks (Tasks, Calendar, People, etc.)
global.Tasks = {
  Tasklists: {
    list: jest.fn(() => ({ items: [] })),
  },
  Tasks: {
    list: jest.fn(() => ({ items: [] })),
    insert: jest.fn((task, listId) => ({ id: 'mock-task-id', ...task })),
    patch: jest.fn((task, listId, taskId) => ({ id: taskId, ...task })),
    remove: jest.fn((listId, taskId) => {}),
  },
};

global.Calendar = {
  CalendarList: {
    list: jest.fn(() => ({ items: [] })),
  },
  Events: {
    list: jest.fn(() => ({ items: [] })),
    insert: jest.fn((event, calendarId) => ({ id: 'mock-event-id', ...event })),
    patch: jest.fn((event, calendarId, eventId) => ({ id: eventId, ...event })),
    remove: jest.fn((calendarId, eventId) => {}),
  },
};

global.CalendarApp = {
  EventColor: {
    BLUE: 'BLUE',
    GREEN: 'GREEN',
    PALE_BLUE: 'PALE_BLUE',
    PALE_GREEN: 'PALE_GREEN',
    MAUVE: 'MAUVE',
    PALE_RED: 'PALE_RED',
    ORANGE: 'ORANGE',
    YELLOW: 'YELLOW',
    GRAY: 'GRAY',
    PALE_YELLOW: 'PALE_YELLOW',
  },
  Visibility: {
    PUBLIC: 'PUBLIC',
    PRIVATE: 'PRIVATE',
  },
  getAllCalendars: jest.fn(() => []),
  getEventById: jest.fn(() => ({
    deleteEvent: jest.fn(),
    setColor: jest.fn(),
    setVisibility: jest.fn(),
  })),
};

global.Drive = {
  Files: {
    get: jest.fn(() => ({})),
    list: jest.fn(() => ({ files: [] })),
    create: jest.fn(() => ({ id: 'mock-file-id' })),
    update: jest.fn(() => ({ id: 'mock-file-id' })),
    remove: jest.fn(() => {}),
  },
  Drives: {
    list: jest.fn(() => ({ drives: [] })),
    get: jest.fn(() => ({ id: 'mock-drive-id' })),
  },
  Permissions: {
    create: jest.fn(() => ({ id: 'mock-perm-id' })),
    update: jest.fn(() => ({ id: 'mock-perm-id' })),
    remove: jest.fn(() => {}),
  }
};

global.DriveApp = {
  getRootFolder: jest.fn(() => ({
    getId: () => 'mock-root-folder-id',
    getName: () => 'My Drive',
  })),
  getFolderById: jest.fn(() => ({
    getId: () => 'mock-folder-id',
    getName: () => 'Mock Folder',
  })),
  getFileById: jest.fn(() => ({
    getId: () => 'mock-file-id',
    getName: () => 'Mock File',
    makeCopy: jest.fn(() => ({
      getId: () => 'mock-copied-file-id',
    })),
  })),
  searchFolders: jest.fn(() => ({
    hasNext: () => false,
    next: () => null,
  })),
  searchFiles: jest.fn(() => ({
    hasNext: () => false,
    next: () => null,
  })),
};

global.DocumentApp = {
  ElementType: {
    PARAGRAPH: 'PARAGRAPH',
    TABLE: 'TABLE',
    LIST_ITEM: 'LIST_ITEM',
  },
  openById: jest.fn(() => ({
    getBody: () => ({
      getText: () => 'mock body text',
      appendParagraph: jest.fn(),
      appendTable: jest.fn(),
      appendListItem: jest.fn(),
    }),
  })),
};

global.FormApp = {
  PageNavigationType: {
    GO_TO_PAGE: 'GO_TO_PAGE',
  },
  openById: jest.fn(() => ({
    getId: () => 'mock-form-id',
  })),
};

global.MailApp = {
  getRemainingDailyQuota: jest.fn(() => 100),
};

global.GmailApp = {
  getDrafts: jest.fn(() => []),
  getDraft: jest.fn(() => ({
    getId: () => 'mock-draft-id',
  })),
  getThreadById: jest.fn(() => ({
    getId: () => 'mock-thread-id',
  })),
  search: jest.fn(() => []),
  sendEmail: jest.fn(),
  createDraft: jest.fn(),
};

global.UrlFetchApp = {
  fetch: jest.fn(() => ({
    getContentText: () => '{}',
    getResponseCode: () => 200,
  })),
  fetchAll: jest.fn((requests) => (requests || []).map(() => ({
    getContentText: () => '{}',
    getResponseCode: () => 200,
  }))),
};

global.Session = {
  getScriptTimeZone: jest.fn(() => 'GMT'),
  getActiveUser: jest.fn(() => ({
    getEmail: () => 'mock-user@example.com',
  })),
};

// Wrap all exposed mock APIs in dynamic proxies for error tracking and call tracing
global.PropertiesService = _createGasMockProxy('PropertiesService', global.PropertiesService);
global.CacheService = _createGasMockProxy('CacheService', global.CacheService);
global.LockService = _createGasMockProxy('LockService', global.LockService);
global.Utilities = _createGasMockProxy('Utilities', global.Utilities);
global.HtmlService = _createGasMockProxy('HtmlService', global.HtmlService);
global.SpreadsheetApp = _createGasMockProxy('SpreadsheetApp', global.SpreadsheetApp);
global.Tasks = _createGasMockProxy('Tasks', global.Tasks);
global.Calendar = _createGasMockProxy('Calendar', global.Calendar);
global.CalendarApp = _createGasMockProxy('CalendarApp', global.CalendarApp);
global.DriveApp = _createGasMockProxy('DriveApp', global.DriveApp);
global.Drive = _createGasMockProxy('Drive', global.Drive);
global.DocumentApp = _createGasMockProxy('DocumentApp', global.DocumentApp);
global.FormApp = _createGasMockProxy('FormApp', global.FormApp);
global.MailApp = _createGasMockProxy('MailApp', global.MailApp);
global.GmailApp = _createGasMockProxy('GmailApp', global.GmailApp);
global.UrlFetchApp = _createGasMockProxy('UrlFetchApp', global.UrlFetchApp);
global.Session = _createGasMockProxy('Session', global.Session);

// Dynamic autoloader for Advanced Google services to prevent cryptic ReferenceErrors
var knownServices = ['People', 'Maps', 'XmlService', 'LanguageApp', 'ContactsApp', 'AdminDirectory', 'AdminReports', 'Gmail', 'Sheets'];
knownServices.forEach(function(service) {
  Object.defineProperty(global, service, {
    get: function() {
      if (!global['_' + service + '_mock_']) {
        global['_' + service + '_mock_'] = _createGasMockProxy(service, {});
      }
      return global['_' + service + '_mock_'];
    },
    configurable: true
  });
});




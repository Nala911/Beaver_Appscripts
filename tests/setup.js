const fs = require('fs');
const path = require('path');
const vm = require('vm');

// 1. Load the mock Google Apps Script context
require('./mocks/gas-mocks');

// 2. Helper to load files into the global context
function loadGlobalFile(filePath) {
  const code = fs.readFileSync(filePath, 'utf8');
  try {
    // Run inside the Jest global context so it shares Jest's global object and mocks
    vm.runInNewContext(code, global, { filename: filePath });
  } catch (err) {
    console.error(`Error loading file: ${filePath}`);
    throw err;
  }
}

// 3. Load core files in alphabetical order
const coreDir = path.join(__dirname, '../core');
if (fs.existsSync(coreDir)) {
  const files = fs.readdirSync(coreDir)
    .filter(f => f.endsWith('.js'))
    .sort();
  files.forEach(f => {
    loadGlobalFile(path.join(coreDir, f));
  });
}

// 4. Load tools files in alphabetical order (including subdirectories recursively)
const toolsDir = path.join(__dirname, '../tools');
function getJsFilesRecursive(dir) {
  let results = [];
  if (!fs.existsSync(dir)) return results;
  const list = fs.readdirSync(dir);
  list.forEach(file => {
    const fullPath = path.join(dir, file);
    const stat = fs.statSync(fullPath);
    if (stat && stat.isDirectory()) {
      results = results.concat(getJsFilesRecursive(fullPath));
    } else if (file.endsWith('.js')) {
      results.push(fullPath);
    }
  });
  return results;
}

if (fs.existsSync(toolsDir)) {
  const files = getJsFilesRecursive(toolsDir).sort();
  files.forEach(f => {
    loadGlobalFile(f);
  });
}

// 5. Automatic state reset before each test
// 5. Automatic state reset before each test
beforeEach(() => {
  global._mockCallsTrace = [];
  if (global._resetMockSpreadsheet) {
    global._resetMockSpreadsheet();
  }
  jest.clearAllMocks();
});

afterEach(() => {
  try {
    const testName = expect.getState().currentTestName;
    const traces = global._mockCallsTrace || [];
    if (traces.length > 0) {
      const traceDir = path.join(__dirname, 'reporters');
      if (!fs.existsSync(traceDir)) {
        fs.mkdirSync(traceDir, { recursive: true });
      }
      const cleanName = testName.replace(/[^a-zA-Z0-9_-]/g, '_');
      const traceFile = path.join(traceDir, `trace-${cleanName}.json`);
      fs.writeFileSync(traceFile, JSON.stringify(traces, null, 2), 'utf8');
    }
  } catch (e) {
    console.error('Failed to save mock calls trace:', e);
  }
});

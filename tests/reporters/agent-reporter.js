const fs = require('fs');
const path = require('path');

class AgentReporter {
  constructor(globalConfig, options) {
    this._globalConfig = globalConfig;
    this._options = options;
    this.failures = [];
  }

  onRunStart() {
    this.failures = [];
    // Clean up old failure report JSON if it exists
    const failureJsonPath = path.join(__dirname, 'failures.json');
    if (fs.existsSync(failureJsonPath)) {
      try {
        fs.unlinkSync(failureJsonPath);
      } catch (e) {}
    }
    // Clean up old trace files
    try {
      const files = fs.readdirSync(__dirname);
      files.forEach(file => {
        if (file.startsWith('trace-') && file.endsWith('.json')) {
          fs.unlinkSync(path.join(__dirname, file));
        }
      });
    } catch (e) {}
  }

  onTestResult(test, testResult, aggregatedResult) {
    if (testResult.numFailingTests === 0) return;

    testResult.testResults.forEach(result => {
      if (result.status !== 'failed') return;

      const failureEntry = {
        testFilePath: testResult.testFilePath,
        title: result.title,
        fullName: result.fullName,
        ancestorTitles: result.ancestorTitles,
        failureMessages: result.failureMessages,
        mockTrace: [],
      };

      // Capture mock trace logs
      const cleanName = result.fullName.replace(/[^a-zA-Z0-9_-]/g, '_');
      const traceFile = path.join(__dirname, `trace-${cleanName}.json`);
      if (fs.existsSync(traceFile)) {
        try {
          failureEntry.mockTrace = JSON.parse(fs.readFileSync(traceFile, 'utf8'));
          fs.unlinkSync(traceFile);
        } catch (e) {}
      }

      // Try to extract exact line details from stack trace
      const stack = result.failureMessages.join('\n');
      const cleanStack = stack.replace(/\x1B\[\d+m/g, '');
      const missingMockMatch = cleanStack.match(/Missing Mock Error:\s+'([^']+)'\s+was\s+called/);
      if (missingMockMatch) {
        const unmockedPath = missingMockMatch[1];
        failureEntry.unmockedMethod = unmockedPath;
        
        const parts = unmockedPath.split('.');
        let suggestedCode = '';
        if (parts.length === 1) {
          suggestedCode = `global.${parts[0]} = jest.fn(() => ({ /* return value */ }));`;
        } else if (parts.length === 2) {
          suggestedCode = `global.${parts[0]} = {\n  ${parts[1]}: jest.fn(() => ({ /* return value */ }))\n};`;
        } else if (parts.length === 3) {
          suggestedCode = `global.${parts[0]} = {\n  ${parts[1]}: {\n    ${parts[2]}: jest.fn(() => ({ /* return value */ }))\n  }\n};`;
        } else {
          suggestedCode = `// Define unmocked path recursively:\n// ${unmockedPath}`;
        }
        failureEntry.suggestedMockScaffolding = suggestedCode;
      }

      const match = stack.match(/at\s+(?:[^\s\(]+\s+\()?\(?([^:\)]+):(\d+):(\d+)\)?/);
      if (match) {
        const filePath = match[1];
        const lineNum = parseInt(match[2], 10);
        
        failureEntry.failureLocation = {
          file: filePath,
          line: lineNum,
        };

        // Attempt to extract source code excerpt
        if (fs.existsSync(filePath)) {
          try {
            const content = fs.readFileSync(filePath, 'utf8');
            const lines = content.split('\n');
            const start = Math.max(0, lineNum - 3);
            const end = Math.min(lines.length, lineNum + 3);
            
            failureEntry.codeExcerpt = lines.slice(start, end).map((line, idx) => {
              const currentLineNum = start + idx + 1;
              const isTarget = currentLineNum === lineNum;
              return `${isTarget ? ' > ' : '   '}${currentLineNum}: ${line}`;
            }).join('\n');
          } catch (e) {
            failureEntry.codeExcerpt = `Could not load source file excerpt: ${e.message}`;
          }
        }
      }

      this.failures.push(failureEntry);
    });
  }

  onRunComplete(contexts, results) {
    if (this.failures.length > 0) {
      const failureJsonPath = path.join(__dirname, 'failures.json');
      fs.writeFileSync(failureJsonPath, JSON.stringify(this.failures, null, 2), 'utf8');
    }
  }
}

module.exports = AgentReporter;

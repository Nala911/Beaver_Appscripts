const fs = require('fs');
const path = require('path');

class AgentReporter {
  constructor(globalConfig, options) {
    this._globalConfig = globalConfig;
    this._options = options;
    this.failures = [];

    // Parse artifact directory from process.argv
    const args = process.argv.slice(2);
    const dirIdx = args.indexOf('--artifact-dir');
    this.artifactDir = (dirIdx !== -1 && args[dirIdx + 1]) ? args[dirIdx + 1] : '';

    this.localReportPath = path.join(__dirname, 'agent_failure_report.md');
    this.localJsonReportPath = path.join(__dirname, 'agent_failure_report.json');
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

    // Clean up old MD/JSON reports from workspace
    if (fs.existsSync(this.localReportPath)) {
      try {
        fs.unlinkSync(this.localReportPath);
      } catch (e) {}
    }
    if (fs.existsSync(this.localJsonReportPath)) {
      try {
        fs.unlinkSync(this.localJsonReportPath);
      } catch (e) {}
    }

    // Clean up old reports in artifact directory if provided
    if (this.artifactDir && fs.existsSync(this.artifactDir)) {
      const artReport = path.join(this.artifactDir, 'agent_failure_report.md');
      const artJsonReport = path.join(this.artifactDir, 'agent_failure_report.json');
      if (fs.existsSync(artReport)) {
        try {
          fs.unlinkSync(artReport);
        } catch (e) {}
      }
      if (fs.existsSync(artJsonReport)) {
        try {
          fs.unlinkSync(artJsonReport);
        } catch (e) {}
      }
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
      // Generate the markdown and JSON failure reports
      this.generateReport(this.failures);
    }
  }

  generateReport(failures) {
    let md = `# 🚨 Agent Unit Test Failure Report\n\n`;
    md += `This diagnostic report was generated automatically. AI agents should review the failures below, fix the corresponding code, and re-run the validation.\n\n`;
    md += `## Summary of Failures\n\n`;
    md += `| Test Suite / File | Test Name | Location |\n`;
    md += `| :--- | :--- | :--- |\n`;

    failures.forEach(f => {
      const fileBase = path.basename(f.testFilePath);
      const loc = f.failureLocation ? `[${path.basename(f.failureLocation.file)}:${f.failureLocation.line}](file:///${f.failureLocation.file.replace(/\\/g, '/')}#L${f.failureLocation.line})` : 'N/A';
      md += `| \`${fileBase}\` | ${f.title} | ${loc} |\n`;
    });

    md += `\n---\n\n## Failure Details\n\n`;

    failures.forEach((f, idx) => {
      md += `### ${idx + 1}. ${f.fullName}\n\n`;
      md += `> [!CAUTION]\n`;
      md += `> **Error message:**\n`;
      md += `> \`\`\`\n`;
      // Clean up stack trace formatting
      const cleanMsg = f.failureMessages.join('\n').replace(/\x1B\[\d+m/g, ''); // strip terminal ANSI colors
      md += `> ${cleanMsg.split('\n').slice(0, 10).join('\n> ')}\n`; // first 10 lines
      md += `> \`\`\`\n\n`;

      if (f.failureLocation) {
        md += `* **File Link:** [${f.failureLocation.file}](file:///${f.failureLocation.file.replace(/\\/g, '/')}#L${f.failureLocation.line})\n`;
      }

      if (f.codeExcerpt) {
        md += `\n**Code Excerpt:**\n`;
        md += `\`\`\`javascript\n`;
        md += `${f.codeExcerpt}\n`;
        md += `\`\`\`\n`;
      }

      if (f.mockTrace && f.mockTrace.length > 0) {
        md += `\n**Mock Google API Calls History:**\n\n`;
        md += `| Time | Service | Method | Arguments | Status / Return Value |\n`;
        md += `| :--- | :--- | :--- | :--- | :--- |\n`;
        f.mockTrace.forEach(t => {
          const timeStr = t.timestamp ? t.timestamp.split('T')[1].slice(0, 8) : '--:--:--'; // HH:MM:SS
          const argsStr = t.arguments ? JSON.stringify(t.arguments) : '[]';
          const retValStr = t.status === 'SUCCESS' 
            ? (t.returnValue !== undefined ? JSON.stringify(t.returnValue) : 'undefined')
            : `❌ ${t.error || 'Unknown Error'}`;
          // Truncate long lines to keep markdown table clean
          const trunc = (str, max) => {
            const s = str || '';
            return s.length > max ? s.slice(0, max) + '...' : s;
          };
          md += `| ${timeStr} | \`${t.service}\` | \`${t.method}\` | \`${trunc(argsStr, 45)}\` | \`${trunc(retValStr, 45)}\` |\n`;
        });
      } else {
        md += `\n*No mock API calls were registered during this test.*\n`;
      }

      md += `\n---\n\n`;
    });

    // Write to workspace directory
    fs.writeFileSync(this.localReportPath, md, 'utf8');
    console.log(`\n📝 Local failure report created: file:///${this.localReportPath.replace(/\\/g, '/')}`);

    const jsonReport = {
      summary: {
        totalFailures: failures.length,
        timestamp: new Date().toISOString()
      },
      failures: failures.map(f => ({
        testSuite: f.testFilePath ? path.relative(__dirname, f.testFilePath) : 'Unknown Suite',
        testName: f.fullName,
        errorType: f.unmockedMethod ? 'MissingMockError' : 'GenericTestFailure',
        unmockedMethod: f.unmockedMethod || null,
        suggestedMockScaffolding: f.suggestedMockScaffolding || null,
        location: f.failureLocation ? {
          file: path.relative(__dirname, f.failureLocation.file),
          line: f.failureLocation.line
        } : null,
        codeExcerpt: f.codeExcerpt || null,
        mockTrace: f.mockTrace || []
      }))
    };
    fs.writeFileSync(this.localJsonReportPath, JSON.stringify(jsonReport, null, 2), 'utf8');
    console.log(`📝 Local JSON diagnostic report created: file:///${this.localJsonReportPath.replace(/\\/g, '/')}`);

    // Write to conversation artifact folder if provided
    if (this.artifactDir && fs.existsSync(this.artifactDir)) {
      const artifactPath = path.join(this.artifactDir, 'agent_failure_report.md');
      fs.writeFileSync(artifactPath, md, 'utf8');
      console.log(`📝 Artifact diagnostic report created: file:///${artifactPath.replace(/\\/g, '/')}`);
      
      const artifactJsonPath = path.join(this.artifactDir, 'agent_failure_report.json');
      fs.writeFileSync(artifactJsonPath, JSON.stringify(jsonReport, null, 2), 'utf8');
      console.log(`📝 Artifact JSON diagnostic report created: file:///${artifactJsonPath.replace(/\\/g, '/')}`);
    }
  }
}

module.exports = AgentReporter;

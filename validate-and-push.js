const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

// Extract artifact directory argument if provided
const args = process.argv.slice(2);
let artifactDir = '';
const dirIdx = args.indexOf('--artifact-dir');
if (dirIdx !== -1 && args[dirIdx + 1]) {
  artifactDir = args[dirIdx + 1];
}

const localReportPath = path.join(__dirname, 'tests/reporters/agent_failure_report.md');
const localJsonReportPath = path.join(__dirname, 'tests/reporters/agent_failure_report.json');

console.log('🔄 Running local unit tests validation...');

// 1. Run Jest
const jestRun = spawnSync('npx', ['jest', '--runInBand'], {
  stdio: 'inherit',
  shell: true,
  env: { ...process.env, NODE_ENV: 'test' }
});

const failureJsonPath = path.join(__dirname, 'tests/reporters/failures.json');

if (jestRun.status !== 0) {
  console.error('\n❌ [VALIDATION FAILED] Unit tests failed. Aborting deploy.');

  if (fs.existsSync(failureJsonPath)) {
    try {
      const failures = JSON.parse(fs.readFileSync(failureJsonPath, 'utf8'));
      generateReport(failures);
    } catch (err) {
      console.error('Failed to generate failure report:', err.message);
    }
  } else {
    console.warn('No structured failure JSON found. Run Jest with AgentReporter to generate diagnostics.');
  }
  process.exit(1);
}

console.log('\n✅ [VALIDATION SUCCESS] All tests passed! Deploying changes to Apps Script editor.');

// Clean up any old local failure reports if tests passed
if (fs.existsSync(localReportPath)) {
  try { fs.unlinkSync(localReportPath); } catch (e) {}
}
if (fs.existsSync(localJsonReportPath)) {
  try {
    fs.unlinkSync(localJsonReportPath);
    console.log('🗑️ Cleaned up old local failure reports (MD/JSON).');
  } catch (e) {}
}

// 2. Run Clasp Push
const claspPush = spawnSync('npx', ['clasp', 'push', '-f'], {
  stdio: 'inherit',
  shell: true
});

if (claspPush.status !== 0) {
  console.error('\n❌ [DEPLOY FAILED] Clasp push encountered errors.');
  process.exit(1);
}

console.log('\n🚀 [DEPLOY SUCCESS] Code successfully validated and pushed!');
process.exit(0);

// --- Report Generation Helper ---
function generateReport(failures) {
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
  fs.writeFileSync(localReportPath, md, 'utf8');
  console.log(`\n📝 Local failure report created: file:///${localReportPath.replace(/\\/g, '/')}`);

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
  fs.writeFileSync(localJsonReportPath, JSON.stringify(jsonReport, null, 2), 'utf8');
  console.log(`📝 Local JSON diagnostic report created: file:///${localJsonReportPath.replace(/\\/g, '/')}`);

  // Write to conversation artifact folder if provided
  if (artifactDir && fs.existsSync(artifactDir)) {
    const artifactPath = path.join(artifactDir, 'agent_failure_report.md');
    fs.writeFileSync(artifactPath, md, 'utf8');
    console.log(`📝 Artifact diagnostic report created: file:///${artifactPath.replace(/\\/g, '/')}`);
    
    const artifactJsonPath = path.join(artifactDir, 'agent_failure_report.json');
    fs.writeFileSync(artifactJsonPath, JSON.stringify(jsonReport, null, 2), 'utf8');
    console.log(`📝 Artifact JSON diagnostic report created: file:///${artifactJsonPath.replace(/\\/g, '/')}`);
  }
}

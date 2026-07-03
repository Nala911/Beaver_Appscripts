const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

// Extract arguments
const args = process.argv.slice(2);
let artifactDir = '';
const dirIdx = args.indexOf('--artifact-dir');
if (dirIdx !== -1 && args[dirIdx + 1]) {
  artifactDir = args[dirIdx + 1];
}

const isForce = args.includes('--force') || args.includes('-f');
const isTestOnly = args.includes('--test-only');
const localReportPath = path.join(__dirname, 'tests/reporters/agent_failure_report.md');
const localJsonReportPath = path.join(__dirname, 'tests/reporters/agent_failure_report.json');

console.log('🔄 Starting deployment pipeline...');

// Clean up old reports in case they exist
try {
  if (fs.existsSync(localReportPath)) fs.unlinkSync(localReportPath);
  if (fs.existsSync(localJsonReportPath)) fs.unlinkSync(localJsonReportPath);
} catch (e) {}

// --- Helper Functions ---

function getChangedFiles() {
  const result = spawnSync('git', ['status', '--porcelain'], { encoding: 'utf8', shell: true });
  if (result.status !== 0) {
    console.warn('⚠️ Warning: Failed to run "git status". Falling back to running all tests.');
    return null;
  }
  
  const lines = result.stdout.split('\n');
  const files = [];
  for (const line of lines) {
    if (!line.trim()) continue;
    
    // git status --porcelain output prefix is 3 chars (XY )
    let filePath = line.substring(3).trim();
    
    // Handle renamed files: "R  old -> new" or "R  \"old\" -> \"new\""
    if (line.startsWith('R ')) {
      const parts = filePath.split(' -> ');
      if (parts.length > 1) {
        filePath = parts[1].replace(/^"(.*)"$/, '$1');
      }
    } else {
      filePath = filePath.replace(/^"(.*)"$/, '$1');
    }
    
    files.push(filePath.replace(/\\/g, '/'));
  }
  return files;
}

function analyzeChanges(files) {
  let runAllTests = false;
  let claspPushNeeded = false;
  const testFilesToRun = new Set();
  
  if (files === null) {
    return { runAllTests: true, claspPushNeeded: true, testFilesToRun: [] };
  }
  
  if (files.length === 0) {
    return { runAllTests: false, claspPushNeeded: false, testFilesToRun: [] };
  }
  
  for (const file of files) {
    // If it's a test file, just run that specific test (no clasp push needed)
    if (file.endsWith('.test.js')) {
      if (fs.existsSync(path.join(__dirname, file))) {
        testFilesToRun.add(file);
      }
      continue;
    }

    // Check if it's a core/setup/configuration change
    if (
      file === 'package.json' ||
      file === 'package-lock.json' ||
      file === 'jest.config.js' ||
      file === 'tests/setup.js' ||
      file.startsWith('tests/mocks/') ||
      file.startsWith('core/')
    ) {
      runAllTests = true;
      claspPushNeeded = true;
      continue;
    }
    
    // Check if it's a tool file
    if (file.startsWith('tools/')) {
      if (file.endsWith('.js') || file.endsWith('.html') || file.endsWith('.json')) {
        claspPushNeeded = true;
      }
      
      const parts = file.split('/');
      let toolName = '';
      if (parts.length > 2) {
        toolName = parts[1];
      } else if (parts.length === 2) {
        const filename = parts[1];
        const match = filename.match(/^([A-Za-z0-9-]+)(?:_Code|_Sidebar)?/);
        if (match) {
          toolName = match[1];
        }
      }
      
      if (toolName) {
        // Find co-located tests for this tool
        const toolDir = path.join(__dirname, 'tools', toolName);
        if (fs.existsSync(toolDir)) {
          const filesInDir = fs.readdirSync(toolDir);
          for (const f of filesInDir) {
            if (f.endsWith('.test.js')) {
              testFilesToRun.add(`tools/${toolName}/${f}`);
            }
          }
        }
      }
      continue;
    }
    
    // Check for general clasp-tracked files at root level
    if (file === 'appsscript.json') {
      claspPushNeeded = true;
      continue;
    }
  }
  
  return {
    runAllTests,
    claspPushNeeded,
    testFilesToRun: Array.from(testFilesToRun)
  };
}

// --- Main Execution Flow ---

let runAllTests = false;
let claspPushNeeded = false;
let testFilesToRun = [];
let changedFiles = [];

if (isForce) {
  console.log('🔄 Force flag active. Running all tests.');
  runAllTests = true;
  claspPushNeeded = true;
} else {
  changedFiles = getChangedFiles();
  if (changedFiles === null) {
    runAllTests = true;
    claspPushNeeded = true;
  } else if (changedFiles.length === 0) {
    console.log('✨ No changes detected. Nothing to test.');
    process.exit(0);
  } else {
    console.log(`🔍 Detected changed files (${changedFiles.length}):`);
    changedFiles.forEach(f => console.log(`   - ${f}`));
    
    const analysis = analyzeChanges(changedFiles);
    runAllTests = analysis.runAllTests;
    claspPushNeeded = analysis.claspPushNeeded;
    testFilesToRun = analysis.testFilesToRun;
  }
}

// Execute tests if needed
let jestPassed = true;

if (runAllTests || testFilesToRun.length > 0) {
  const target = runAllTests ? 'all tests' : `related tests:\n${testFilesToRun.map(t => `   - ${t}`).join('\n')}`;
  console.log(`\n🧪 Running ${target}...`);
  
  // Forward all args EXCEPT pipeline-specific flags so --artifact-dir reaches the AgentReporter
  const jestArgs = ['jest', '--runInBand', ...args.filter(arg => arg !== '--test-only' && arg !== '--force' && arg !== '-f')];
  if (!runAllTests) {
    jestArgs.push(...testFilesToRun);
  }
  
  const jestRun = spawnSync('npx', jestArgs, {
    stdio: 'inherit',
    shell: true,
    env: { ...process.env, NODE_ENV: 'test' }
  });
  
  if (jestRun.status !== 0) {
    jestPassed = false;
  }
} else {
  console.log('\nℹ️ No related tests found or needed. Skipping test execution.');
}

if (!jestPassed) {
  console.error('\n❌ [VALIDATION FAILED] Unit tests failed. Aborting deploy.');
  // The AgentReporter has already generated the diagnostic reports.
  process.exit(1);
}

console.log('\n✅ [VALIDATION SUCCESS] All tests passed / verification succeeded!');

// Run clasp push if needed
if (isTestOnly) {
  console.log('\nℹ️ Test-only mode active. Skipping clasp push.');
} else if (claspPushNeeded) {
  console.log('\n🚀 Deploying changes to Apps Script editor...');
  
  console.log('\n🏗️ Running build step...');
  const buildRun = spawnSync('node', ['build.js'], {
    stdio: 'inherit',
    shell: true
  });

  if (buildRun.status !== 0) {
    console.error('\n❌ [DEPLOY FAILED] Build step encountered errors.');
    process.exit(1);
  }

  const claspPush = spawnSync('npx', ['clasp', 'push', '-f'], {
    stdio: 'inherit',
    shell: true
  });

  if (claspPush.status !== 0) {
    console.error('\n❌ [DEPLOY FAILED] Clasp push encountered errors.');
    process.exit(1);
  }
  console.log('\n🚀 [DEPLOY SUCCESS] Code successfully validated and pushed!');
} else {
  console.log('\nℹ️ No Google Apps Script code changes detected. Skipping clasp push.');
}

process.exit(0);

#!/usr/bin/env node

/**
 * Workspace Sync Appscripts - Unified Agent CLI Harness
 * 
 * Main entry point for AI agents in Antigravity to:
 *  - Test: Smart, CPU-optimized test runner with selective file execution & pre-flight syntax checks.
 *  - Build: Validate core evaluation order, check syntax, and bundle files to dist/Code.js.
 *  - Deploy: Complete validation, build, and clasp push pipeline.
 *  - Scaffold: Auto-generate new Beaver tools with code, view, and test harnesses.
 *  - Status: Full system diagnostics (core order, tool list, test coverage, last test report).
 *  - Benchmark: Measure test suite performance without overloading CPU.
 */

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const crypto = require('crypto');

const ROOT_DIR = path.resolve(__dirname, '..');
const CORE_DIR = path.join(ROOT_DIR, 'core');
const TOOLS_DIR = path.join(ROOT_DIR, 'tools');
const TESTS_DIR = path.join(ROOT_DIR, 'tests');
const DIST_DIR = path.join(ROOT_DIR, 'dist');
const CACHE_FILE = path.join(ROOT_DIR, '.agent_cache.json');
const REPORTERS_DIR = path.join(TESTS_DIR, 'reporters');
const FAILURE_MD = path.join(REPORTERS_DIR, 'agent_failure_report.md');
const FAILURE_JSON = path.join(REPORTERS_DIR, 'agent_failure_report.json');
const STATUS_JSON = path.join(REPORTERS_DIR, 'agent_status_report.json');
const STATUS_MD = path.join(REPORTERS_DIR, 'agent_status_report.md');

// Utility: parse CLI flags
function parseArgs() {
  const rawArgs = process.argv.slice(2);
  const command = rawArgs[0] && !rawArgs[0].startsWith('-') ? rawArgs[0] : 'help';
  const flags = {};
  
  for (let i = (command === 'help' ? 0 : 1); i < rawArgs.length; i++) {
    const arg = rawArgs[i];
    if (arg.startsWith('--')) {
      const parts = arg.substring(2).split('=');
      const key = parts[0];
      const val = parts.length > 1 ? parts.slice(1).join('=') : true;
      flags[key] = val;
    } else if (arg.startsWith('-')) {
      flags[arg.substring(1)] = true;
    }
  }
  
  return { command, flags, rawArgs };
}

// Exit codes standard:
// 0 = Success
// 1 = Unit Test Failure
// 2 = Syntax / Build Error
// 3 = Invalid Usage / Scaffolding Error
// 4 = Deployment Error
function exitProcess(code, message) {
  if (message) {
    if (code === 0) {
      console.log(`\n✨ ${message}`);
    } else {
      console.error(`\n❌ [EXIT CODE ${code}] ${message}`);
    }
  }
  process.exit(code);
}

// Fast node syntax check (--check) on a file
function checkSyntax(filePath) {
  const result = spawnSync(process.execPath, ['--check', filePath], { encoding: 'utf8' });
  if (result.status !== 0) {
    return {
      valid: false,
      error: result.stderr || result.stdout || 'Syntax check failed'
    };
  }
  return { valid: true };
}

// Validate core load order (files must start with digits, e.g., 00_..., 01_...)
function validateCoreOrder() {
  if (!fs.existsSync(CORE_DIR)) return { valid: true, files: [] };
  
  const files = fs.readdirSync(CORE_DIR)
    .filter(f => f.endsWith('.js') && !f.endsWith('.test.js'))
    .sort();
    
  const errors = [];
  files.forEach(f => {
    if (!/^\d{2}_/.test(f)) {
      errors.push(`File "core/${f}" does not follow numerical prefix naming convention (e.g. 00_Config.js).`);
    }
  });

  return {
    valid: errors.length === 0,
    errors,
    files
  };
}

// Pre-flight syntax check across changed JS files or all JS files
function runSyntaxPreflight(filesToCheck) {
  console.log('⚡ Running fast pre-flight syntax check...');
  const invalidFiles = [];

  for (const relPath of filesToCheck) {
    const absPath = path.join(ROOT_DIR, relPath);
    if (fs.existsSync(absPath) && absPath.endsWith('.js')) {
      const check = checkSyntax(absPath);
      if (!check.valid) {
        invalidFiles.push({ file: relPath, error: check.error });
      }
    }
  }

  if (invalidFiles.length > 0) {
    console.error('\n🚨 [SYNTAX PRE-FLIGHT FAILED] Found syntax errors in source files:');
    invalidFiles.forEach(item => {
      console.error(`   - ${item.file}`);
      console.error(`     ${item.error.trim().replace(/\n/g, '\n     ')}`);
    });
    return false;
  }

  console.log('✅ Syntax pre-flight passed cleanly.');
  return true;
}

// Git status changed files helper
function getGitChangedFiles() {
  const result = spawnSync('git', ['status', '--porcelain'], { encoding: 'utf8', shell: true });
  if (result.status !== 0) {
    return null;
  }
  
  const lines = result.stdout.split('\n');
  const files = [];
  for (const line of lines) {
    if (!line.trim()) continue;
    let filePath = line.substring(3).trim();
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

// File hash helper for smart caching
function getFileHash(filePath) {
  if (!fs.existsSync(filePath)) return '';
  const content = fs.readFileSync(filePath);
  return crypto.createHash('sha256').update(content).digest('hex');
}

// Load / save cache
function loadCache() {
  if (fs.existsSync(CACHE_FILE)) {
    try {
      return JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8'));
    } catch (e) {}
  }
  return { hashes: {}, lastTestResult: null };
}

function saveCache(cache) {
  try {
    fs.writeFileSync(CACHE_FILE, JSON.stringify(cache, null, 2), 'utf8');
  } catch (e) {}
}

// --- COMMAND IMPLEMENTATIONS ---

// 1. COMMAND: TEST
function commandTest(flags) {
  console.log('🧪 [AGENT HARNESS] Starting smart CPU-optimized test runner...');
  const startTime = Date.now();

  const isForce = flags.force || flags.f;
  const targetTool = flags.tool;
  const targetFile = flags.file;
  const lowCpu = flags['low-cpu'] || true; // Default to CPU-friendly mode
  const artifactDir = flags['artifact-dir'] || '';

  let testFiles = new Set();
  let filesToSyntaxCheck = [];
  let runAll = false;

  if (targetFile) {
    testFiles.add(targetFile);
    filesToSyntaxCheck.push(targetFile);
  } else if (targetTool) {
    const toolDir = path.join(TOOLS_DIR, targetTool);
    if (!fs.existsSync(toolDir)) {
      exitProcess(3, `Tool directory "tools/${targetTool}" does not exist.`);
    }
    const items = fs.readdirSync(toolDir);
    items.forEach(item => {
      if (item.endsWith('.js')) {
        filesToSyntaxCheck.push(`tools/${targetTool}/${item}`);
        if (item.endsWith('.test.js')) {
          testFiles.add(`tools/${targetTool}/${item}`);
        }
      }
    });
    if (testFiles.size === 0) {
      console.warn(`⚠️ No test files (.test.js) found for tool "${targetTool}".`);
    }
  } else if (isForce) {
    runAll = true;
  } else {
    const changed = getGitChangedFiles();
    if (changed === null) {
      runAll = true;
    } else if (changed.length === 0) {
      console.log('✨ No modified files detected in Git. Using content hash cache check...');
      const cache = loadCache();
      let anyHashChanged = false;

      // Scan core and tools for hash changes
      const scanDir = (dir, prefix) => {
        if (!fs.existsSync(dir)) return;
        const list = fs.readdirSync(dir);
        list.forEach(item => {
          const full = path.join(dir, item);
          const rel = `${prefix}/${item}`;
          if (fs.statSync(full).isDirectory()) {
            scanDir(full, rel);
          } else if (item.endsWith('.js') || item.endsWith('.html')) {
            const h = getFileHash(full);
            if (cache.hashes[rel] !== h) {
              cache.hashes[rel] = h;
              anyHashChanged = true;
              filesToSyntaxCheck.push(rel);
            }
          }
        });
      };

      scanDir(CORE_DIR, 'core');
      scanDir(TOOLS_DIR, 'tools');

      if (!anyHashChanged && cache.lastTestResult === 'SUCCESS') {
        console.log('✨ Code content and test files are unchanged. All previous tests passed! (0ms CPU cost)');
        exitProcess(0, 'Test runner completed (Cached success).');
      } else {
        runAll = true;
      }
    } else {
      console.log(`🔍 Detected ${changed.length} changed file(s):`);
      changed.forEach(f => console.log(`   - ${f}`));

      changed.forEach(file => {
        if (file.endsWith('.js')) filesToSyntaxCheck.push(file);

        if (file.endsWith('.test.js')) {
          testFiles.add(file);
        } else if (file.startsWith('core/') || file === 'package.json' || file === 'jest.config.js' || file === 'tests/setup.js') {
          runAll = true;
        } else if (file.startsWith('tools/')) {
          const parts = file.split('/');
          if (parts.length >= 2) {
            const tName = parts[1];
            const tDir = path.join(TOOLS_DIR, tName);
            if (fs.existsSync(tDir)) {
              fs.readdirSync(tDir).forEach(f => {
                if (f.endsWith('.test.js')) testFiles.add(`tools/${tName}/${f}`);
              });
            }
          }
        }
      });
    }
  }

  // Pre-flight syntax check
  if (filesToSyntaxCheck.length > 0) {
    if (!runSyntaxPreflight(filesToSyntaxCheck)) {
      exitProcess(2, 'Pre-flight syntax check failed.');
    }
  }

  // Build Jest arguments with CPU governor
  const jestArgs = ['jest'];
  
  // CPU Optimization: limit worker concurrency to prevent CPU spikes / overheating
  if (lowCpu || testFiles.size <= 2) {
    jestArgs.push('--maxWorkers=50%'); // Gentle CPU load
  }

  if (artifactDir) {
    jestArgs.push('--artifact-dir', artifactDir);
  }

  if (!runAll && testFiles.size > 0) {
    const list = Array.from(testFiles);
    console.log(`\n🎯 Running targeted test suites (${list.length}):`);
    list.forEach(t => console.log(`   - ${t}`));
    jestArgs.push(...list);
  } else {
    console.log('\n🧪 Running full test suite...');
  }

  const jestResult = spawnSync('npx', jestArgs, {
    stdio: 'inherit',
    shell: true,
    env: { ...process.env, NODE_ENV: 'test' }
  });

  const durationSec = ((Date.now() - startTime) / 1000).toFixed(2);
  const cache = loadCache();

  if (jestResult.status !== 0) {
    cache.lastTestResult = 'FAILED';
    saveCache(cache);
    console.error(`\n❌ [TEST FAILED] Execution finished in ${durationSec}s.`);
    if (fs.existsSync(FAILURE_MD)) {
      console.error(`📝 Diagnostic failure report: file:///${FAILURE_MD.replace(/\\/g, '/')}`);
    }
    if (fs.existsSync(FAILURE_JSON)) {
      console.error(`📝 Diagnostic JSON report: file:///${FAILURE_JSON.replace(/\\/g, '/')}`);
    }
    exitProcess(1, 'Unit tests failed.');
  } else {
    cache.lastTestResult = 'SUCCESS';
    // Update content hashes
    const updateHashes = (dir, prefix) => {
      if (!fs.existsSync(dir)) return;
      fs.readdirSync(dir).forEach(item => {
        const full = path.join(dir, item);
        const rel = `${prefix}/${item}`;
        if (fs.statSync(full).isDirectory()) {
          updateHashes(full, rel);
        } else if (item.endsWith('.js') || item.endsWith('.html')) {
          cache.hashes[rel] = getFileHash(full);
        }
      });
    };
    updateHashes(CORE_DIR, 'core');
    updateHashes(TOOLS_DIR, 'tools');
    saveCache(cache);

    console.log(`\n✅ [TEST SUCCESS] All tests passed in ${durationSec}s.`);
    exitProcess(0, 'Test runner completed successfully.');
  }
}

// 2. COMMAND: BUILD
function commandBuild(flags) {
  console.log('🏗️ [AGENT HARNESS] Starting build & validation process...');
  const startTime = Date.now();

  // 1. Verify Core Load/Evaluation Order
  console.log('🔍 Checking core file evaluation order...');
  const coreCheck = validateCoreOrder();
  if (!coreCheck.valid) {
    console.error('❌ Core evaluation order validation failed:');
    coreCheck.errors.forEach(e => console.error(`   - ${e}`));
    exitProcess(2, 'Core load order invalid.');
  }
  console.log(`✅ Core load order verified (${coreCheck.files.length} files in order).`);

  // 2. Collect all JS source files for syntax check
  const allJs = [];
  const collectJs = (dir, prefix) => {
    if (!fs.existsSync(dir)) return;
    fs.readdirSync(dir).forEach(item => {
      const full = path.join(dir, item);
      const rel = `${prefix}/${item}`;
      if (fs.statSync(full).isDirectory()) {
        collectJs(full, rel);
      } else if (item.endsWith('.js') && !item.endsWith('.test.js')) {
        allJs.push(rel);
      }
    });
  };
  collectJs(CORE_DIR, 'core');
  collectJs(TOOLS_DIR, 'tools');

  if (!runSyntaxPreflight(allJs)) {
    exitProcess(2, 'Build aborted due to syntax errors.');
  }

  if (flags['check-only']) {
    exitProcess(0, 'Syntax and core evaluation order checks passed (--check-only).');
  }

  // 3. Run build.js process
  console.log('📦 Bundling distribution code...');
  const buildRun = spawnSync('node', ['build.js'], { stdio: 'inherit', shell: true });
  if (buildRun.status !== 0) {
    exitProcess(2, 'Build bundling step failed.');
  }

  const durationSec = ((Date.now() - startTime) / 1000).toFixed(2);
  console.log(`\n✨ Build bundle successfully created in dist/Code.js (${durationSec}s).`);
  exitProcess(0, 'Build completed successfully.');
}

// 3. COMMAND: DEPLOY
function commandDeploy(flags) {
  console.log('🚀 [AGENT HARNESS] Starting full validation and deployment pipeline...');
  const isForce = flags.force || flags.f;
  const isTestOnly = flags['test-only'];
  const lowCpu = flags['low-cpu'] || true;

  // Step A: Run tests
  console.log('\n--- PHASE 1: TESTING ---');
  
  const testRun = spawnSync(process.execPath, [path.join(__dirname, 'agent-cli.js'), 'test', ...(isForce ? ['--force'] : [])], { stdio: 'inherit' });
  if (testRun.status !== 0) {
    exitProcess(1, 'Deployment aborted due to test failure.');
  }

  if (isTestOnly) {
    exitProcess(0, 'Test-only pipeline completed successfully.');
  }

  // Step B: Run build
  console.log('\n--- PHASE 2: BUILDING ---');
  const buildRun = spawnSync(process.execPath, [path.join(__dirname, 'agent-cli.js'), 'build'], { stdio: 'inherit' });
  if (buildRun.status !== 0) {
    exitProcess(2, 'Deployment aborted due to build error.');
  }

  // Step C: Run Clasp Push
  console.log('\n--- PHASE 3: CLASP PUSH ---');
  console.log('📡 Pushing code to Google Apps Script editor...');
  const claspPush = spawnSync('npx', ['clasp', 'push', '-f'], { stdio: 'inherit', shell: true });
  if (claspPush.status !== 0) {
    exitProcess(4, 'Clasp push failed.');
  }

  console.log('\n🎉 [DEPLOY SUCCESS] All code validated, tested, bundled, and deployed cleanly!');
  exitProcess(0, 'Deploy pipeline finished.');
}

// 4. COMMAND: SCAFFOLD
function commandScaffold(flags) {
  const toolName = flags.name;
  const description = flags.description || `${toolName} Tool for Workspace Sync`;

  if (!toolName || !/^[A-Za-z0-9_]+$/.test(toolName)) {
    console.error('❌ Error: Must specify valid tool name using --name=ToolName (alphanumeric & underscore only).');
    console.error('   Example: node scripts/agent-cli.js scaffold --name=InvoiceGenerator --description="Generates PDFs"');
    exitProcess(3, 'Invalid tool name.');
  }

  const toolDir = path.join(TOOLS_DIR, toolName);
  if (fs.existsSync(toolDir)) {
    console.error(`❌ Error: Tool directory "tools/${toolName}" already exists.`);
    exitProcess(3, 'Tool already exists.');
  }

  console.log(`🛠️ Scaffolding new tool: "${toolName}"...`);
  fs.mkdirSync(toolDir, { recursive: true });

  // 1. Code.js template
  const codeContent = `/**
 * ${toolName}_Code.js
 * ${description}
 */

(function () {
  /**
   * Main entry point for ${toolName} execution.
   * @param {Object} params Input parameters from sidebar or triggers.
   * @return {Object} Result payload.
   */
  function run${toolName}(params) {
    try {
      if (typeof Logger !== 'undefined') {
        Logger.info('${toolName}', 'run', 'Starting execution', params);
      }
      
      // Implement tool logic here
      const result = {
        success: true,
        message: '${toolName} executed successfully.',
        timestamp: new Date().toISOString()
      };

      return result;
    } catch (err) {
      if (typeof Logger !== 'undefined') {
        Logger.error('${toolName}', 'run', err);
      }
      if (typeof SyncEngine !== 'undefined' && SyncEngine.Utils) {
        return SyncEngine.Utils.translateApiError(err, '${toolName}');
      }
      return { success: false, message: err.message || String(err) };
    }
  }

  // Register in global SyncEngine namespace
  if (typeof globalThis.SyncEngine === 'undefined') {
    globalThis.SyncEngine = {};
  }
  globalThis.SyncEngine.${toolName} = {
    run: run${toolName}
  };

  // Expose global function for Apps Script UI calling
  globalThis.run${toolName} = run${toolName};
})();
`;
  fs.writeFileSync(path.join(toolDir, 'Code.js'), codeContent, 'utf8');
  console.log(`   + Created tools/${toolName}/Code.js`);

  // 2. Sidebar.html template
  const sidebarContent = `<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <title>${toolName}</title>
    <?!= HtmlService.createHtmlOutputFromFile('core/SidebarShared').getContent(); ?>
    <style>
      .tool-container {
        padding: 16px;
      }
    </style>
  </head>
  <body>
    <div class="tool-container">
      <h2>${toolName}</h2>
      <p class="text-secondary">${description}</p>
      
      <div style="margin-top: 20px;">
        <button id="btnRun" class="btn btn-primary" onclick="handleRun()">Run ${toolName}</button>
      </div>

      <div id="statusOutput" style="margin-top: 15px;"></div>
    </div>

    <script>
      function handleRun() {
        const btn = document.getElementById('btnRun');
        const statusDiv = document.getElementById('statusOutput');
        btn.disabled = true;
        statusDiv.innerHTML = '<span class="spinner"></span> Processing...';

        google.script.run
          .withSuccessHandler(function(response) {
            btn.disabled = false;
            if (response && response.success) {
              statusDiv.innerHTML = '<div class="alert alert-success">' + response.message + '</div>';
            } else {
              statusDiv.innerHTML = '<div class="alert alert-danger">' + (response.message || 'Execution failed') + '</div>';
            }
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            statusDiv.innerHTML = '<div class="alert alert-danger">Error: ' + err.message + '</div>';
          })
          .run${toolName}({});
      }
    </script>
  </body>
</html>
`;
  fs.writeFileSync(path.join(toolDir, 'Sidebar.html'), sidebarContent, 'utf8');
  console.log(`   + Created tools/${toolName}/Sidebar.html`);

  // 3. Code.test.js template
  const testContent = `/**
 * ${toolName} Tool Unit Tests
 */

describe('${toolName} Tool', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('should execute run${toolName} successfully', () => {
    const result = global.run${toolName}({});
    expect(result).toBeDefined();
    expect(result.success).toBe(true);
    expect(result.message).toContain('${toolName} executed successfully');
  });
});
`;
  fs.writeFileSync(path.join(toolDir, 'Code.test.js'), testContent, 'utf8');
  console.log(`   + Created tools/${toolName}/Code.test.js`);

  console.log(`\n✨ Tool "${toolName}" successfully scaffolded!`);
  console.log(`   Run test with: npm run agent:test -- --tool=${toolName}`);
  exitProcess(0, 'Scaffolding complete.');
}

// 5. COMMAND: STATUS
function commandStatus(flags) {
  console.log('📊 [AGENT HARNESS] Generating System Status Diagnostics...');

  const coreOrder = validateCoreOrder();
  const toolsList = [];

  if (fs.existsSync(TOOLS_DIR)) {
    fs.readdirSync(TOOLS_DIR).forEach(t => {
      const tPath = path.join(TOOLS_DIR, t);
      if (fs.statSync(tPath).isDirectory()) {
        const files = fs.readdirSync(tPath);
        const hasCode = files.includes('Code.js');
        const hasSidebar = files.includes('Sidebar.html');
        const hasTest = files.some(f => f.endsWith('.test.js'));
        toolsList.push({ name: t, hasCode, hasSidebar, hasTest, filesCount: files.length });
      }
    });
  }

  const gitFiles = getGitChangedFiles();
  const cache = loadCache();
  const lastFailureExists = fs.existsSync(FAILURE_JSON);

  const statusReport = {
    timestamp: new Date().toISOString(),
    coreEvaluationOrder: {
      valid: coreOrder.valid,
      totalFiles: coreOrder.files.length,
      files: coreOrder.files,
      errors: coreOrder.errors
    },
    toolsSummary: {
      totalTools: toolsList.length,
      toolsWithTests: toolsList.filter(t => t.hasTest).length,
      tools: toolsList
    },
    gitState: {
      dirty: gitFiles !== null && gitFiles.length > 0,
      changedFilesCount: gitFiles ? gitFiles.length : 0,
      changedFiles: gitFiles || []
    },
    pipelineState: {
      lastTestResult: cache.lastTestResult || 'UNKNOWN',
      hasActiveFailureReport: lastFailureExists
    }
  };

  // Write status report files
  if (!fs.existsSync(REPORTERS_DIR)) {
    fs.mkdirSync(REPORTERS_DIR, { recursive: true });
  }
  fs.writeFileSync(STATUS_JSON, JSON.stringify(statusReport, null, 2), 'utf8');

  // Generate Status Markdown
  let md = `# 📊 Agent Workspace Diagnostics & Status Report\n\n`;
  md += `**Generated:** ${statusReport.timestamp}\n\n`;
  md += `## 1. Core Evaluation Order (${statusReport.coreEvaluationOrder.valid ? '✅ VALID' : '❌ INVALID'})\n\n`;
  statusReport.coreEvaluationOrder.files.forEach((f, i) => {
    md += `${i + 1}. \`core/${f}\`\n`;
  });

  md += `\n## 2. Tools Inventory (${statusReport.toolsSummary.totalTools} Tools, ${statusReport.toolsSummary.toolsWithTests} with Tests)\n\n`;
  md += `| Tool Name | Has Code.js | Has Sidebar | Has Tests |\n`;
  md += `| :--- | :--- | :--- | :--- |\n`;
  statusReport.toolsSummary.tools.forEach(t => {
    md += `| \`${t.name}\` | ${t.hasCode ? '✅' : '❌'} | ${t.hasSidebar ? '✅' : '❌'} | ${t.hasTest ? '✅' : '⚠️ Missing'} |\n`;
  });

  md += `\n## 3. Git Status & Uncommitted Changes\n\n`;
  if (statusReport.gitState.changedFiles.length > 0) {
    md += `Changed files (${statusReport.gitState.changedFilesCount}):\n`;
    statusReport.gitState.changedFiles.forEach(f => md += `- \`${f}\`\n`);
  } else {
    md += `*Working tree clean.*\n`;
  }

  md += `\n## 4. Pipeline Cache State\n\n`;
  md += `- **Last Test Execution:** \`${statusReport.pipelineState.lastTestResult}\`\n`;
  md += `- **Failure Report Active:** \`${statusReport.pipelineState.hasActiveFailureReport}\`\n`;

  fs.writeFileSync(STATUS_MD, md, 'utf8');

  if (flags.json) {
    console.log(JSON.stringify(statusReport, null, 2));
  } else {
    console.log(`\n✅ Status Diagnostics Report created:`);
    console.log(`   - Markdown: file:///${STATUS_MD.replace(/\\/g, '/')}`);
    console.log(`   - JSON:     file:///${STATUS_JSON.replace(/\\/g, '/')}`);
    console.log(`\nSummary: ${toolsList.length} tools, ${coreOrder.files.length} core files, Last test: ${statusReport.pipelineState.lastTestResult}`);
  }

  exitProcess(0);
}

// 6. COMMAND: BENCHMARK
function commandBenchmark(flags) {
  console.log('⏱️ [AGENT HARNESS] Running low-CPU test suite performance benchmark...');
  const startTime = Date.now();

  const jestRun = spawnSync('npx', ['jest', '--maxWorkers=50%', '--verbose'], { stdio: 'inherit', shell: true });
  const duration = ((Date.now() - startTime) / 1000).toFixed(2);

  if (jestRun.status === 0) {
    console.log(`\n⚡ Benchmark complete. All test suites executed cleanly in ${duration}s.`);
    console.log(`   Memory footprint: ${(process.memoryUsage().heapUsed / 1024 / 1024).toFixed(2)} MB`);
    exitProcess(0, 'Benchmark complete.');
  } else {
    exitProcess(1, 'Benchmark failed due to test errors.');
  }
}

// 7. COMMAND: HELP
function commandHelp() {
  console.log(`
🤖 Workspace Sync Appscripts - AI Agent CLI Harness Help

Usage:
  node scripts/agent-cli.js <command> [flags]
  npm run agent:<command> -- [flags]

Available Commands:
  test        Smart, CPU-optimized test runner with syntax pre-flight & change detection.
              Flags: --tool=ToolName, --file=path/to/test.js, --force, --low-cpu, --artifact-dir=dir

  build       Validates core load order, runs syntax check, bundles code to dist/Code.js.
              Flags: --check-only

  deploy      Full pipeline: pre-flight check -> tests -> build -> clasp push.
              Flags: --force, --test-only, --low-cpu

  scaffold    Auto-generates new Beaver tool files with Code.js, Sidebar.html, and Code.test.js.
              Flags: --name=ToolName (required), --description="Description text"

  status      Outputs system diagnostics (core order, tool coverage, git status, last test state).
              Flags: --json

  benchmark   Runs test suites in low-CPU mode to measure execution speed and memory footprint.

  help        Displays this reference menu.

Exit Codes:
  0 = Success
  1 = Unit Test Failure
  2 = Syntax or Core Load Order / Build Error
  3 = Usage or Scaffolding Error
  4 = Deployment / Clasp Error
`);
  exitProcess(0);
}

// Main Router
function main() {
  const { command, flags } = parseArgs();

  switch (command.toLowerCase()) {
    case 'test':
      commandTest(flags);
      break;
    case 'build':
      commandBuild(flags);
      break;
    case 'deploy':
      commandDeploy(flags);
      break;
    case 'scaffold':
      commandScaffold(flags);
      break;
    case 'status':
      commandStatus(flags);
      break;
    case 'benchmark':
      commandBenchmark(flags);
      break;
    case 'help':
    default:
      commandHelp();
      break;
  }
}

main();

# 🤖 AI Agent Entry Points & Execution Harness Guide

This sitemap and reference document provides AI Agents in Antigravity with complete entry points, execution commands, diagnostic reporting formats, and scaffolding instructions for maintaining and refactoring the **Workspace Sync Appscripts** codebase.

---

## ⚡ Quick Reference Commands

AI Agents can execute all workflows directly via `npm run agent:*` or `node scripts/agent-cli.js <command>`:

| Task / Purpose | Command | Description |
| :--- | :--- | :--- |
| **Run Smart Tests** | `npm run agent:test` | Runs change-detected tests with fast syntax pre-flight & CPU throttling. |
| **Target Tool Test** | `npm run agent:test -- --tool=CalendarSync` | Runs tests only for a specific tool. |
| **Force All Tests** | `npm run agent:test -- --force` | Bypasses content-hash cache and runs all test suites. |
| **Build & Validate** | `npm run agent:build` | Audits core evaluation order, checks JS syntax, and bundles to `dist/Code.js`. |
| **Syntax Check Only**| `npm run agent:build -- --check-only` | Validates JavaScript syntax and numerical prefix order without writing bundle. |
| **Full Deploy** | `npm run agent:deploy` | Runs syntax check -> unit tests -> build -> `clasp push`. |
| **Scaffold Tool** | `npm run agent:scaffold -- --name=NewTool` | Auto-generates `Code.js`, `Sidebar.html`, and `Code.test.js` template. |
| **System Status** | `npm run agent:status` | Outputs core order state, tool inventory, git status, and pipeline health. |
| **CPU Benchmark** | `npm run agent:benchmark` | Measures test suite execution timing and heap memory consumption. |

---

## 🔒 Low CPU Processor Policy

To protect the local hardware CPU processor from overheating or freezing:
1. **Worker Concurrency Limit**: All test runners use `--maxWorkers=50%` (or maximum 2 workers) instead of spawning workers for every CPU core.
2. **Fast Pre-flight Syntax Checking**: Node's native `node --check` validates syntax in <5ms per file before spawning Jest processes.
3. **Content Hash Caching**: Hashes are stored in `.agent_cache.json`. If files haven't changed since the last green test run, testing completes in 0ms CPU time.

---

## 📑 Diagnostic Reports & Logs for AI Agents

When tests fail or status is requested, structured diagnostic logs are automatically generated at:

1. **Failure Reports**:
   - Markdown: [`tests/reporters/agent_failure_report.md`](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Workspace%20Sync%20Appscripts/tests/reporters/agent_failure_report.md)
   - JSON: [`tests/reporters/agent_failure_report.json`](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Workspace%20Sync%20Appscripts/tests/reporters/agent_failure_report.json)
   - Content: Clickable `file:///` URLs with line numbers (`#L42`), code excerpts, missing mock scaffolding code, and Google API mock call traces.

2. **Status Diagnostics**:
   - Markdown: [`tests/reporters/agent_status_report.md`](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Workspace%20Sync%20Appscripts/tests/reporters/agent_status_report.md)
   - JSON: [`tests/reporters/agent_status_report.json`](file:///c:/Users/fazil_uxry2im/Documents/Beaver/Workspace%20Sync%20Appscripts/tests/reporters/agent_status_report.json)

---

## 🛠️ Exit Codes Matrix

AI Agents should evaluate script process exit codes:

- **`0`**: Success (All tests passed / Build succeeded / Deploy complete).
- **`1`**: Unit Test Failure (Check `agent_failure_report.json`).
- **`2`**: Syntax Error or Core Evaluation Order Error (Check output logs).
- **`3`**: Invalid Usage / Scaffolding Error.
- **`4`**: Deployment / Clasp Error.

---

## 🏗️ Scaffolding New Beaver Tools

When creating a new tool, run:
```bash
npm run agent:scaffold -- --name=MyNewTool --description="Tool description here"
```
This automatically creates:
- `tools/MyNewTool/Code.js` (with `SyncEngine.MyNewTool` namespace wrapper)
- `tools/MyNewTool/Sidebar.html` (with shared styling header and `google.script.run` bridge)
- `tools/MyNewTool/Code.test.js` (with Jest unit test harness)

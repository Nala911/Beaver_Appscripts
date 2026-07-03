const fs = require('fs');
const path = require('path');

const DIST_DIR = path.join(__dirname, 'dist');

// Helper to clear directory recursively
function clearDir(dirPath) {
  if (fs.existsSync(dirPath)) {
    fs.rmSync(dirPath, { recursive: true, force: true });
  }
}

// Helper to ensure parent directory exists
function ensureParentDir(filePath) {
  const parent = path.dirname(filePath);
  if (!fs.existsSync(parent)) {
    fs.mkdirSync(parent, { recursive: true });
  }
}

console.log('🏗️ Starting build process...');

try {
  // 1. Clean and create dist directory
  clearDir(DIST_DIR);
  fs.mkdirSync(DIST_DIR, { recursive: true });

  // 2. Concatenate JavaScript files
  let concatenatedJs = '// =========================================================\n';
  concatenatedJs += '// BUNDLED GOOGLE APPS SCRIPT CODE (AUTOMATICALLY GENERATED)\n';
  concatenatedJs += '// =========================================================\n\n';

  // Read core/ JS files
  const coreDir = path.join(__dirname, 'core');
  if (fs.existsSync(coreDir)) {
    const files = fs.readdirSync(coreDir)
      .filter(f => f.endsWith('.js') && !f.endsWith('.test.js'))
      .sort(); // Sorts numerically/alphabetically

    files.forEach(file => {
      const filePath = path.join(coreDir, file);
      console.log(`   🔹 Bundling core file: core/${file}`);
      const content = fs.readFileSync(filePath, 'utf8');
      concatenatedJs += `\n// --- FILE: core/${file} ---\n`;
      concatenatedJs += content;
      concatenatedJs += '\n';
    });
  }

  // Read tools/ Code.js files
  const toolsDir = path.join(__dirname, 'tools');
  if (fs.existsSync(toolsDir)) {
    const tools = fs.readdirSync(toolsDir).sort();
    tools.forEach(tool => {
      const toolDir = path.join(toolsDir, tool);
      if (fs.statSync(toolDir).isDirectory()) {
        const jsFiles = fs.readdirSync(toolDir)
          .filter(f => f.endsWith('.js') && !f.endsWith('.test.js'))
          .sort();
        
        jsFiles.forEach(file => {
          const filePath = path.join(toolDir, file);
          console.log(`   🔹 Bundling tool file: tools/${tool}/${file}`);
          const content = fs.readFileSync(filePath, 'utf8');
          concatenatedJs += `\n// --- FILE: tools/${tool}/${file} ---\n`;
          concatenatedJs += content;
          concatenatedJs += '\n';
        });
      }
    });
  }

  // Write concatenated JS to dist/Code.js
  fs.writeFileSync(path.join(DIST_DIR, 'Code.js'), concatenatedJs, 'utf8');
  console.log('✅ JavaScript files bundled into dist/Code.js');

  // 3. Copy appsscript.json
  const appsscriptSrc = path.join(__dirname, 'appsscript.json');
  if (fs.existsSync(appsscriptSrc)) {
    fs.copyFileSync(appsscriptSrc, path.join(DIST_DIR, 'appsscript.json'));
    console.log('✅ appsscript.json copied');
  }

  // 4. Copy HTML files while preserving directory structure
  function copyHtmlFiles(srcDir, destDir) {
    if (!fs.existsSync(srcDir)) return;
    const items = fs.readdirSync(srcDir);
    items.forEach(item => {
      const srcPath = path.join(srcDir, item);
      const relative = path.relative(__dirname, srcPath);
      const destPath = path.join(destDir, relative);

      if (fs.statSync(srcPath).isDirectory()) {
        copyHtmlFiles(srcPath, destDir);
      } else if (item.endsWith('.html')) {
        console.log(`   🔹 Copying HTML template: ${relative.replace(/\\/g, '/')}`);
        ensureParentDir(destPath);
        fs.copyFileSync(srcPath, destPath);
      }
    });
  }

  copyHtmlFiles(coreDir, DIST_DIR);
  copyHtmlFiles(toolsDir, DIST_DIR);
  console.log('✅ HTML template files copied');

  console.log('✨ Build completed successfully!');
} catch (error) {
  console.error('❌ Build failed:', error);
  process.exit(1);
}

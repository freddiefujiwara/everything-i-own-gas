const fs = require("fs");
const path = require("path");

// The source is a set of ES modules under src/, while Apps Script
// runs a single file. src/Code.js is the bundle entry: each of its
// import lines names one module, and the build replaces the line
// with that module's contents, in order. The modules also import
// each other (that is how the tests load them), so the concatenated
// source still holds import statements and export markers; both are
// stripped below, leaving one plain script in a single scope.
const ENTRY_IMPORT_LINE =
  /^import\s+(?:[^'"]*?\s+from\s+)?['"](\.[^'"]+)['"];?\s*$/;

const inlineModules = (entrySource, readModule) =>
  entrySource
    .split("\n")
    .map(line => {
      const match = ENTRY_IMPORT_LINE.exec(line);
      return match ? readModule(match[1]) : line;
    })
    .join("\n");

const stripModuleSyntax = source =>
  source
    .replace(/^import\s+[\s\S]*?from\s+['"][^'"]+['"];?[ \t]*$/gm, "")
    .replace(/^import\s+['"][^'"]+['"];?[ \t]*$/gm, "")
    .replace(/^export\s+/gm, "")
    .replace(/^export\s*\{[^}]*\};?\s*$/gm, "");

const build = ({
  rootDir = process.cwd(),
  fsModule = fs,
  pathModule = path,
} = {}) => {
  const srcDir = pathModule.join(rootDir, "src");
  const appsscriptPath = pathModule.join(rootDir, "appsscript.json");
  const distDir = pathModule.join(rootDir, "dist");
  const distPath = pathModule.join(distDir, "Code.gs");
  const distManifestPath = pathModule.join(distDir, "appsscript.json");

  fsModule.mkdirSync(distDir, { recursive: true });
  const entry = fsModule.readFileSync(pathModule.join(srcDir, "Code.js"), "utf8");
  const inlined = inlineModules(entry, specifier =>
    fsModule.readFileSync(pathModule.join(srcDir, specifier), "utf8")
  );
  fsModule.writeFileSync(distPath, stripModuleSyntax(inlined));
  fsModule.copyFileSync(appsscriptPath, distManifestPath);
};

const runIfMain = ({ main = require.main, current = module, buildFn = build } = {}) => {
  if (main === current) {
    buildFn();
    return true;
  }
  return false;
};

runIfMain();

module.exports = { build, runIfMain };

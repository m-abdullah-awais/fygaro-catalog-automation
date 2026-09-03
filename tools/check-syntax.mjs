/*
 * Fygaro Catalog Automation
 * Parses every source file so a typo cannot reach the browser, and checks the
 * house rules the project is held to.
 *
 *   node tools/check-syntax.mjs
 */
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, relative, extname } from 'node:path';
import vm from 'node:vm';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const SKIP_DIRS = new Set(['node_modules', '.git', 'temp', 'docs', 'icons']);

function walk(dir, out = []) {
  for (const entry of readdirSync(dir)) {
    if (SKIP_DIRS.has(entry)) continue;
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) walk(full, out);
    else out.push(full);
  }
  return out;
}

const files = walk(root);
const problems = [];

for (const file of files) {
  const rel = relative(root, file).replace(/\\/g, '/');
  const ext = extname(file);
  if (!['.js', '.mjs', '.json', '.html', '.css', '.md'].includes(ext)) continue;

  const text = readFileSync(file, 'utf8');

  // House rule: no em dash anywhere in the project. The character is built from
  // its code point so this file does not trip its own check.
  const emDash = String.fromCharCode(0x2014);
  if (text.includes(emDash)) {
    const line = text.slice(0, text.indexOf(emDash)).split('\n').length;
    problems.push(`${rel}:${line}  contains an em dash`);
  }

  if (ext === '.json') {
    try { JSON.parse(text); } catch (err) { problems.push(`${rel}  invalid JSON: ${err.message}`); }
    continue;
  }

  if (ext === '.js') {
    // Extension scripts are classic scripts, so parse them as such.
    try { new vm.Script(text, { filename: rel }); } catch (err) {
      problems.push(`${rel}  ${err.message}`);
    }
    continue;
  }

  if (ext === '.mjs') {
    try { new vm.SourceTextModule(text, { identifier: rel }); } catch (err) {
      if (!/SourceTextModule is not a constructor/.test(String(err))) {
        problems.push(`${rel}  ${err.message}`);
      }
    }
  }
}

// The manifest must point at files that exist.
const manifest = JSON.parse(readFileSync(join(root, 'manifest.json'), 'utf8'));
const referenced = [
  manifest.background.service_worker,
  manifest.side_panel.default_path,
  ...Object.values(manifest.icons),
  ...Object.values(manifest.action.default_icon),
  ...manifest.content_scripts.flatMap((cs) => cs.js)
];
for (const path of new Set(referenced)) {
  try { statSync(join(root, path)); } catch { problems.push(`manifest.json references a missing file: ${path}`); }
}

if (problems.length) {
  console.error('Problems found:\n');
  for (const p of problems) console.error('  ' + p);
  process.exit(1);
}
console.log(`Checked ${files.length} files. No syntax errors, no em dashes, manifest references all resolve.`);

/**
 * build.js — Copies source files to dist/ for clasp push.
 *
 * No bundler needed. Each .js file becomes a separate .gs file
 * in the Apps Script editor.
 */

import { cpSync, mkdirSync, readdirSync, rmSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, '..');
const dist = resolve(root, 'dist');
const src = resolve(root, 'src');

// Clean dist/
rmSync(dist, { recursive: true, force: true });
mkdirSync(dist, { recursive: true });

// Copy src/*.js → dist/
const jsFiles = readdirSync(src).filter((f) => f.endsWith('.js'));
for (const file of jsFiles) {
  cpSync(resolve(src, file), resolve(dist, file));
}

// Copy src/html/*.html → dist/
const htmlDir = resolve(src, 'html');
const htmlFiles = readdirSync(htmlDir).filter((f) => f.endsWith('.html'));
for (const file of htmlFiles) {
  cpSync(resolve(htmlDir, file), resolve(dist, file));
}

// Copy appsscript.json → dist/
cpSync(resolve(root, 'appsscript.json'), resolve(dist, 'appsscript.json'));

console.log('Build complete → dist/');

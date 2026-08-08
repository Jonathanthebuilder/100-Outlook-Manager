import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, '..');
const dist = resolve(root, 'dist');
const htmlPath = resolve(dist, 'outlook-manager.html');
const standalonePath = resolve(dist, 'outlook-manager-standalone.html');

const html = await readFile(htmlPath, 'utf8');
const scriptMatch = html.match(/<script[^>]+src="\.\/([^"]+)"[^>]*><\/script>/);
const styleMatch = html.match(/<link[^>]+href="\.\/([^"]+)"[^>]*>/);

if (!scriptMatch || !styleMatch) {
  throw new Error('Could not find built JS/CSS asset references in dist/outlook-manager.html');
}

const [scriptAsset, styleAsset] = [scriptMatch[1], styleMatch[1]];
const [script, style] = await Promise.all([
  readFile(resolve(dist, scriptAsset), 'utf8'),
  readFile(resolve(dist, styleAsset), 'utf8'),
]);

const standalone = html
  .replace(styleMatch[0], () => `<style>${style}</style>`)
  .replace(scriptMatch[0], '')
  .replace('</body>', () => `    <script>${script}</script>\n  </body>`);

await mkdir(dist, { recursive: true });
await writeFile(standalonePath, standalone, 'utf8');
await writeFile(htmlPath, standalone, 'utf8');

console.log(`Standalone file written: ${standalonePath}`);
console.log(`Offline entry updated: ${htmlPath}`);

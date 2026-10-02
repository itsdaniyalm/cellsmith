// Fills in manifest.template.xml.
// Usage: node scripts/make-manifest.mjs [baseUrl] [outFile]
//   baseUrl  where the built files are served from (default https://localhost:3000)
//   outFile  where to write the manifest (default manifest.xml)
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const baseUrl = (process.argv[2] ?? 'https://localhost:3000').replace(/\/+$/, '');
const outFile = path.resolve(root, process.argv[3] ?? 'manifest.xml');

if (!/^https:\/\//i.test(baseUrl)) {
  console.error(`Office add-ins must be served over HTTPS, but the base URL is "${baseUrl}".`);
  process.exit(1);
}

const xml = readFileSync(path.join(root, 'manifest.template.xml'), 'utf8')
  // The header comment documents the placeholder, so keep it literal there.
  .replace(/<!--[\s\S]*?-->\s*/, '')
  .replaceAll('${BASE_URL}', baseUrl);

mkdirSync(path.dirname(outFile), { recursive: true });
writeFileSync(outFile, xml);
console.log(`wrote ${path.relative(root, outFile)} for ${baseUrl}`);

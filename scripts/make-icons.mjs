// Renders the add-in icons from the SVG masters in assets/.
// Run: npm run icons   (needs the @resvg/resvg-js dev dependency)
//
//   assets/icon.svg        full design, used for 64 px and up
//   assets/icon-small.svg  simplified for 16 and 32 px, where fine detail disappears
//
// Output:
//   public/assets/icon-{16,32,64,80}.png   referenced by the add-in manifest (served with the app)
//   assets/icon-{128,512}.png              for store listings and social previews (repo only)
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { Resvg } from '@resvg/resvg-js';

const root = path.resolve(import.meta.dirname, '..');
const full = readFileSync(path.join(root, 'assets', 'icon.svg'), 'utf8');
const small = readFileSync(path.join(root, 'assets', 'icon-small.svg'), 'utf8');

const jobs = [
  { size: 16, svg: small, dir: 'public/assets' },
  { size: 32, svg: small, dir: 'public/assets' },
  { size: 64, svg: full, dir: 'public/assets' },
  { size: 80, svg: full, dir: 'public/assets' },
  { size: 128, svg: full, dir: 'assets' },
  { size: 512, svg: full, dir: 'assets' },
];

for (const { size, svg, dir } of jobs) {
  const png = new Resvg(svg, { fitTo: { mode: 'width', value: size } }).render().asPng();
  mkdirSync(path.join(root, dir), { recursive: true });
  writeFileSync(path.join(root, dir, `icon-${size}.png`), png);
  console.log(`wrote ${dir}/icon-${size}.png`);
}

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { defineConfig } from 'vitest/config';

/**
 * Office add-ins must be served over HTTPS. `npm run certs` installs a trusted localhost
 * certificate (via office-addin-dev-certs); we use it when present. Set NO_HTTPS=1 to force
 * plain HTTP (handy for previewing the UI in an ordinary browser).
 */
function httpsOptions() {
  if (process.env.NO_HTTPS) return undefined;
  const dir = path.join(os.homedir(), '.office-addin-dev-certs');
  try {
    return {
      key: fs.readFileSync(path.join(dir, 'localhost.key')),
      cert: fs.readFileSync(path.join(dir, 'localhost.crt')),
      ca: fs.readFileSync(path.join(dir, 'ca.crt')),
    };
  } catch {
    return undefined;
  }
}

export default defineConfig({
  base: './',
  resolve: {
    alias: {
      // Monaco's ESM sources, imported selectively (see src/editor/monaco.ts) so none of its
      // 80+ bundled languages end up in the build.
      '@monaco': path.resolve(import.meta.dirname, 'node_modules/monaco-editor/esm/vs'),
    },
  },
  server: {
    host: 'localhost',
    port: 3000,
    strictPort: true,
    https: httpsOptions(),
    headers: { 'Access-Control-Allow-Origin': '*' },
  },
  worker: { format: 'es' },
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 4000,
    rollupOptions: { input: { taskpane: path.resolve(import.meta.dirname, 'taskpane.html') } },
  },
  test: { include: ['tests/**/*.test.ts'] },
});

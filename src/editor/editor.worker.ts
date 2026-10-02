// Monaco's generic editor worker (text diffing, word lookup, etc.). It lives in its own local
// entry so Vite bundles it as a worker; importing it straight out of node_modules through the
// dependency optimizer breaks the `?worker` default export.
import '@monaco/editor/editor.worker.js';

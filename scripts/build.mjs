import { build, context } from 'esbuild';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { chmodSync, readFileSync, writeFileSync } from 'fs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, '..');

// Sync manifest.xml version with package.json
const pkg = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));
const manifestPath = resolve(root, 'manifest.xml');
const manifest = readFileSync(manifestPath, 'utf8');
const updatedManifest = manifest.replace(/<Version>[^<]*<\/Version>/, `<Version>${pkg.version}</Version>`);
if (updatedManifest !== manifest) {
  writeFileSync(manifestPath, updatedManifest);
}

const watch = process.argv.includes('--watch');

const shared = {
  bundle: true,
  sourcemap: true,
  logLevel: 'info',
};

const serverConfig = {
  ...shared,
  entryPoints: [resolve(root, 'src/server/daemon.ts')],
  outfile: resolve(root, 'dist/daemon.js'),
  platform: 'node',
  target: 'node18',
  format: 'cjs',
  banner: { js: '#!/usr/bin/env node' },
  external: ['@modelcontextprotocol/sdk'],
};

const entrypointConfig = {
  ...shared,
  entryPoints: [resolve(root, 'src/server/entrypoint.ts')],
  outfile: resolve(root, 'dist/server.js'),
  platform: 'node',
  target: 'node18',
  format: 'cjs',
  banner: { js: '#!/usr/bin/env node' },
  external: [],
};

const taskpaneConfig = {
  ...shared,
  entryPoints: [resolve(root, 'src/taskpane/index.ts')],
  outfile: resolve(root, 'dist/taskpane-app.js'),
  platform: 'browser',
  target: 'es2020',
  format: 'iife',
  globalName: 'TaskpaneApp',
  // Bake the package version into the bundle so the taskpane can display it.
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
};

if (watch) {
  const serverCtx = await context(serverConfig);
  const entrypointCtx = await context(entrypointConfig);
  const taskpaneCtx = await context(taskpaneConfig);
  await serverCtx.watch();
  await entrypointCtx.watch();
  await taskpaneCtx.watch();
  console.log('Watching for changes...');
} else {
  await build(serverConfig);
  chmodSync(resolve(root, 'dist/daemon.js'), 0o755);
  await build(entrypointConfig);
  chmodSync(resolve(root, 'dist/server.js'), 0o755);
  await build(taskpaneConfig);
}

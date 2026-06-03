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
  entryPoints: [resolve(root, 'src/server/index.ts')],
  outfile: resolve(root, 'dist/server.js'),
  platform: 'node',
  target: 'node18',
  format: 'cjs',
  banner: { js: '#!/usr/bin/env node' },
  external: ['@modelcontextprotocol/sdk'],
};

const taskpaneConfig = {
  ...shared,
  entryPoints: [resolve(root, 'src/taskpane/index.ts')],
  outfile: resolve(root, 'dist/taskpane-app.js'),
  platform: 'browser',
  target: 'es2020',
  format: 'iife',
  globalName: 'TaskpaneApp',
};

if (watch) {
  const serverCtx = await context(serverConfig);
  const taskpaneCtx = await context(taskpaneConfig);
  await serverCtx.watch();
  await taskpaneCtx.watch();
  console.log('Watching for changes...');
} else {
  await build(serverConfig);
  chmodSync(resolve(root, 'dist/server.js'), 0o755);
  await build(taskpaneConfig);
}

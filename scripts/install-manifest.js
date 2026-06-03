#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const os = require('os');

const manifestSrc = path.join(__dirname, '..', 'manifest.xml');

if (process.platform === 'darwin') {
  const wefDir = path.join(os.homedir(), 'Library/Containers/com.microsoft.Word/Data/Documents/wef');
  if (!fs.existsSync(wefDir)) fs.mkdirSync(wefDir, { recursive: true });
  const dest = path.join(wefDir, 'manifest.xml');
  fs.copyFileSync(manifestSrc, dest);
  console.log('✓ Manifest installed to: ' + dest);
  console.log('  Restart Word, then: Home → Add-ins → MCP Word Bridge');
} else if (process.platform === 'win32') {
  const wefDir = path.join(os.homedir(), 'AppData', 'Local', 'Microsoft', 'Office', '16.0', 'Wef');
  try {
    if (!fs.existsSync(wefDir)) fs.mkdirSync(wefDir, { recursive: true });
    const dest = path.join(wefDir, 'manifest.xml');
    fs.copyFileSync(manifestSrc, dest);
    console.log('✓ Manifest installed to: ' + dest);
    console.log('  Restart Word, then: Insert → My Add-ins → Developer Add-ins');
  } catch {
    console.log('Manual steps: Copy ' + manifestSrc + ' to a local folder and add as Trusted Catalog in Word');
  }
} else {
  console.log('Platform not supported for local sideloading. Manifest: ' + manifestSrc);
}

// Starts ALL THREE websites: Operator (4001), Enterprise (4002), Customer phone (4003). Press Ctrl+C to stop both.
const { spawn } = require('child_process');
const path = require('path');
const kids = ['operator', 'enterprise', 'customer'].map(role => spawn(process.execPath, [path.join(__dirname, 'server.js'), role], { stdio: 'inherit' }));
const stop = () => kids.forEach(k => { try { k.kill(); } catch (e) {} });
process.on('SIGINT', () => { stop(); process.exit(0); });
process.on('SIGTERM', () => { stop(); process.exit(0); });
kids.forEach(k => k.on('exit', code => { if (code) { stop(); process.exit(code); } }));

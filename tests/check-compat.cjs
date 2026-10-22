const fs = require('fs');
const path = require('path');
const assert = require('assert');
const acorn = require('acorn');

const dist = path.resolve(__dirname, '../dist');
const bundle = fs.readFileSync(path.join(dist, 'assets/js/office-compat.bundle.js'), 'utf8');
const html = fs.readFileSync(path.join(dist, 'taskpane.html'), 'utf8');

assert(bundle.length > 10000, 'Expected application and polyfills in bundle');
acorn.parse(bundle, { ecmaVersion: 5, sourceType: 'script' });
assert(html.includes('Office.onReady(function (info)'), 'Missing early Office onReady callback');
assert(html.includes('assets/js/office-compat.bundle.js'), 'Missing bundle reference');
assert(!html.includes('assets/js/app.js'), 'Untranspiled app.js referenced');
assert(!html.includes('legacy-message'), 'Unconditional legacy warning remains');
assert(fs.existsSync(path.join(dist, 'assets/js/commands.js')), 'Commands JS missing');
assert(fs.existsSync(path.join(dist, 'manifest.excel.prod.xml')), 'Manifest missing');
assert(fs.existsSync(path.join(dist, 'assets/logo.png')), 'Logo missing');
const ieCss = fs.readFileSync(path.join(dist, 'assets/css/app.ie11.css'), 'utf8');
assert(!ieCss.includes('var(--'), 'IE11 CSS still contains unresolved variables');
assert(ieCss.includes('[data-theme="dark"] .app'), 'Dark-mode legacy styles missing');
assert(html.includes('assets/css/app.ie11.css'), 'Missing runtime legacy CSS loader');
console.log('PASS: ES5 syntax and Office compatibility distribution checks');

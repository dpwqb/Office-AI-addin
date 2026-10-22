const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const config = require('../webpack.config.cjs');

const ROOT = path.resolve(__dirname, '..');
const DIST = path.join(ROOT, 'dist');
const SCRIPT_NAMES = [
  'config', 'state', 'markdown', 'host', 'host-excel',
  'host-word', 'host-powerpoint', 'api', 'ui', 'app'
];

function copyIfPresent(relativePath) {
  const from = path.join(ROOT, relativePath);
  const to = path.join(DIST, relativePath);
  if (!fs.existsSync(from)) return;
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.cpSync(from, to, { recursive: true });
}

function buildHtml() {
  const source = fs.readFileSync(path.join(ROOT, 'taskpane.html'), 'utf8');
  const originalScripts = SCRIPT_NAMES
    .map(name => '<script src="assets/js/' + name + '.js"></script>')
    .join('\n  ');
  if (!source.includes(originalScripts)) {
    throw new Error('Expected application script sequence not found in taskpane.html');
  }
  const officeJsTag =
    '<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>';
  if (!source.includes(officeJsTag)) {
    throw new Error('Office.js bootstrap script not found');
  }
  let html = source.replace(
    originalScripts,
    '<script src="assets/js/office-compat.bundle.js"></script>'
  );
  if (!html.includes('window.__officeReadyInfo')) {
    html = html.replace(officeJsTag, officeJsTag + '\n  <script>\n' +
      '    // Office 2019: register Office.onReady before any application bundle.\n' +
      '    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {\n' +
      '      Office.onReady(function (info) { window.__officeReadyInfo = info; });\n' +
      '    }\n' +
      '  </script>');
  }
  html = html.replace(/\s*<div id="legacy-message"[^>]*>[\s\S]*?<\/div>/, '');
  fs.writeFileSync(path.join(DIST, 'taskpane.html'), html);
}

fs.rmSync(DIST, { recursive: true, force: true });
webpack(config, (error, stats) => {
  if (error || !stats || stats.hasErrors()) {
    console.error(error || stats.toString({ all: false, errors: true }));
    process.exitCode = 1;
    return;
  }
  try {
    ['assets/css', 'assets/logo.png', 'assets/vendor',
      'commands.html', 'assets/js/commands.js', 'dpoqb-in-office.html', 'CNAME'
    ].forEach(copyIfPresent);
    fs.readdirSync(ROOT).filter(name => /^manifest\..+\.xml$/.test(name))
      .forEach(copyIfPresent);
    buildHtml();
    console.log('Office compatibility site built in dist/. Serve dist/ as the web root.');
    console.log(stats.toString({ all: false, assets: true, timings: true }));
  } catch (e) {
    console.error(e);
    process.exitCode = 1;
  }
});

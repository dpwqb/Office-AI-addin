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
  const cssTag = '<link rel="stylesheet" href="assets/css/app.css" />';
  if (!html.includes(cssTag)) throw new Error('Stylesheet reference not found');
  const cssBootstrap =
    '\n  <script>\n' +
    '    // IE11 does not support CSS custom properties (var(--...)).\n' +
    '    if (!window.CSS || !window.CSS.supports || !window.CSS.supports("color", "var(--test)")) {\n' +
    '      var sheet = document.createElement("link");\n' +
    '      sheet.rel = "stylesheet"; sheet.href = "assets/css/app.ie11.css";\n' +
    '      document.getElementsByTagName("head")[0].appendChild(sheet);\n' +
    '    }\n' +
    '  </script>';
  html = html.replace(cssTag, cssTag + cssBootstrap);
  html = html.replace(/\s*<div id="legacy-message"[^>]*>[\s\S]*?<\/div>/, '');
  fs.writeFileSync(path.join(DIST, 'taskpane.html'), html);
}

function buildLegacyCss() {
  const source = fs.readFileSync(path.join(ROOT, 'assets/css/app.css'), 'utf8');
  const rootBlock = source.match(/:root\s*\{([^}]*)\}/);
  const darkBlock = source.match(/\[data-theme="dark"\]\s*\{([^}]*)\}/);
  if (!rootBlock || !darkBlock) throw new Error('Missing light or dark CSS variables');
  function variables(block, fallback) {
    const values = Object.assign({}, fallback || {});
    block.replace(/(--[\w-]+)\s*:\s*([^;]+);/g, function (_, key, value) {
      values[key] = value.trim();
      return '';
    });
    return values;
  }
  const light = variables(rootBlock[1]);
  const dark = variables(darkBlock[1], light);
  function resolve(css, values) {
    return css.replace(/var\(\s*(--[\w-]+)\s*\)/g, function (match, key) {
      return Object.prototype.hasOwnProperty.call(values, key) ? values[key] : match;
    });
  }
  const darkRules = [];
  const ruleRe = /([^{}]+)\{([^{}]*)\}/g;
  let rule;
  while ((rule = ruleRe.exec(source)) !== null) {
    const selector = rule[1].trim();
    if (!rule[2].includes('var(') || selector[0] === '@' ||
        selector === ':root' || selector === '[data-theme="dark"]') continue;
    const scoped = selector.split(',').map(function (part) {
      return '[data-theme="dark"] ' + part.trim();
    }).join(', ');
    darkRules.push(scoped + ' {' + resolve(rule[2], dark) + '}');
  }
  const legacy = resolve(source, light) +
    '\n/* IE11 dark-mode overrides (CSS custom properties are unavailable). */\n' +
    darkRules.join('\n');
  if (legacy.includes('var(--')) throw new Error('Unresolved IE11 CSS variable');
  fs.writeFileSync(path.join(DIST, 'assets/css/app.ie11.css'), legacy);
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
    buildLegacyCss();
    buildHtml();
    console.log('Office compatibility site built in dist/. Serve dist/ as the web root.');
    console.log(stats.toString({ all: false, assets: true, timings: true }));
  } catch (e) {
    console.error(e);
    process.exitCode = 1;
  }
});

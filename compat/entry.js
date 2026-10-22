// Ordered classic-script bundle; all app modules use window.App.
require('core-js/stable');
require('regenerator-runtime/runtime');
require('whatwg-fetch');
require('abortcontroller-polyfill/dist/abortcontroller-polyfill-only');

require('../assets/js/config.js');
require('../assets/js/state.js');
require('../assets/js/markdown.js');
require('../assets/js/host.js');
require('../assets/js/host-excel.js');
require('../assets/js/host-word.js');
require('../assets/js/host-powerpoint.js');
require('../assets/js/api.js');
require('../assets/js/ui.js');
require('../assets/js/app.js');

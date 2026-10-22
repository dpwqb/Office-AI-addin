// Ordered classic-script bundle; all app modules use window.App.
require('core-js/stable');
require('regenerator-runtime/runtime');
require('whatwg-fetch');
require('abortcontroller-polyfill/dist/abortcontroller-polyfill-only');

// IE11 lacks Element.closest; delegated UI click handlers depend on it.
if (typeof Element !== 'undefined') {
  if (!Element.prototype.matches) {
    Element.prototype.matches = Element.prototype.msMatchesSelector ||
      Element.prototype.webkitMatchesSelector || function (selector) {
        var all = (this.document || this.ownerDocument).querySelectorAll(selector);
        var index = 0;
        while (all[index] && all[index] !== this) index++;
        return !!all[index];
      };
  }
  if (!Element.prototype.closest) {
    Element.prototype.closest = function (selector) {
      var element = this;
      while (element && element.nodeType === 1) {
        if (element.matches(selector)) return element;
        element = element.parentElement || element.parentNode;
      }
      return null;
    };
  }
}

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

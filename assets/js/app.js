(function () {
  'use strict';
  const App = (window.App = window.App || {});

  async function initOffice() {
    if (navigator.userAgent.indexOf('Trident') !== -1 || navigator.userAgent.indexOf('Edge') !== -1) {
      const legacy = document.getElementById('legacy-message');
      if (legacy) legacy.hidden = false;
    }
    if (typeof Office !== 'undefined' && Office.onReady) {
      try {
        await Office.onReady();
        if (App.hasOffice()) {
          App.state.workbookId = await App.getWorkbookId();
          const md = await App.getWorkbookMetadata().catch(() => null);
          if (md) App.state.workbookLabel = md.workbookName || ''; // 不再显示当前选择范围，避免展示过期状态
        }
      } catch (e) {
        console.warn('[Office init]', e);
      }
    }
    App.ensureSession();
    App.render();
  }

  window.dpoqbExcelTools = App.TOOL_EXECUTORS;
  initOffice();
})();

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
        const info = await Office.onReady();
        // 优先用 Office 报告的宿主类型选择提供者，回退到运行时探测。
        const hostType = App.hostTypeFromOffice(info && info.host) || App.detectHostType();
        App.selectHost(hostType);
        if (App.hasOffice()) {
          App.state.workbookId = await App.getDocumentId();
          const md = await App.host.getMetadata().catch(() => null);
          // 仅 Excel 提供 workbookName；其余宿主保持空标签，避免展示过期状态。
          if (md && md.workbookName) App.state.workbookLabel = md.workbookName;
        }
      } catch (e) {
        console.warn('[Office init]', e);
      }
    }
    window.dpoqbExcelTools = App.host.toolExecutors;
    App.ensureSession();
    App.render();
  }

  initOffice();
})();

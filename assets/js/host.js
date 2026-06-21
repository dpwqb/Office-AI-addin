(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 宿主提供者注册表。各 host-*.js 在加载时把自己注册进来。
  App.HOSTS = App.HOSTS || {};

  // ---- 通用 Office 文档设置存取（宿主无关，Excel/Word/PowerPoint 通用）----
  function saveDocSetting(key, value) {
    return new Promise((resolve, reject) => {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message || 'Failed to save document setting')));
    });
  }
  function loadDocSetting(key, fallback) {
    try { return Office.context.document.settings.get(key) || fallback; } catch { return fallback; }
  }
  async function getDocumentId() {
    if (!App.hasOffice()) return 'browser';
    return new Promise((resolve, reject) => {
      const settings = Office.context.document.settings;
      let v = settings.get(App.STORAGE_KEYS.workbookId);
      if (v) return resolve(v);
      v = App.id();
      settings.set(App.STORAGE_KEYS.workbookId, v);
      settings.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(v) : reject(new Error(r.error?.message || 'Failed to save document ID')));
    });
  }
  function requireOffice() { if (!App.hasOffice()) throw new Error(App.t('demo')); }

  // ---- 非 Office 环境 / 未识别宿主时的占位提供者 ----
  function makeDemoHost(hostType) {
    const unavailable = async () => { throw new Error(App.t('demo')); };
    return {
      hostType: hostType || 'none',
      available: false,
      metadataLabel: 'Document metadata',
      systemPrompt: 'You are an AI assistant integrated into Microsoft Office.',
      toolDefinitions: [],
      toolExecutors: {},
      defaultArgsForTool() { return '{}'; },
      evalToolName: 'eval_officejs',
      getMetadata: unavailable,
      navigateCitation: unavailable,
      follow: async () => {},
      i18n: { zh: {}, en: {} }
    };
  }

  // ---- 宿主检测与选择 ----
  function detectHostType() {
    if (typeof Excel !== 'undefined' && Excel.run) return 'excel';
    if (typeof Word !== 'undefined' && Word.run) return 'word';
    if (typeof PowerPoint !== 'undefined' && PowerPoint.run) return 'powerpoint';
    return null;
  }
  function hostTypeFromOffice(officeHost) {
    if (typeof Office === 'undefined' || !Office.HostType) return null;
    switch (officeHost) {
      case Office.HostType.Excel: return 'excel';
      case Office.HostType.Word: return 'word';
      case Office.HostType.PowerPoint: return 'powerpoint';
      default: return null;
    }
  }
  function selectHost(type) {
    const provider = type && App.HOSTS[type];
    App.host = provider || makeDemoHost(type);
    return App.host;
  }

  // ---- 稳定的委托封装：api.js / ui.js 在加载时取这些别名，调用时再转发到当前宿主 ----
  function executeToolByName(name, args) {
    const ex = App.host && App.host.toolExecutors;
    const fn = ex && ex[name];
    if (!fn) throw new Error(`Tool ${name} not found`);
    return fn(args || {});
  }
  function maybeFollow(result) { return App.host && App.host.follow ? App.host.follow(result) : Promise.resolve(); }
  function navigateCitation(ref) { return App.host && App.host.navigateCitation ? App.host.navigateCitation(ref) : Promise.resolve(); }
  function getWorkbookMetadata() { return App.host.getMetadata(); }

  App.saveDocSetting = saveDocSetting;
  App.loadDocSetting = loadDocSetting;
  App.getDocumentId = getDocumentId;
  App.getWorkbookId = getDocumentId; // 向后兼容旧调用名
  App.requireOffice = requireOffice;
  App.makeDemoHost = makeDemoHost;
  App.detectHostType = detectHostType;
  App.hostTypeFromOffice = hostTypeFromOffice;
  App.selectHost = selectHost;
  App.executeToolByName = executeToolByName;
  App.maybeFollow = maybeFollow;
  App.navigateCitation = navigateCitation;
  App.getWorkbookMetadata = getWorkbookMetadata;

  // 默认先挂占位提供者，保证 App.host 在首屏渲染前一定存在；app.js 会在 Office.onReady 后重新选定。
  App.host = makeDemoHost(detectHostType());
})();

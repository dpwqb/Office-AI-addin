(function () {
  'use strict';
  const App = (window.App = window.App || {});

  function getStoredItem(key) { try { return (typeof window !== 'undefined' && window.localStorage) ? window.localStorage.getItem(key) : null; } catch { return null; } }
  function setStoredItem(key, value) { try { if (typeof window !== 'undefined' && window.localStorage) window.localStorage.setItem(key, value); } catch {} }
  function prefersDarkMode() { return typeof window !== 'undefined' && typeof window.matchMedia === 'function' && window.matchMedia('(prefers-color-scheme: dark)').matches; }
  function getNavigatorLanguage() { return (typeof navigator !== 'undefined' && navigator.language) ? navigator.language : 'en'; }

  function t(key) {
    const hostI18n = App.host && App.host.i18n && App.host.i18n[state.locale];
    if (hostI18n && hostI18n[key] != null) return hostI18n[key];
    return (App.I18N[state.locale] && App.I18N[state.locale][key]) || key;
  }
  function loadLocale() { const v = getStoredItem(App.STORAGE_KEYS.locale); return v === 'en' || v === 'zh' ? v : (getNavigatorLanguage().startsWith('zh') ? 'zh' : 'en'); }
  function loadSettings() { try { return Object.assign({}, App.DEFAULT_SETTINGS, JSON.parse(getStoredItem(App.STORAGE_KEYS.settings) || '{}')); } catch { return { ...App.DEFAULT_SETTINGS }; } }
  function saveSettings() { setStoredItem(App.STORAGE_KEYS.settings, JSON.stringify(state.settings)); }
  function loadSessions() { try { return JSON.parse(getStoredItem(App.STORAGE_KEYS.sessions) || '[]'); } catch { return []; } }
  function saveSessions() { setStoredItem(App.STORAGE_KEYS.sessions, JSON.stringify(state.sessions)); }
  function id() { const c = typeof crypto !== 'undefined' ? crypto : null; return (c && typeof c.randomUUID === 'function') ? c.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`; }
  function now() { return Date.now(); }
  function hasOffice() {
    if (typeof Office === 'undefined') return false;
    const runtimeReady = (typeof Excel !== 'undefined' && !!Excel.run) || (typeof Word !== 'undefined' && !!Word.run) || (typeof PowerPoint !== 'undefined' && !!PowerPoint.run);
    // 同时要求已选定一个可用的宿主提供者，避免 provider 选定失败时 UI 仍呈现为可用 Office 会话。
    return runtimeReady && !!(App.host && App.host.available);
  }
  function escapeHtml(s) { return String(s ?? '').replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch])); }
  function pretty(v) { try { return JSON.stringify(v, null, 2); } catch { return String(v); } }
  function normalizeBaseUrl(url) { return String(url || '').replace(/\/+$/, ''); }
  function chatEndpoint() { const base = normalizeBaseUrl(state.settings.customPrefixUrl || App.DEFAULT_BASE_URL); return base.endsWith('/chat/completions') ? base : `${base}/chat/completions`; }

  function getProviderModels(provider) {
    const p = App.PROVIDERS[provider] || App.PROVIDERS.Custom;
    const models = p.models.slice();
    const current = state.settings.model;
    if (current && !models.some(m => m.id === current)) models.unshift({ id: current, name: current });
    return models;
  }

  const state = {
    locale: loadLocale(),
    theme: getStoredItem(App.STORAGE_KEYS.theme) || (prefersDarkMode() ? 'dark' : 'light'),
    tab: 'chat',
    settings: loadSettings(),
    workbookId: null,
    sessions: loadSessions(),
    currentSessionId: null,
    messages: [],
    isWorking: false,
    stopRequested: false,
    abortController: null,
    error: null,
    stats: { inputTokens: 0, outputTokens: 0, totalCost: 0, calls: 0 },
    toolOutput: '',
    workbookLabel: '',
    sessionMenuOpen: false,
    pendingDeleteSessionId: null,
    pendingEval: null
  };

  function activeWorkbookKey() { return state.workbookId || 'browser'; }
  function sessionsForCurrentWorkbook() {
    const key = activeWorkbookKey();
    return state.sessions.filter(s => (s.workbookId || 'browser') === key);
  }
  function makeSession(name) {
    return { id: id(), workbookId: activeWorkbookKey(), name: name || (state.locale === 'zh' ? '新对话' : 'NEW CHAT'), messages: [], createdAt: now(), updatedAt: now() };
  }
  function ensureSession() {
    const available = sessionsForCurrentWorkbook();
    if (state.currentSessionId && available.some(s => s.id === state.currentSessionId)) return;
    let session = available[0];
    if (!session) { session = makeSession(); state.sessions.unshift(session); saveSessions(); }
    state.currentSessionId = session.id;
    state.messages = session.messages || [];
  }
  function persistCurrentSession() {
    const s = state.sessions.find(x => x.id === state.currentSessionId);
    if (!s) return;
    s.messages = state.messages;
    s.updatedAt = now();
    const firstUser = state.messages.find(m => m.role === 'user');
    if (firstUser && (!s.name || s.name === '新对话' || s.name === 'NEW CHAT')) {
      const text = firstUser.content.trim();
      s.name = text.length > 40 ? text.slice(0, 37) + '...' : text;
    }
    state.sessions.sort((a, b) => b.updatedAt - a.updatedAt);
    saveSessions();
  }

  App.state = state;
  App.getStoredItem = getStoredItem;
  App.setStoredItem = setStoredItem;
  App.prefersDarkMode = prefersDarkMode;
  App.getNavigatorLanguage = getNavigatorLanguage;
  App.t = t;
  App.loadLocale = loadLocale;
  App.loadSettings = loadSettings;
  App.saveSettings = saveSettings;
  App.loadSessions = loadSessions;
  App.saveSessions = saveSessions;
  App.id = id;
  App.now = now;
  App.hasOffice = hasOffice;
  App.escapeHtml = escapeHtml;
  App.pretty = pretty;
  App.normalizeBaseUrl = normalizeBaseUrl;
  App.chatEndpoint = chatEndpoint;
  App.getProviderModels = getProviderModels;
  App.activeWorkbookKey = activeWorkbookKey;
  App.sessionsForCurrentWorkbook = sessionsForCurrentWorkbook;
  App.makeSession = makeSession;
  App.ensureSession = ensureSession;
  App.persistCurrentSession = persistCurrentSession;
})();

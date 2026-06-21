(function () {
  'use strict';
  const App = (window.App = window.App || {});

  // 本模块在其余脚本之后加载，取别名时这些成员均已挂载到 App。
  const state = App.state;
  const t = App.t;
  const escapeHtml = App.escapeHtml;
  const pretty = App.pretty;
  const now = App.now;
  const hasOffice = App.hasOffice;
  const PROVIDERS = App.PROVIDERS;
  const STORAGE_KEYS = App.STORAGE_KEYS;
  const TOOL_DEFINITIONS = App.TOOL_DEFINITIONS;
  const defaultArgsForTool = App.defaultArgsForTool;
  const getProviderModels = App.getProviderModels;
  const sessionsForCurrentWorkbook = App.sessionsForCurrentWorkbook;
  const renderMarkdown = App.renderMarkdown;
  const trimToolText = App.trimToolText;
  const makeSession = App.makeSession;
  const saveSessions = App.saveSessions;
  const ensureSession = App.ensureSession;
  const persistCurrentSession = App.persistCurrentSession;
  const saveSettings = App.saveSettings;
  const setStoredItem = App.setStoredItem;
  const runAgentLoop = App.runAgentLoop;
  const markStoppedMessage = App.markStoppedMessage;
  const executeToolByName = App.executeToolByName;
  const maybeFollow = App.maybeFollow;
  const navigateCitation = App.navigateCitation;

  function captureUiState() {
    const snap = { focused: null, atBottom: true, scrollTop: 0 };
    try {
      const el = document.activeElement;
      if (el && (el.id === 'chat-input' || el.id === 'manual-args')) {
        snap.focused = { id: el.id, value: el.value, start: el.selectionStart, end: el.selectionEnd };
      }
      const msgEl = document.getElementById('messages');
      if (msgEl) {
        snap.scrollTop = msgEl.scrollTop;
        snap.atBottom = msgEl.scrollHeight - msgEl.scrollTop - msgEl.clientHeight < 40;
      }
    } catch {}
    return snap;
  }
  function restoreUiState(snap) {
    if (!snap) return;
    if (snap.focused) {
      const el = document.getElementById(snap.focused.id);
      if (el) {
        el.value = snap.focused.value;
        el.focus();
        try { el.setSelectionRange(snap.focused.start ?? el.value.length, snap.focused.end ?? el.value.length); } catch {}
      }
    }
    const msgEl = document.getElementById('messages');
    if (msgEl) msgEl.scrollTop = snap.atBottom ? msgEl.scrollHeight : snap.scrollTop;
  }

  function render() {
    const snap = captureUiState();
    document.documentElement.dataset.theme = state.theme;
    ensureSession();
    const configured = Boolean(state.settings.apiKey && state.settings.model && state.settings.customPrefixUrl);
    const currentProvider = PROVIDERS[state.settings.provider] || PROVIDERS.Custom;
    const models = getProviderModels(state.settings.provider);
    const visibleSessions = sessionsForCurrentWorkbook();
    const sessionsHtml = visibleSessions.length ? visibleSessions.map(s => {
      const pendingDelete = state.pendingDeleteSessionId === s.id;
      return `
      <div class="session-item ${s.id === state.currentSessionId ? 'active' : ''} ${pendingDelete ? 'confirming-delete' : ''}">
        <button data-action="switch-session" data-id="${s.id}" title="${escapeHtml(s.name)}" ${pendingDelete ? 'disabled' : ''}><span>${escapeHtml(s.name)}</span><small>${new Date(s.updatedAt).toLocaleString()}</small></button>
        ${pendingDelete ? `<div class="delete-confirm"><small>${t('confirmDeleteSession')}</small><div><button class="confirm-delete-btn" data-action="confirm-delete-session" data-id="${s.id}">${t('confirm')}</button><button class="cancel-delete-btn" data-action="cancel-delete-session" data-id="${s.id}">${t('cancel')}</button></div></div>` : `<button class="delete-btn" data-action="delete-session" data-id="${s.id}" title="${t('deleteSession')}">×</button>`}
      </div>`;
    }).join('') : `<div class="session-empty">${t('noSessions')}</div>`;
    const messagesHtml = state.messages.length ? state.messages.map(renderMessage).join('') : `
      <div class="empty">
        <img class="logo" src="assets/logo.png" alt="logo" />
        <h2>${t('title')}</h2>
        <p>${t('subtitle')}</p>
        <div class="prompt-grid">
          <button class="prompt-card" data-prompt="${escapeHtml(state.locale === 'zh' ? '请根据当前表格数据结构，生成合适的可视化图表' : 'Generate charts to visualize my data and apply professional styling')}"><strong>${t('chart')}</strong><span>${t('chartDesc')}</span></button>
          <button class="prompt-card" data-prompt="${escapeHtml(state.locale === 'zh' ? '帮我检查当前表格中的错误内容，定位问题并给出修复后的正确结果' : 'Check my spreadsheet for formula errors and fix them automatically')}"><strong>${t('fix')}</strong><span>${t('fixDesc')}</span></button>
          <button class="prompt-card" data-prompt="${escapeHtml(state.locale === 'zh' ? '帮我全面分析表中所有内容，并输出汇总结果与关键分析结论' : 'Analyze all workbook contents and summarize key conclusions')}"><strong>${t('analyze')}</strong><span>${t('analyzeDesc')}</span></button>
        </div>
      </div>`;
    document.getElementById('container').innerHTML = `
      <div class="app">
        <header class="header">
          <img class="logo" src="assets/logo.png" alt="logo" />
          <div class="title"><strong>dpoqb in Excel</strong><span>${escapeHtml(state.workbookLabel || (configured ? t('configuredShort') : t('notConfiguredShort')))}</span></div>
          <div class="session-menu-wrap">
            <button class="session-trigger" data-action="toggle-session-menu" title="${t('sessionHistory')}"><span>${t('sessions')}</span><strong>⌄</strong></button>
            ${state.sessionMenuOpen ? `<div class="session-menu">
              <div class="session-menu-head"><strong>${t('sessionHistory')}</strong><button class="secondary-btn compact" data-action="new-chat">＋ ${t('newChat')}</button></div>
              <div class="session-list header-session-list">${sessionsHtml}</div>
            </div>` : ''}
          </div>
          <button class="icon-btn" data-action="toggle-follow" title="${state.settings.followMode ? t('followOn') : t('followOff')}">${state.settings.followMode ? '◎' : '○'}</button>
          <button class="icon-btn" data-action="toggle-theme" title="${state.theme === 'dark' ? t('light') : t('dark')}">${state.theme === 'dark' ? '☀' : '☾'}</button>
          <button class="icon-btn" data-action="toggle-locale" title="language">${t('lang')}</button>
        </header>
        <nav class="tabs">
          <button class="tab ${state.tab === 'chat' ? 'active' : ''}" data-tab="chat">${t('chat')}</button>
          <button class="tab ${state.tab === 'settings' ? 'active' : ''}" data-tab="settings">${t('settings')}</button>
          <button class="tab ${state.tab === 'tools' ? 'active' : ''}" data-tab="tools">${t('tools')}</button>
        </nav>
        <main class="main">
          <section class="panel chat-panel ${state.tab === 'chat' ? 'active' : ''}">
            <div class="messages" id="messages">${messagesHtml}</div>
            <div class="composer">
              ${state.error ? `<div class="error">${escapeHtml(state.error)}</div>` : ''}
              <div class="input-wrap">
                <textarea id="chat-input" rows="2" placeholder="${configured ? t('input') : t('noConfig')}" ${state.isWorking || !configured ? 'disabled' : ''}></textarea>
                <button class="send-btn" data-action="send" ${!configured ? 'disabled' : ''}>${state.isWorking ? t('stop') : t('send')}</button>
              </div>
            </div>
          </section>
          <section class="panel settings-panel ${state.tab === 'settings' ? 'active' : ''}">
            <div class="section">
              <h3>API Configuration</h3>
              <div class="field"><label>${t('provider')}</label><select data-bind="provider">${Object.entries(PROVIDERS).map(([k, p]) => `<option value="${k}" ${state.settings.provider === k ? 'selected' : ''}>${escapeHtml(p.label)}</option>`).join('')}</select></div>
              <div class="field"><label>${t('env')}</label><select data-action="pick-base-url">${currentProvider.baseUrls.map(x => `<option value="${escapeHtml(x.value)}" ${state.settings.customPrefixUrl === x.value ? 'selected' : ''}>${escapeHtml(x.label || x.value)}</option>`).join('')}<option value="__custom__">自定义 / Custom</option></select></div>
              <div class="field"><label>${t('baseUrl')}</label><input data-bind="customPrefixUrl" value="${escapeHtml(state.settings.customPrefixUrl || '')}" placeholder="https://.../v1" /></div>
              <div class="field"><label>${t('model')}</label><select data-action="pick-model">${models.map(m => `<option value="${escapeHtml(m.id)}" ${state.settings.model === m.id ? 'selected' : ''}>${escapeHtml(m.name || m.id)}</option>`).join('')}<option value="__custom__">自定义模型 / Custom model</option></select><input data-bind="model" value="${escapeHtml(state.settings.model || '')}" placeholder="model id" /></div>
              <div class="field"><label>${t('apiKey')}</label><input data-bind="apiKey" value="${escapeHtml(state.settings.apiKey || '')}" type="password" placeholder="sk-..." /></div>
              <div class="field"><label>${t('thinking')}</label><select data-bind="thinking"><option value="none" ${state.settings.thinking === 'none' ? 'selected' : ''}>off</option><option value="low" ${state.settings.thinking === 'low' ? 'selected' : ''}>low</option><option value="medium" ${state.settings.thinking === 'medium' ? 'selected' : ''}>medium</option><option value="high" ${state.settings.thinking === 'high' ? 'selected' : ''}>high</option></select></div>
              <label class="check"><input type="checkbox" data-bind="followMode" ${state.settings.followMode ? 'checked' : ''}/> ${t('followOn')}</label>
              <div class="hint">${configured ? t('configured') : t('notConfigured')}</div>
            </div>
            <div class="section"><h3>${t('usage')}</h3><div class="stats"><div class="stat"><span>Input</span><strong>${state.stats.inputTokens}</strong></div><div class="stat"><span>Output</span><strong>${state.stats.outputTokens}</strong></div><div class="stat"><span>Calls</span><strong>${state.stats.calls}</strong></div><div class="stat"><span>Cost</span><strong>${state.stats.totalCost.toFixed(6)}</strong></div></div></div>
            <div class="section"><h3>${t('about')}</h3><p class="hint">${t('aboutText')}</p>${hasOffice() ? '' : `<p class="hint">${t('demo')}</p>`}</div>
          </section>
          <section class="panel tools-panel ${state.tab === 'tools' ? 'active' : ''}">
            <div class="section tool-runner">
              <h3>${t('manualTool')}</h3>
              <div class="field"><label>${t('toolName')}</label><select id="manual-tool">${TOOL_DEFINITIONS.map(x => `<option value="${x.function.name}">${x.function.name}</option>`).join('')}</select></div>
              <div class="field"><label>${t('params')}</label><textarea id="manual-args">${escapeHtml(defaultArgsForTool(TOOL_DEFINITIONS[0].function.name))}</textarea></div>
              <button class="primary-btn" data-action="run-tool">${t('run')}</button>
              <pre class="output">${escapeHtml(state.toolOutput || t('output'))}</pre>
            </div>
          </section>
        </main>
        <footer class="footer"><span>dpoqb in Excel · Plain Edition</span><span>${escapeHtml(state.settings.model || '')}</span></footer>
        ${state.pendingEval ? `<div class="modal-overlay">
          <div class="modal">
            <div class="modal-title">${t('confirmEval')}</div>
            <pre class="modal-code">${escapeHtml(String(state.pendingEval.code || '').slice(0, 4000))}</pre>
            <div class="modal-actions">
              <button class="secondary-btn" data-action="eval-deny">${t('evalDeny')}</button>
              <button class="primary-btn" data-action="eval-allow">${t('evalAllow')}</button>
            </div>
          </div>
        </div>` : ''}
      </div>`;
    restoreUiState(snap);
  }

  function assistantBubbleHtml(m) {
    let body = m.content ? `<div class="markdown">${renderMarkdown(m.content)}</div>` : '';
    if (m.toolCalls && m.toolCalls.length) {
      body += m.toolCalls.map(tc => `<div class="tool-block"><div class="tool-head"><span>${escapeHtml(tc.name)}</span><span class="status ${tc.status || 'running'}">${escapeHtml(tc.status || 'running')}</span></div><div class="tool-body">${escapeHtml(pretty(tc.args))}${tc.result ? '\n\n' + escapeHtml(trimToolText(tc.result)) : ''}</div></div>`).join('');
    }
    return body || '...';
  }

  function renderMessage(m) {
    if (m.role === 'tool') {
      return `<div class="msg tool"><div class="bubble">${escapeHtml(m.name || 'tool')}\n${escapeHtml(trimToolText(m.content))}</div><div class="meta">${new Date(m.timestamp || now()).toLocaleTimeString()}</div></div>`;
    }
    if (m.role === 'assistant') {
      return `<div class="msg assistant"><div class="bubble">${assistantBubbleHtml(m)}</div><div class="meta">${new Date(m.timestamp || now()).toLocaleTimeString()}</div></div>`;
    }
    return `<div class="msg user"><div class="bubble">${escapeHtml(m.content)}</div><div class="meta">${new Date(m.timestamp || now()).toLocaleTimeString()}</div></div>`;
  }

  function patchStreamingMessage(assistantUi) {
    const msgEl = document.getElementById('messages');
    if (!msgEl) { render(); return; }
    const bubbles = msgEl.querySelectorAll('.msg.assistant .bubble');
    const bubble = bubbles[bubbles.length - 1];
    if (!bubble) { render(); return; }
    const atBottom = msgEl.scrollHeight - msgEl.scrollTop - msgEl.clientHeight < 40;
    bubble.innerHTML = assistantBubbleHtml(assistantUi);
    if (atBottom) msgEl.scrollTop = msgEl.scrollHeight;
  }

  function fillPrompt(text) {
    const input = document.getElementById('chat-input');
    if (!input) return;
    input.value = text;
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
  }

  function stopActiveRequest() {
    if (!state.isWorking) return;
    state.stopRequested = true;
    if (state.abortController && typeof state.abortController.abort === 'function') state.abortController.abort();
    if (state.pendingEval) { const p = state.pendingEval; state.pendingEval = null; p.resolve(false); }
    state.isWorking = false;
    render();
  }

  function confirmEvalCode(code) {
    return new Promise(resolve => {
      state.pendingEval = { code, resolve };
      render();
    });
  }
  function resolvePendingEval(allow) {
    const pending = state.pendingEval;
    if (!pending) return;
    state.pendingEval = null;
    render();
    pending.resolve(allow);
  }

  async function sendUserMessage(text) {
    if (state.isWorking) return;
    state.error = null;
    state.stopRequested = false;
    const myController = typeof AbortController !== 'undefined' ? new AbortController() : { signal: undefined };
    state.abortController = myController;
    ensureSession();
    state.messages.push({ role: 'user', content: text, timestamp: now() });
    state.isWorking = true;
    persistCurrentSession();
    render();
    try {
      await runAgentLoop();
      persistCurrentSession();
    } catch (e) {
      if (state.stopRequested || e?.name === 'AbortError') {
        markStoppedMessage();
        persistCurrentSession();
      } else {
        state.error = e.message || String(e);
      }
    } finally {
      // 仅当全局 controller 仍是本次请求时才复位，避免停止后立即重发时覆盖新请求的状态
      if (state.abortController === myController) {
        state.isWorking = false;
        state.abortController = null;
      }
      render();
    }
  }

  async function runManualTool() {
    const name = document.getElementById('manual-tool').value;
    const raw = document.getElementById('manual-args').value;
    state.toolOutput = 'running...'; render();
    try {
      const args = raw.trim() ? JSON.parse(raw) : {};
      const out = await executeToolByName(name, args);
      state.toolOutput = pretty(out);
      await maybeFollow(out);
    } catch (e) {
      state.toolOutput = 'ERROR: ' + (e.message || String(e));
    }
    render();
  }

  document.addEventListener('click', async (ev) => {
    const clickedInsideSessionMenu = ev.target.closest('.session-menu-wrap');
    const shouldCloseSessionMenu = state.sessionMenuOpen && !clickedInsideSessionMenu;
    const cite = ev.target.closest('a.citation');
    if (cite) {
      ev.preventDefault();
      const ref = cite.getAttribute('href').replace('#cite:', '');
      await navigateCitation(ref).catch(e => alert(e.message));
      return;
    }
    const tab = ev.target.closest('[data-tab]');
    if (tab) { if (shouldCloseSessionMenu) { state.sessionMenuOpen = false; state.pendingDeleteSessionId = null; } state.tab = tab.dataset.tab; render(); return; }
    const prompt = ev.target.closest('[data-prompt]');
    if (prompt) { fillPrompt(prompt.dataset.prompt || ''); return; }
    const actionEl = ev.target.closest('[data-action]');
    if (!actionEl) { if (shouldCloseSessionMenu) { state.sessionMenuOpen = false; state.pendingDeleteSessionId = null; render(); } return; }
    const action = actionEl.dataset.action;
    if (action === 'toggle-session-menu') {
      state.sessionMenuOpen = !state.sessionMenuOpen;
      state.pendingDeleteSessionId = null;
      render();
    } else if (action === 'send') {
      if (state.isWorking) { stopActiveRequest(); return; }
      const input = document.getElementById('chat-input');
      if (input && input.value.trim()) await sendUserMessage(input.value.trim());
    } else if (action === 'new-chat') {
      const s = makeSession(); state.sessions.unshift(s); state.currentSessionId = s.id; state.messages = []; state.sessionMenuOpen = false; state.pendingDeleteSessionId = null; saveSessions(); render();
    } else if (action === 'clear') {
      state.messages = []; persistCurrentSession(); render();
    } else if (action === 'toggle-theme') {
      state.theme = state.theme === 'dark' ? 'light' : 'dark'; setStoredItem(STORAGE_KEYS.theme, state.theme); render();
    } else if (action === 'toggle-locale') {
      state.locale = state.locale === 'zh' ? 'en' : 'zh'; setStoredItem(STORAGE_KEYS.locale, state.locale); render();
    } else if (action === 'toggle-follow') {
      state.settings.followMode = !state.settings.followMode; saveSettings(); render();
    } else if (action === 'switch-session') {
      const s = state.sessions.find(x => x.id === actionEl.dataset.id); if (s) { state.currentSessionId = s.id; state.messages = s.messages || []; state.sessionMenuOpen = false; state.pendingDeleteSessionId = null; render(); }
    } else if (action === 'delete-session') {
      state.pendingDeleteSessionId = actionEl.dataset.id;
      render();
    } else if (action === 'cancel-delete-session') {
      state.pendingDeleteSessionId = null;
      render();
    } else if (action === 'confirm-delete-session') {
      const deleteId = actionEl.dataset.id;
      state.sessions = state.sessions.filter(x => x.id !== deleteId);
      state.pendingDeleteSessionId = null;
      if (state.currentSessionId === deleteId) { state.currentSessionId = null; state.messages = []; }
      ensureSession(); saveSessions(); render();
    } else if (action === 'run-tool') {
      await runManualTool();
    } else if (action === 'eval-allow') {
      resolvePendingEval(true);
    } else if (action === 'eval-deny') {
      resolvePendingEval(false);
    }
  });

  document.addEventListener('change', (ev) => {
    const bind = ev.target.dataset.bind;
    if (bind) {
      if (ev.target.type === 'checkbox') state.settings[bind] = ev.target.checked;
      else state.settings[bind] = ev.target.value;
      if (bind === 'provider') {
        const provider = PROVIDERS[state.settings.provider] || PROVIDERS.Custom;
        state.settings.customPrefixUrl = provider.baseUrls[0].value || state.settings.customPrefixUrl;
        state.settings.model = provider.models[0].id;
      }
      saveSettings(); render(); return;
    }
    if (ev.target.dataset.action === 'pick-base-url') {
      if (ev.target.value !== '__custom__') state.settings.customPrefixUrl = ev.target.value;
      saveSettings(); render(); return;
    }
    if (ev.target.dataset.action === 'pick-model') {
      if (ev.target.value !== '__custom__') state.settings.model = ev.target.value;
      saveSettings(); render(); return;
    }
    if (ev.target.id === 'manual-tool') {
      const args = document.getElementById('manual-args');
      if (args) args.value = defaultArgsForTool(ev.target.value);
    }
  });

  document.addEventListener('input', (ev) => {
    const bind = ev.target.dataset.bind;
    if (bind) {
      state.settings[bind] = ev.target.value;
      saveSettings();
    }
  });
  document.addEventListener('keydown', async (ev) => {
    if (ev.target.id === 'chat-input' && ev.key === 'Enter' && !ev.shiftKey) {
      ev.preventDefault();
      if (ev.target.value.trim() && !state.isWorking) await sendUserMessage(ev.target.value.trim());
    }
  });

  App.render = render;
  App.patchStreamingMessage = patchStreamingMessage;
  App.confirmEvalCode = confirmEvalCode;
  App.sendUserMessage = sendUserMessage;
})();

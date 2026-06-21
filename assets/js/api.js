(function () {
  'use strict';
  const App = (window.App = window.App || {});

  function makeAbortError() {
    const e = new Error('Request aborted');
    e.name = 'AbortError';
    return e;
  }

  function markStoppedMessage() {
    const state = App.state;
    const text = App.t('stopped');
    const last = state.messages[state.messages.length - 1];
    if (last && last.role === 'assistant') {
      if (!last.content && !(last.toolCalls && last.toolCalls.length)) last.content = text;
      else if (last.content && !last.content.includes(text)) last.content += `\n\n_${text}_`;
    } else {
      state.messages.push({ role: 'assistant', content: text, timestamp: App.now() });
    }
  }

  async function runAgentLoop() {
    const state = App.state;
    if (!state.settings.apiKey) throw new Error('Please configure API key first');
    let metadataText = '';
    if (App.hasOffice()) {
      try { metadataText = `\n\n${App.host.metadataLabel || 'Document metadata'}:\n` + JSON.stringify(await App.host.getMetadata(), null, 2); }
      catch (e) { metadataText = '\n\nDocument metadata unavailable: ' + e.message; }
    }
    const conversation = toApiMessages(state.messages);
    const messages = [{ role: 'system', content: App.host.systemPrompt + metadataText }, ...conversation];
    const ms = Number(state.settings.maxAgentSteps);
    const maxSteps = Number.isFinite(ms) && ms > 0 ? ms : 8;

    for (let step = 0; step < maxSteps; step++) {
      if (state.stopRequested) throw makeAbortError();

      const assistantUi = { role: 'assistant', content: '', timestamp: App.now(), toolCalls: [] };
      state.messages.push(assistantUi);
      App.render();

      let lastRender = 0;
      const scheduleStreamRender = () => {
        const ts = Date.now();
        if (ts - lastRender > 80) { lastRender = ts; App.patchStreamingMessage(assistantUi); }
      };

      const res = await callChatCompletions(messages, {
        signal: state.abortController?.signal,
        onContent(delta) {
          assistantUi.content += delta;
          scheduleStreamRender();
        }
      });
      updateUsage(res.usage);
      const msg = res.choices?.[0]?.message || {};

      if (msg.content && !assistantUi.content) assistantUi.content = msg.content;

      if (msg.tool_calls && msg.tool_calls.length) {
        messages.push({ role: 'assistant', content: msg.content || '', tool_calls: msg.tool_calls });
        for (const tc of msg.tool_calls) {
          if (state.stopRequested) throw makeAbortError();
          const name = tc.function?.name;
          let args = {};
          try { args = JSON.parse(tc.function?.arguments || '{}'); } catch (e) { args = { _parseError: e.message, raw: tc.function?.arguments || '' }; }
          const uiCall = { id: tc.id, name, args, status: 'running', result: '' };
          assistantUi.toolCalls.push(uiCall); App.render();
          let toolResult;
          if (name === App.host.evalToolName) {
            const code = String(args?.code || '');
            const ok = await App.confirmEvalCode(code);
            if (!ok) {
              toolResult = { success: false, declined: true, error: 'User declined to run eval_officejs code.' };
              uiCall.status = 'error'; uiCall.result = toolResult;
              messages.push({ role: 'tool', tool_call_id: tc.id, name, content: JSON.stringify(toolResult) });
              App.render();
              continue;
            }
          }
          try {
            toolResult = await App.executeToolByName(name, args);
            if (state.stopRequested) throw makeAbortError();
            uiCall.status = toolResult && toolResult.success === false ? 'error' : 'complete';
            uiCall.result = toolResult;
            await App.maybeFollow(toolResult);
          } catch (e) {
            if (state.stopRequested || e?.name === 'AbortError') {
              uiCall.status = 'stopped';
              uiCall.result = { success: false, stopped: true };
              App.render();
              throw makeAbortError();
            }
            toolResult = { success: false, error: e.message || String(e) };
            uiCall.status = 'error'; uiCall.result = toolResult;
          }
          messages.push({ role: 'tool', tool_call_id: tc.id, name, content: JSON.stringify(toolResult) });
          App.render();
        }
        App.persistCurrentSession();
        continue;
      }

      if (!assistantUi.content && !(assistantUi.toolCalls && assistantUi.toolCalls.length)) {
        state.messages = state.messages.filter(m => m !== assistantUi);
      }
      App.render();
      return;
    }

    state.messages.push({ role: 'assistant', content: state.locale === 'zh' ? '工具调用步骤已达到上限，请继续发送指令让我接着完成。' : 'The tool-call step limit was reached. Send another instruction to continue.', timestamp: App.now() });
  }

  function toApiMessages(messages) {
    const out = [];
    for (const m of messages) {
      if (m.role === 'user') {
        if (String(m.content || '').trim()) out.push({ role: 'user', content: m.content });
      } else if (m.role === 'assistant') {
        const parts = [];
        if (String(m.content || '').trim()) parts.push(m.content);
        const toolSummary = summarizeToolCalls(m.toolCalls || []);
        if (toolSummary) parts.push(toolSummary);
        const content = parts.join('\n\n');
        if (content.trim()) out.push({ role: 'assistant', content });
      }
    }
    return out;
  }

  function summarizeToolCalls(toolCalls) {
    if (!Array.isArray(toolCalls) || !toolCalls.length) return '';
    const lines = toolCalls.map(tc => {
      const result = tc.result ? trimToolText(tc.result) : '';
      return `- ${tc.name || 'tool'} (${tc.status || 'unknown'}): ${result}`;
    });
    return `Tool execution summary from previous turn:\n${lines.join('\n')}`;
  }

  async function callChatCompletions(messages, options = {}) {
    const state = App.state;
    const tmp = Number(state.settings.temperature);
    const temperature = Number.isFinite(tmp) ? tmp : 0.2;
    const tools = App.host.toolDefinitions;
    const body = { model: state.settings.model, messages, temperature, stream: true };
    // 仅在有工具时才声明 tools/tool_choice；空数组配 tool_choice:'auto' 会被部分 OpenAI 兼容后端拒绝。
    if (Array.isArray(tools) && tools.length) { body.tools = tools; body.tool_choice = 'auto'; }
    if (state.settings.thinking && state.settings.thinking !== 'none') body.reasoning_effort = state.settings.thinking;
    const res = await fetch(App.chatEndpoint(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${state.settings.apiKey}` },
      body: JSON.stringify(body),
      signal: options.signal
    });
    if (!res.ok) {
      let text = await res.text().catch(() => '');
      try { const j = JSON.parse(text); text = j.error?.message || j.message || text; } catch {}
      throw new Error(`API ${res.status}: ${text || res.statusText}`);
    }

    const contentType = res.headers && res.headers.get ? (res.headers.get('content-type') || '') : '';
    if (!res.body || typeof res.body.getReader !== 'function' || contentType.includes('application/json')) {
      const json = await res.json();
      const content = json.choices?.[0]?.message?.content || '';
      if (content && options.onContent) options.onContent(content);
      return json;
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    const toolCalls = [];
    let content = '';
    let usage = null;
    let buffer = '';

    const processLine = (line) => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith(':')) return;
      if (!trimmed.startsWith('data:')) return;
      const data = trimmed.slice(5).trim();
      if (!data || data === '[DONE]') return;
      let chunk;
      try { chunk = JSON.parse(data); } catch { return; }
      if (chunk.usage) usage = chunk.usage;
      const choice = chunk.choices && chunk.choices[0];
      const delta = choice && choice.delta ? choice.delta : {};
      if (typeof delta.content === 'string') {
        content += delta.content;
        if (options.onContent) options.onContent(delta.content, content);
      }
      if (Array.isArray(delta.tool_calls)) mergeToolCallDeltas(toolCalls, delta.tool_calls);
    };

    while (true) {
      const { value, done } = await reader.read();
      buffer += decoder.decode(value || new Uint8Array(), { stream: !done });
      const lines = buffer.split(/\r?\n/);
      buffer = lines.pop() || '';
      for (const line of lines) processLine(line);
      if (done) break;
    }
    if (buffer.trim()) processLine(buffer);

    const message = { role: 'assistant', content };
    const normalizedToolCalls = toolCalls.filter(Boolean).map(tc => ({
      id: tc.id || `call_${App.id().replace(/[^a-zA-Z0-9_]/g, '')}`,
      type: tc.type || 'function',
      function: { name: tc.function.name || '', arguments: tc.function.arguments || '{}' }
    })).filter(tc => tc.function.name);
    if (normalizedToolCalls.length) message.tool_calls = normalizedToolCalls;
    return { choices: [{ message }], usage };
  }

  function mergeToolCallDeltas(toolCalls, deltas) {
    for (const part of deltas) {
      const index = Number.isInteger(part.index) ? part.index : 0;
      const target = toolCalls[index] || { id: '', type: 'function', function: { name: '', arguments: '' } };
      if (part.id) target.id = part.id;
      if (part.type) target.type = part.type;
      if (part.function) {
        if (part.function.name) target.function.name += part.function.name;
        if (part.function.arguments) target.function.arguments += part.function.arguments;
      }
      toolCalls[index] = target;
    }
  }

  function updateUsage(usage) {
    const state = App.state;
    state.stats.calls++;
    if (!usage) return;
    state.stats.inputTokens += usage.prompt_tokens || usage.input_tokens || 0;
    state.stats.outputTokens += usage.completion_tokens || usage.output_tokens || 0;
    state.stats.totalCost += 0;
  }

  function trimToolText(v) { const s = typeof v === 'string' ? v : App.pretty(v); return s.length > 4000 ? s.slice(0, 4000) + '\n... truncated ...' : s; }

  App.runAgentLoop = runAgentLoop;
  App.markStoppedMessage = markStoppedMessage;
  App.trimToolText = trimToolText;
})();

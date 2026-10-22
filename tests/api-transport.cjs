const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const apiSource = fs.readFileSync(path.resolve(__dirname, '../assets/js/api.js'), 'utf8');

async function testScenario(streaming) {
  const requests = [];
  const app = {
    state: {
      settings: { apiKey: 'test-key', model: 'test-model', temperature: 0.2 },
      messages: [], stopRequested: false,
      stats: { calls: 0, inputTokens: 0, outputTokens: 0, totalCost: 0 }
    },
    host: { systemPrompt: 'system', toolDefinitions: [], metadataLabel: 'Document' },
    hasOffice: () => false,
    chatEndpoint: () => 'https://example.invalid/v1/chat/completions',
    t: key => key,
    now: () => Date.now(),
    render: () => {},
    patchStreamingMessage: () => {},
    persistCurrentSession: () => {},
    id: () => 'test-id'
  };
  const responseMessage = { choices: [{ message: { content: 'Hello' } }], usage: {} };
  let index = 0;
  const body = new TextEncoder().encode(
    'data: {"choices":[{"delta":{"content":"Hello"}}]}\n\n' +
    'data: [DONE]\n\n'
  );
  const response = streaming
    ? {
        ok: true,
        headers: { get: () => 'text/event-stream' },
        body: {
          getReader: () => ({
            read: async () => index++ === 0
              ? { done: false, value: body }
              : { done: true, value: new Uint8Array() }
          })
        }
      }
    : {
        ok: true,
        headers: { get: () => 'application/json' },
        json: async () => responseMessage
      };

  function MockResponse() {}
  MockResponse.prototype.body = null;
  const context = {
    window: { App: app },
    fetch: async (url, init) => {
      requests.push({ url, body: JSON.parse(init.body) });
      return response;
    },
    ReadableStream: streaming ? function MockReadableStream() {} : undefined,
    TextDecoder: streaming ? TextDecoder : undefined,
    Response: streaming ? MockResponse : undefined,
    Uint8Array,
    console
  };
  vm.runInNewContext(apiSource, context, { filename: 'api.js' });
  await app.runAgentLoop();
  assert.strictEqual(requests.length, 1);
  assert.strictEqual(requests[0].body.stream, streaming);
  assert.strictEqual(requests[0].body.tools, undefined);
  assert.strictEqual(app.state.messages[0].content, 'Hello');
  assert.strictEqual(app.state.stats.calls, 1);
}

(async () => {
  await testScenario(false);
  await testScenario(true);
  console.log('PASS: JSON fallback and normal SSE streaming');
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});

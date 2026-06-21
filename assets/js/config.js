(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const STORAGE_KEYS = {
    settings: 'officedpoqb-provider-config',
    locale: 'dpoqboffice-locale',
    theme: 'dpoqboffice-theme',
    sessions: 'OpenExcelDB_v3.sessions.plain',
    workbookId: 'dpoqboffice-workbook-id', // 通用 documentId（沿用旧键名，保持向后兼容）
    sheetMap: 'dpoqbexcel-sheet-map-v1'   // 仅 Excel 使用
  };

  const DEFAULT_BASE_URL = 'https://api.dpoqb.top/v1';
  const DEFAULT_SETTINGS = {
    provider: 'dpoqb',
    model: 'gpt-oss-120b',
    apiKey: '',
    customPrefixUrl: DEFAULT_BASE_URL,
    thinking: 'none',
    followMode: true,
    temperature: 0.2
  };

  const PROVIDERS = {
    dpoqb: {
      label: 'dpoqb LLM',
      baseUrls: [
        { label: 'api通用接口', value: DEFAULT_BASE_URL },
        { label: '测试接口', value: 'https://api.dpoqb.top/t1' }
      ],
      models: [
        { id: 'gpt-oss-120b', name: 'GPT OSS 120B', contextWindow: 200000, maxTokens: 128000 },
        { id: 'gpt-oss-20b', name: 'GPT OSS 20B', contextWindow: 128000, maxTokens: 64000 }
      ]
    },
    OpenAI: {
      label: 'OpenAI Compatible',
      baseUrls: [{ label: 'OpenAI', value: 'https://api.openai.com/v1' }],
      models: [
        { id: 'gpt-5.2-chat-latest', name: 'GPT-5.2 Chat' },
        { id: 'gpt-5.1-chat-latest', name: 'GPT-5.1 Chat' },
        { id: 'gpt-4.1', name: 'GPT-4.1' },
        { id: 'gpt-4o', name: 'GPT-4o' }
      ]
    },
    DeepSeek: {
      label: 'DeepSeek Compatible',
      baseUrls: [{ label: 'DeepSeek', value: 'https://api.deepseek.com/v1' }],
      models: [
        { id: 'deepseek-chat', name: 'DeepSeek Chat' },
        { id: 'deepseek-reasoner', name: 'DeepSeek Reasoner' }
      ]
    },
    OpenRouter: {
      label: 'OpenRouter',
      baseUrls: [{ label: 'OpenRouter', value: 'https://openrouter.ai/api/v1' }],
      models: [
        { id: 'openai/gpt-5.2-chat', name: 'OpenAI GPT-5.2 Chat' },
        { id: 'deepseek/deepseek-v3.2', name: 'DeepSeek V3.2' },
        { id: 'moonshotai/kimi-k2-thinking', name: 'Kimi K2 Thinking' }
      ]
    },
    Groq: {
      label: 'Groq',
      baseUrls: [{ label: 'Groq', value: 'https://api.groq.com/openai/v1' }],
      models: [
        { id: 'openai/gpt-oss-120b', name: 'GPT OSS 120B' },
        { id: 'deepseek-r1-distill-llama-70b', name: 'DeepSeek R1 Distill Llama 70B' }
      ]
    },
    Custom: {
      label: 'Custom OpenAI Compatible',
      baseUrls: [{ label: '自定义', value: '' }],
      models: [{ id: 'custom-model', name: 'custom-model' }]
    }
  };

  const I18N = {
    zh: {
      chat: '对话', settings: '设置', tools: '工具', newChat: '新建对话', clear: '清空消息', deleteSession: '删除当前会话',
      followOn: '跟随模式：已开启', followOff: '跟随模式：已关闭', light: '浅色', dark: '深色',
      title: '准备好处理你的文档', subtitle: '你可以让我分析、生成或修改内容',
      input: '告诉我你想如何处理…', noConfig: '请先在设置中配置 API 密钥', send: '发送', stop: '停止', stopped: '已停止生成。',
      chart: '智能生成', chartDesc: '根据需求自动生成内容',
      fix: '检查与修正', fixDesc: '定位问题并给出修复结果',
      analyze: '结构智能解析', analyzeDesc: '提炼结构并输出结论',
      chartPrompt: '请根据当前文档内容，帮我生成合适的内容',
      fixPrompt: '帮我检查当前文档的问题，并给出修复后的结果',
      analyzePrompt: '帮我全面分析当前文档内容，并输出汇总结论',
      provider: '服务商', env: '接口环境', model: '模型', apiKey: 'API 密钥', baseUrl: 'Base URL', thinking: '思考模式',
      configured: 'API 已配置', notConfigured: '填写以上信息即可开始', about: '关于',
      aboutText: 'dpoqb Office 助手使用 AI 为你的 Office 文档提供智能对话能力。API 密钥仅保存在浏览器本地。',
      sessions: '会话', usage: '用量统计', advanced: '高级工具', manualTool: '手动执行工具', run: '执行',
      toolName: '工具名称', params: '参数 JSON', output: '输出', demo: '当前不在 Office 环境中，文档工具只能在插件侧边栏里运行。',
      lang: 'EN', configuredShort: '已配置', notConfiguredShort: '未配置',
      sessionHistory: '会话历史', noSessions: '暂无历史会话', confirmDeleteSession: '确定删除该会话吗？删除后不可恢复。', currentSession: '当前会话', confirm: '确认', cancel: '取消',
      confirmEval: 'AI 请求执行以下 Office.js 代码，是否允许？', evalAllow: '允许执行', evalDeny: '拒绝'
    },
    en: {
      chat: 'Chat', settings: 'Settings', tools: 'Tools', newChat: 'New Chat', clear: 'Clear messages', deleteSession: 'Delete current session',
      followOn: 'Follow mode: ON', followOff: 'Follow mode: OFF', light: 'Light', dark: 'Dark',
      title: 'Ready to work with your document', subtitle: 'Ask me to analyze, generate, or modify content',
      input: 'Tell me what to do…', noConfig: 'Configure API key in settings to get started', send: 'Send', stop: 'Stop', stopped: 'Stopped.',
      chart: 'Smart Generation', chartDesc: 'Generate content on demand',
      fix: 'Check & Fix', fixDesc: 'Detect issues and fix them',
      analyze: 'Structure Analysis', analyzeDesc: 'Extract structure and conclusions',
      chartPrompt: 'Generate suitable content based on the current document',
      fixPrompt: 'Check the current document for issues and fix them',
      analyzePrompt: 'Analyze the current document and summarize key conclusions',
      provider: 'Provider', env: 'Environment', model: 'Model', apiKey: 'API Key', baseUrl: 'Base URL', thinking: 'Thinking Mode',
      configured: 'API configured', notConfigured: 'Fill in all fields above to get started', about: 'About',
      aboutText: 'dpoqb Office Assistant gives your Office documents AI chat and action capabilities. API keys are only stored locally in your browser.',
      sessions: 'Sessions', usage: 'Usage', advanced: 'Advanced tools', manualTool: 'Manual tool runner', run: 'Run',
      toolName: 'Tool name', params: 'Params JSON', output: 'Output', demo: 'Not currently running inside Office. Document tools only work in the add-in task pane.',
      lang: '中文', configuredShort: 'Configured', notConfiguredShort: 'Not configured',
      sessionHistory: 'Session History', noSessions: 'No sessions yet', confirmDeleteSession: 'Delete this session? This cannot be undone.', currentSession: 'Current session', confirm: 'Confirm', cancel: 'Cancel',
      confirmEval: 'The AI requests to run the following Office.js code. Allow it?', evalAllow: 'Allow', evalDeny: 'Deny'
    }
  };

  App.STORAGE_KEYS = STORAGE_KEYS;
  App.DEFAULT_BASE_URL = DEFAULT_BASE_URL;
  App.DEFAULT_SETTINGS = DEFAULT_SETTINGS;
  App.PROVIDERS = PROVIDERS;
  App.I18N = I18N;
})();

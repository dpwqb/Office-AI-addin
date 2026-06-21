(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const STORAGE_KEYS = {
    settings: 'exceldpoqb-provider-config',
    locale: 'dpoqbexcel-locale',
    theme: 'dpoqbexcel-theme',
    sessions: 'OpenExcelDB_v3.sessions.plain',
    workbookId: 'dpoqbexcel-workbook-id',
    sheetMap: 'dpoqbexcel-sheet-map-v1'
  };

  const DEFAULT_BASE_URL = 'https://api.dpoqb.top/v1';
  const DEFAULT_SETTINGS = {
    provider: 'dpoqb',
    model: 'gpt-oss-120b',
    apiKey: '',
    customPrefixUrl: DEFAULT_BASE_URL,
    thinking: 'none',
    followMode: true,
    temperature: 0.2,
    maxAgentSteps: 8
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
      title: '准备好处理你的 Excel 数据', subtitle: '你可以让我分析、可视化或转换你的数据',
      input: '告诉我你想如何处理这份表格…', noConfig: '请先在设置中配置 API 密钥', send: '发送', stop: '停止', stopped: '已停止生成。',
      chart: '智能图表生成', chartDesc: '自动推荐图表类型并一键生成',
      fix: '公式错误诊断', fixDesc: '定位报错原因并生成修复公式',
      analyze: '跨表智能解析', analyzeDesc: '自动关联多表数据并输出结论',
      provider: '服务商', env: '接口环境', model: '模型', apiKey: 'API 密钥', baseUrl: 'Base URL', thinking: '思考模式',
      configured: 'API 已配置', notConfigured: '填写以上信息即可开始', about: '关于',
      aboutText: 'dpoqb Excel 助手使用 AI 为你的 Excel 提供智能对话能力。API 密钥仅保存在浏览器本地。',
      sessions: '会话', usage: '用量统计', advanced: '高级工具', manualTool: '手动执行工具', run: '执行',
      toolName: '工具名称', params: '参数 JSON', output: '输出', demo: '当前不在 Excel/Office 环境中，Excel 工具只能在插件侧边栏里运行。',
      lang: 'EN', configuredShort: '已配置', notConfiguredShort: '未配置',
      sessionHistory: '会话历史', noSessions: '暂无历史会话', confirmDeleteSession: '确定删除该会话吗？删除后不可恢复。', currentSession: '当前会话', confirm: '确认', cancel: '取消',
      confirmEval: 'AI 请求执行以下 Office.js 代码，是否允许？', evalAllow: '允许执行', evalDeny: '拒绝'
    },
    en: {
      chat: 'Chat', settings: 'Settings', tools: 'Tools', newChat: 'New Chat', clear: 'Clear messages', deleteSession: 'Delete current session',
      followOn: 'Follow mode: ON', followOff: 'Follow mode: OFF', light: 'Light', dark: 'Dark',
      title: 'Ready to work with your Excel data', subtitle: 'Ask me to analyze, visualize, or transform your data',
      input: 'Tell me what to do with this workbook…', noConfig: 'Configure API key in settings to get started', send: 'Send', stop: 'Stop', stopped: 'Stopped.',
      chart: 'Chart Generation', chartDesc: 'One-click visualization & styling',
      fix: 'Error Fix', fixDesc: 'Auto-detect & fix formula errors',
      analyze: 'Multi-Sheet Analysis', analyzeDesc: 'Cross-sheet automation and conclusions',
      provider: 'Provider', env: 'Environment', model: 'Model', apiKey: 'API Key', baseUrl: 'Base URL', thinking: 'Thinking Mode',
      configured: 'API configured', notConfigured: 'Fill in all fields above to get started', about: 'About',
      aboutText: 'dpoqb Excel Assistant gives your Excel workbook AI chat and action capabilities. API keys are only stored locally in your browser.',
      sessions: 'Sessions', usage: 'Usage', advanced: 'Advanced tools', manualTool: 'Manual tool runner', run: 'Run',
      toolName: 'Tool name', params: 'Params JSON', output: 'Output', demo: 'Not currently running inside Excel/Office. Excel tools only work in the add-in task pane.',
      lang: '中文', configuredShort: 'Configured', notConfiguredShort: 'Not configured',
      sessionHistory: 'Session History', noSessions: 'No sessions yet', confirmDeleteSession: 'Delete this session? This cannot be undone.', currentSession: 'Current session', confirm: 'Confirm', cancel: 'Cancel',
      confirmEval: 'The AI requests to run the following Office.js code. Allow it?', evalAllow: 'Allow', evalDeny: 'Deny'
    }
  };

  const SYSTEM_PROMPT = `You are an AI assistant integrated into Microsoft Excel with full access to read and modify spreadsheet data.

Available tools:
READ:
- get_cell_ranges: Read cell values, formulas, and formatting
- get_range_as_csv: Get data as CSV, useful for analysis
- search_data: Find text across the spreadsheet
- get_all_objects: List charts, pivot tables, and other objects

WRITE:
- set_cell_range: Write values, formulas, notes, and formatting
- clear_cell_range: Clear contents or formatting
- copy_to: Copy ranges with formula translation
- modify_sheet_structure: Insert/delete/hide/unhide rows/columns, freeze panes
- modify_workbook_structure: Create/delete/rename/duplicate sheets
- resize_range: Adjust column widths and row heights
- modify_object: Create/update/delete charts and pivot tables
- eval_officejs: Execute Office.js code when the listed tools are not enough

Citations: Use markdown links with #cite: hash to reference sheets/cells. Clicking navigates there.
- Sheet only: [Sheet Name](#cite:sheetId)
- Cell/range: [A1:B10](#cite:sheetId!A1:B10)
Example: [Exchange Ratio](#cite:3) or [see cell B5](#cite:3!B5)

When the user asks about their workbook data, read it first. Be concise. Use A1 notation for cell references. Before overwriting existing data, confirm unless the user explicitly asks to replace or overwrite.`;

  const TOOL_DEFINITIONS = [
    { type: 'function', function: { name: 'get_cell_ranges', description: 'Read cell values, formulas, and formatting from specified ranges in a worksheet. Returns cells as a sparse object with A1-notation keys.', parameters: { type: 'object', properties: { sheetId: { type: 'number', description: 'The worksheet ID (1-based index)' }, ranges: { type: 'array', items: { type: 'string' }, description: "Array of ranges in A1 notation, e.g. ['A1:C10']" }, includeStyles: { type: 'boolean', description: 'Include font/fill styling info. Default true' }, cellLimit: { type: 'number', description: 'Maximum cells to return. Default 2000' }, explanation: { type: 'string' } }, required: ['sheetId', 'ranges'] } } },
    { type: 'function', function: { name: 'get_range_as_csv', description: 'Read cell data from a range and return it as CSV format. Great for analysis.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, includeHeaders: { type: 'boolean' }, maxRows: { type: 'number' }, explanation: { type: 'string' } }, required: ['sheetId', 'range'] } } },
    { type: 'function', function: { name: 'search_data', description: 'Find text or values across the spreadsheet. Supports regex and case-sensitive search.', parameters: { type: 'object', properties: { searchTerm: { type: 'string' }, sheetId: { type: 'number' }, range: { type: 'string' }, offset: { type: 'number' }, options: { type: 'object', properties: { matchCase: { type: 'boolean' }, matchEntireCell: { type: 'boolean' }, matchFormulas: { type: 'boolean' }, useRegex: { type: 'boolean' }, maxResults: { type: 'number' } } }, explanation: { type: 'string' } }, required: ['searchTerm'] } } },
    { type: 'function', function: { name: 'get_all_objects', description: 'List all charts, pivot tables, and other objects in the workbook.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, id: { type: 'string' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'set_cell_range', description: "WRITE. Write values, formulas, notes, and formatting to cells. By default fails if target cells contain data. Retry with allow_overwrite=true after confirmation.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, cells: { type: 'array', items: { type: 'array', items: { type: 'object', properties: { value: {}, formula: { type: 'string' }, note: { type: 'string' }, cellStyles: { type: 'object' }, borderStyles: { type: 'object' } } } } }, copyToRange: { type: 'string' }, resizeWidth: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, resizeHeight: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, allow_overwrite: { type: 'boolean' }, explanation: { type: 'string' } }, required: ['sheetId', 'range', 'cells'] } } },
    { type: 'function', function: { name: 'clear_cell_range', description: "Clear contents, formatting, or both from a range. clearType: contents/formats/all.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, clearType: { enum: ['contents', 'formats', 'all'] }, explanation: { type: 'string' } }, required: ['sheetId', 'range'] } } },
    { type: 'function', function: { name: 'copy_to', description: 'Copy a range to another location with formula translation. If destination is larger, source pattern repeats.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, sourceRange: { type: 'string' }, destinationRange: { type: 'string' }, explanation: { type: 'string' } }, required: ['sheetId', 'sourceRange', 'destinationRange'] } } },
    { type: 'function', function: { name: 'modify_sheet_structure', description: "Insert, delete, hide, unhide, or freeze rows and columns. Use reference like '5' or 'C'.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, operation: { enum: ['insert', 'delete', 'hide', 'unhide', 'freeze', 'unfreeze'] }, dimension: { enum: ['rows', 'columns'] }, reference: { type: 'string' }, count: { type: 'number' }, position: { enum: ['before', 'after'] }, explanation: { type: 'string' } }, required: ['sheetId', 'operation', 'dimension'] } } },
    { type: 'function', function: { name: 'modify_workbook_structure', description: 'Create, delete, rename, or duplicate worksheets.', parameters: { type: 'object', properties: { operation: { enum: ['create', 'delete', 'rename', 'duplicate'] }, sheetId: { type: 'number' }, sheetName: { type: 'string' }, newName: { type: 'string' }, tabColor: { type: 'string' }, explanation: { type: 'string' } }, required: ['operation'] } } },
    { type: 'function', function: { name: 'resize_range', description: "Adjust column widths or row heights. Use 'A:D' for columns, '1:5' for rows, or omit range for entire sheet.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, width: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, height: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, explanation: { type: 'string' } }, required: ['sheetId'] } } },
    { type: 'function', function: { name: 'modify_object', description: 'Create, update, or delete charts and pivot tables.', parameters: { type: 'object', properties: { operation: { enum: ['create', 'update', 'delete'] }, sheetId: { type: 'number' }, objectType: { enum: ['pivotTable', 'chart'] }, id: { type: 'string' }, properties: { type: 'object', properties: { name: { type: 'string' }, source: { type: 'string' }, range: { type: 'string' }, anchor: { type: 'string' }, rows: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' } } } }, columns: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' } } } }, values: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' }, summarizeBy: { enum: ['sum', 'count', 'average', 'max', 'min'] } } } }, title: { type: 'string' }, chartType: { enum: ['columnClustered', 'barClustered', 'line', 'pie', 'scatter', 'area', 'doughnut'] } } }, explanation: { type: 'string' } }, required: ['operation', 'sheetId', 'objectType'] } } },
    { type: 'function', function: { name: 'eval_officejs', description: 'Execute arbitrary Office.js code in Excel.run. Escape hatch for unsupported operations. Code receives context and Excel.', parameters: { type: 'object', properties: { code: { type: 'string' }, explanation: { type: 'string' } }, required: ['code'] } } }
  ];

  function defaultArgsForTool(name) {
    const samples = {
      get_cell_ranges: { sheetId: 1, ranges: ['A1:D10'], includeStyles: true, cellLimit: 2000 },
      get_range_as_csv: { sheetId: 1, range: 'A1:D10', includeHeaders: true, maxRows: 500 },
      search_data: { searchTerm: 'keyword', options: { matchCase: false, useRegex: false, maxResults: 100 } },
      get_all_objects: {},
      set_cell_range: { sheetId: 1, range: 'A1:B2', cells: [[{ value: '标题1', cellStyles: { fontWeight: 'bold' } }, { value: '标题2', cellStyles: { fontWeight: 'bold' } }], [{ value: 1 }, { formula: '=A2*2' }]], allow_overwrite: false },
      clear_cell_range: { sheetId: 1, range: 'A1:B2', clearType: 'contents' },
      copy_to: { sheetId: 1, sourceRange: 'A1:B2', destinationRange: 'D1:E2' },
      modify_sheet_structure: { sheetId: 1, operation: 'insert', dimension: 'rows', reference: '5', count: 1, position: 'before' },
      modify_workbook_structure: { operation: 'create', sheetName: 'AI分析结果', tabColor: '#134cff' },
      resize_range: { sheetId: 1, range: 'A:D', width: { type: 'points', value: 90 } },
      modify_object: { operation: 'create', sheetId: 1, objectType: 'chart', properties: { source: 'A1:B10', chartType: 'columnClustered', anchor: 'E2', title: 'Chart' } },
      eval_officejs: { code: "const range = context.workbook.worksheets.getActiveWorksheet().getRange('A1');\nrange.load('values');\nawait context.sync();\nreturn range.values;" }
    };
    return App.pretty(samples[name] || {});
  }

  App.STORAGE_KEYS = STORAGE_KEYS;
  App.DEFAULT_BASE_URL = DEFAULT_BASE_URL;
  App.DEFAULT_SETTINGS = DEFAULT_SETTINGS;
  App.PROVIDERS = PROVIDERS;
  App.I18N = I18N;
  App.SYSTEM_PROMPT = SYSTEM_PROMPT;
  App.TOOL_DEFINITIONS = TOOL_DEFINITIONS;
  App.defaultArgsForTool = defaultArgsForTool;
})();

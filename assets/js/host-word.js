(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const requireOffice = () => App.requireOffice();

  // ---- 段落锚点 id：用 Word 内容控件 tag 作为稳定引用，回退到段落序号 ----
  function clampText(s, n) { s = String(s == null ? '' : s); return s.length > n ? s.slice(0, n) + '…' : s; }

  async function getDocumentOutline(args = {}) {
    requireOffice();
    const maxParagraphs = Math.max(1, Number(args.maxParagraphs || 400));
    return Word.run(async context => {
      const body = context.document.body;
      const paras = body.paragraphs;
      paras.load('items');
      await context.sync();
      const items = paras.items.slice(0, maxParagraphs);
      items.forEach(p => p.load('text,styleBuiltIn,style,isListItem,outlineLevel'));
      await context.sync();
      const outline = [];
      const paragraphs = [];
      items.forEach((p, idx) => {
        const text = (p.text || '').trim();
        const style = p.styleBuiltIn || p.style || '';
        const isHeading = /heading/i.test(String(style));
        const level = isHeading ? (parseInt(String(style).replace(/\D/g, ''), 10) || 1) : 0;
        if (text || isHeading) paragraphs.push({ index: idx, text: clampText(text, 200), style: String(style) });
        if (isHeading && text) outline.push({ index: idx, level, text: clampText(text, 200) });
      });
      return { success: true, paragraphCount: paras.items.length, returned: items.length, hasMore: paras.items.length > items.length, outline, paragraphs };
    });
  }

  async function getSelection() {
    requireOffice();
    return Word.run(async context => {
      const sel = context.document.getSelection();
      sel.load('text,styleBuiltIn,style,font/bold,font/italic,font/size,font/name,font/color');
      await context.sync();
      return { success: true, text: sel.text || '', style: String(sel.styleBuiltIn || sel.style || ''), font: { bold: sel.font.bold, italic: sel.font.italic, size: sel.font.size, name: sel.font.name, color: sel.font.color } };
    });
  }

  async function getParagraphs(args = {}) {
    requireOffice();
    const start = Math.max(0, Number(args.start || 0));
    const count = Math.max(1, Number(args.count || 20));
    return Word.run(async context => {
      const paras = context.document.body.paragraphs;
      paras.load('items');
      await context.sync();
      const slice = paras.items.slice(start, start + count);
      slice.forEach(p => p.load('text,styleBuiltIn'));
      await context.sync();
      return { success: true, total: paras.items.length, paragraphs: slice.map((p, i) => ({ index: start + i, text: p.text || '', style: String(p.styleBuiltIn || '') })) };
    });
  }

  async function getTables(args = {}) {
    requireOffice();
    const maxTables = Math.max(1, Number(args.maxTables || 20));
    return Word.run(async context => {
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      const slice = tables.items.slice(0, maxTables);
      slice.forEach(t => t.load('values,rowCount'));
      await context.sync();
      return { success: true, tableCount: tables.items.length, tables: slice.map((t, i) => ({ index: i, rowCount: t.rowCount, values: t.values })) };
    });
  }

  async function searchText(args) {
    requireOffice();
    const { query, matchCase = false, matchWholeWord = false, useWildcards = false, maxResults = 100 } = args;
    if (!query) throw new Error('query is required');
    return Word.run(async context => {
      const results = context.document.body.search(query, { matchCase, matchWholeWord, matchWildcards: useWildcards });
      results.load('items');
      await context.sync();
      const slice = results.items.slice(0, maxResults);
      slice.forEach(r => r.load('text'));
      await context.sync();
      return { success: true, totalFound: results.items.length, returned: slice.length, matches: slice.map((r, i) => ({ index: i, text: clampText(r.text, 160) })) };
    });
  }

  // 仅在跟随模式开启时选中目标 range（select 会滚动视图）。统一受 followMode 控制，
  // 避免在工具内部硬编码 select 而绕过开关。在 range 仍有效的 Word.run 内调用。
  function followSelect(range) {
    if (App.state.settings.followMode) { try { range.select(); } catch {} }
  }

  const WHERE_MAP = { Replace: 'Replace', Start: 'Start', End: 'End', Before: 'Before', After: 'After' };
  async function insertText(args) {
    requireOffice();
    const { text = '', location = 'End', style } = args;
    const where = WHERE_MAP[location] || 'End';
    return Word.run(async context => {
      const body = context.document.body;
      let range;
      if (where === 'Replace' || where === 'Before' || where === 'After') {
        const sel = context.document.getSelection();
        range = sel.insertText(text, where === 'Replace' ? Word.InsertLocation.replace : (where === 'Before' ? Word.InsertLocation.before : Word.InsertLocation.after));
      } else {
        range = body.insertText(text, where === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end);
      }
      if (style) range.styleBuiltIn = style;
      followSelect(range);
      range.load('text');
      await context.sync();
      return { success: true, inserted: clampText(text, 120), location: where, _navTarget: { kind: 'selection' } };
    });
  }

  async function replaceText(args) {
    requireOffice();
    const { query, replacement = '', matchCase = false, matchWholeWord = false, useWildcards = false, replaceAll = true } = args;
    if (!query) throw new Error('query is required');
    return Word.run(async context => {
      const results = context.document.body.search(query, { matchCase, matchWholeWord, matchWildcards: useWildcards });
      results.load('items');
      await context.sync();
      const targets = replaceAll ? results.items : results.items.slice(0, 1);
      targets.forEach(r => r.insertText(replacement, Word.InsertLocation.replace));
      await context.sync();
      return { success: true, replaced: targets.length, totalFound: results.items.length, _navTarget: { kind: 'selection' } };
    });
  }

  async function applyStyle(args) {
    requireOffice();
    const { target = 'selection', style, font = {} } = args;
    return Word.run(async context => {
      const range = target === 'selection' ? context.document.getSelection() : context.document.body.getRange();
      if (style) range.styleBuiltIn = style;
      if (font.bold != null) range.font.bold = !!font.bold;
      if (font.italic != null) range.font.italic = !!font.italic;
      if (font.underline != null) range.font.underline = font.underline ? 'Single' : 'None';
      if (font.size) range.font.size = font.size;
      if (font.name) range.font.name = font.name;
      if (font.color) range.font.color = font.color;
      await context.sync();
      return { success: true, target, style: style || null, _navTarget: { kind: 'selection' } };
    });
  }

  async function setParagraphFormat(args) {
    requireOffice();
    const { target = 'selection', alignment, lineSpacing, leftIndent, spaceBefore, spaceAfter } = args;
    return Word.run(async context => {
      const range = target === 'selection' ? context.document.getSelection() : context.document.body.getRange();
      const pf = range.paragraphFormat;
      if (alignment) pf.alignment = alignment; // Left/Centered/Right/Justified
      if (lineSpacing) pf.lineSpacing = lineSpacing;
      if (leftIndent != null) pf.leftIndent = leftIndent;
      if (spaceBefore != null) pf.spaceBefore = spaceBefore;
      if (spaceAfter != null) pf.spaceAfter = spaceAfter;
      await context.sync();
      return { success: true, target, _navTarget: { kind: 'selection' } };
    });
  }

  async function insertHeading(args) {
    requireOffice();
    const { text = '', level = 1, location = 'End' } = args;
    return Word.run(async context => {
      const body = context.document.body;
      const range = body.insertParagraph(text, location === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end);
      range.styleBuiltIn = `Heading${Math.min(Math.max(Number(level) || 1, 1), 6)}`;
      followSelect(range);
      await context.sync();
      return { success: true, text: clampText(text, 120), level, _navTarget: { kind: 'selection' } };
    });
  }

  async function insertTable(args) {
    requireOffice();
    const { rows, location = 'End', headerRow = true } = args;
    if (!Array.isArray(rows) || !rows.length || !rows.every(Array.isArray)) throw new Error('rows must be a non-empty 2D array');
    const rowCount = rows.length;
    const colCount = rows[0].length;
    return Word.run(async context => {
      const body = context.document.body;
      const values = rows.map(r => r.map(c => (c == null ? '' : String(c))));
      const table = body.insertTable(rowCount, colCount, location === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end, values);
      if (headerRow && rowCount > 0) {
        // 首行作为表头：加粗。不用 styleFirstRow，因其依赖表样式是否定义了表头格式，跨版本不可靠。
        try { table.getCell(0, 0).parentRow.font.bold = true; } catch {}
      }
      followSelect(table);
      await context.sync();
      return { success: true, rowCount, columnCount: colCount, _navTarget: { kind: 'selection' } };
    });
  }

  async function insertPageBreak(args = {}) {
    requireOffice();
    const { location = 'End' } = args;
    return Word.run(async context => {
      const body = context.document.body;
      const range = location === 'selection' ? context.document.getSelection() : body.getRange(location === 'Start' ? 'Start' : 'End');
      range.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
      followSelect(range);
      await context.sync();
      return { success: true, _navTarget: { kind: 'selection' } };
    });
  }

  async function insertImage(args) {
    requireOffice();
    const { base64, location = 'End' } = args;
    if (!base64) throw new Error('base64 image data is required');
    return Word.run(async context => {
      const body = context.document.body;
      const range = location === 'selection' ? context.document.getSelection() : body;
      const pic = range.insertInlinePictureFromBase64(base64, location === 'Start' ? Word.InsertLocation.start : Word.InsertLocation.end);
      followSelect(pic);
      await context.sync();
      return { success: true, _navTarget: { kind: 'selection' } };
    });
  }

  async function manageComment(args) {
    requireOffice();
    const { operation = 'add', text = '' } = args;
    return Word.run(async context => {
      if (operation === 'add') {
        const sel = context.document.getSelection();
        if (typeof sel.insertComment !== 'function') throw new Error('Comment API not available in this Word host');
        sel.insertComment(text);
        await context.sync();
        return { success: true, operation, _navTarget: { kind: 'selection' } };
      }
      throw new Error(`Unsupported comment operation: ${operation}`);
    });
  }

  // 内容控件作为可导航锚点：add 时返回 tag，navigateCitation 时按 tag/title 选中
  async function manageContentControl(args) {
    requireOffice();
    const { operation = 'add', tag, title, text } = args;
    return Word.run(async context => {
      if (operation === 'add') {
        const sel = context.document.getSelection();
        const cc = sel.insertContentControl();
        if (tag) cc.tag = tag;
        if (title) cc.title = title;
        if (text) cc.insertText(text, Word.InsertLocation.replace);
        cc.load('id,tag,title');
        await context.sync();
        return { success: true, operation, id: cc.id, tag: cc.tag, title: cc.title, _navTarget: { kind: 'contentControl', tag: cc.tag || tag } };
      }
      if (operation === 'list') {
        const ccs = context.document.contentControls;
        ccs.load('items');
        await context.sync();
        ccs.items.forEach(c => c.load('id,tag,title,text'));
        await context.sync();
        return { success: true, controls: ccs.items.map(c => ({ id: c.id, tag: c.tag, title: c.title, text: clampText(c.text, 120) })) };
      }
      throw new Error(`Unsupported content control operation: ${operation}`);
    });
  }

  async function evalOfficeJs(args) {
    requireOffice();
    const code = args.code || '';
    return Word.run(async context => {
      const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
      const fn = new AsyncFunction('context', 'Word', code);
      const result = await fn(context, Word);
      return { success: true, result: result ?? null };
    });
  }

  async function selectByParagraphIndex(index) {
    return Word.run(async context => {
      const paras = context.document.body.paragraphs;
      paras.load('items');
      await context.sync();
      const p = paras.items[Number(index)];
      if (!p) throw new Error(`Paragraph ${index} not found`);
      p.select();
      await context.sync();
      return { success: true };
    });
  }
  async function selectByContentControlTag(tag) {
    return Word.run(async context => {
      const ccs = context.document.contentControls;
      ccs.load('items');
      await context.sync();
      ccs.items.forEach(c => c.load('tag'));
      await context.sync();
      const cc = ccs.items.find(c => c.tag === tag);
      if (!cc) throw new Error(`Content control '${tag}' not found`);
      cc.select();
      await context.sync();
      return { success: true };
    });
  }

  async function navigateCitation(ref) {
    requireOffice();
    // 格式: "p:<index>" 段落序号 | "cc:<tag>" 内容控件 | 纯数字回退为段落序号
    if (/^cc:/.test(ref)) return selectByContentControlTag(ref.slice(3));
    if (/^p:/.test(ref)) return selectByParagraphIndex(ref.slice(2));
    return selectByParagraphIndex(ref);
  }

  async function maybeFollow(result) {
    if (!App.state.settings.followMode || !result) return;
    const nav = result._navTarget;
    if (!nav) return;
    try {
      if (nav.kind === 'contentControl' && nav.tag) await selectByContentControlTag(nav.tag);
      // kind === 'selection'：插入类工具已对新内容调用 select()，选区即落在新内容上，无需再次定位。
    } catch (e) { console.warn(e); }
  }

  const TOOL_EXECUTORS = {
    get_document_outline: getDocumentOutline,
    get_selection: getSelection,
    get_paragraphs: getParagraphs,
    get_tables: getTables,
    search_text: searchText,
    insert_text: insertText,
    replace_text: replaceText,
    apply_style: applyStyle,
    set_paragraph_format: setParagraphFormat,
    insert_heading: insertHeading,
    insert_table: insertTable,
    insert_page_break: insertPageBreak,
    insert_image: insertImage,
    manage_comment: manageComment,
    manage_content_control: manageContentControl,
    eval_officejs: evalOfficeJs
  };

  async function getDocumentMetadata() {
    requireOffice();
    return Word.run(async context => {
      const body = context.document.body;
      const paras = body.paragraphs;
      paras.load('items');
      const sel = context.document.getSelection();
      sel.load('text,styleBuiltIn');
      await context.sync();
      const headingItems = paras.items.slice(0, 300);
      headingItems.forEach(p => p.load('text,styleBuiltIn'));
      await context.sync();
      const outline = [];
      headingItems.forEach((p, idx) => {
        const style = String(p.styleBuiltIn || '');
        if (/heading/i.test(style) && (p.text || '').trim()) {
          outline.push({ index: idx, level: parseInt(style.replace(/\D/g, ''), 10) || 1, text: clampText(p.text.trim(), 120), refId: `p:${idx}` });
        }
      });
      return { success: true, documentId: App.state.workbookId || 'document', paragraphCount: paras.items.length, selection: { text: clampText(sel.text || '', 200), style: String(sel.styleBuiltIn || '') }, outline };
    });
  }

  const SYSTEM_PROMPT = `You are an AI assistant integrated into Microsoft Word with full access to read and modify the document.

Available tools:
READ:
- get_document_outline: Read heading structure and paragraph summaries
- get_selection: Read the currently selected text and its formatting
- get_paragraphs: Read a range of paragraphs by index
- get_tables: Read table contents
- search_text: Find text (supports match case / whole word / wildcards)

WRITE:
- insert_text: Insert text at Start/End of document, or Replace/Before/After the selection
- replace_text: Find and replace text (replaceAll optional)
- apply_style: Apply a built-in style (e.g. Heading1, Normal, Quote) and font formatting to selection or whole document
- set_paragraph_format: Alignment / line spacing / indent
- insert_heading: Insert a heading paragraph at a given level
- insert_table: Insert a table from a 2D array of rows
- insert_page_break: Insert a page break
- insert_image: Insert an inline image from base64
- manage_comment: Add a comment to the selection
- manage_content_control: Add/list content controls (use as named, navigable anchors)
- eval_officejs: Execute Word.run code when the listed tools are not enough

Citations: Use markdown links with #cite: hash to reference document locations. Clicking navigates there.
- Paragraph by index: [intro](#cite:p:0)
- Content control by tag: [summary](#cite:cc:summary)
Example: [see the introduction](#cite:p:0)

Read the document before making large edits. Be concise. Before overwriting or replacing existing content, confirm unless the user explicitly asks to replace.`;

  const TOOL_DEFINITIONS = [
    { type: 'function', function: { name: 'get_document_outline', description: 'Read the document heading structure and a paragraph summary. Returns outline (headings with index/level) and paragraphs (index/text/style).', parameters: { type: 'object', properties: { maxParagraphs: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_selection', description: 'Read the currently selected text and its style/font.', parameters: { type: 'object', properties: { explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_paragraphs', description: 'Read a range of paragraphs by index.', parameters: { type: 'object', properties: { start: { type: 'number' }, count: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_tables', description: 'Read the contents of tables in the document.', parameters: { type: 'object', properties: { maxTables: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'search_text', description: 'Find text in the document. Supports matchCase, matchWholeWord, useWildcards.', parameters: { type: 'object', properties: { query: { type: 'string' }, matchCase: { type: 'boolean' }, matchWholeWord: { type: 'boolean' }, useWildcards: { type: 'boolean' }, maxResults: { type: 'number' }, explanation: { type: 'string' } }, required: ['query'] } } },
    { type: 'function', function: { name: 'insert_text', description: "WRITE. Insert text. location: Start/End (document) or Replace/Before/After (current selection). Optional built-in style.", parameters: { type: 'object', properties: { text: { type: 'string' }, location: { enum: ['Start', 'End', 'Replace', 'Before', 'After'] }, style: { type: 'string' }, explanation: { type: 'string' } }, required: ['text'] } } },
    { type: 'function', function: { name: 'replace_text', description: 'WRITE. Find and replace text. replaceAll defaults true.', parameters: { type: 'object', properties: { query: { type: 'string' }, replacement: { type: 'string' }, matchCase: { type: 'boolean' }, matchWholeWord: { type: 'boolean' }, useWildcards: { type: 'boolean' }, replaceAll: { type: 'boolean' }, explanation: { type: 'string' } }, required: ['query', 'replacement'] } } },
    { type: 'function', function: { name: 'apply_style', description: 'WRITE. Apply a built-in style and/or font formatting to the selection or whole document.', parameters: { type: 'object', properties: { target: { enum: ['selection', 'document'] }, style: { type: 'string' }, font: { type: 'object', properties: { bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' }, size: { type: 'number' }, name: { type: 'string' }, color: { type: 'string' } } }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'set_paragraph_format', description: 'WRITE. Set paragraph alignment / line spacing / indent / spacing.', parameters: { type: 'object', properties: { target: { enum: ['selection', 'document'] }, alignment: { enum: ['Left', 'Centered', 'Right', 'Justified'] }, lineSpacing: { type: 'number' }, leftIndent: { type: 'number' }, spaceBefore: { type: 'number' }, spaceAfter: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'insert_heading', description: 'WRITE. Insert a heading paragraph at a level (1-6).', parameters: { type: 'object', properties: { text: { type: 'string' }, level: { type: 'number' }, location: { enum: ['Start', 'End'] }, explanation: { type: 'string' } }, required: ['text'] } } },
    { type: 'function', function: { name: 'insert_table', description: 'WRITE. Insert a table from a 2D array of rows.', parameters: { type: 'object', properties: { rows: { type: 'array', items: { type: 'array', items: {} } }, location: { enum: ['Start', 'End'] }, headerRow: { type: 'boolean' }, explanation: { type: 'string' } }, required: ['rows'] } } },
    { type: 'function', function: { name: 'insert_page_break', description: 'WRITE. Insert a page break.', parameters: { type: 'object', properties: { location: { enum: ['Start', 'End', 'selection'] }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'insert_image', description: 'WRITE. Insert an inline image from base64 (no data: prefix).', parameters: { type: 'object', properties: { base64: { type: 'string' }, location: { enum: ['Start', 'End', 'selection'] }, explanation: { type: 'string' } }, required: ['base64'] } } },
    { type: 'function', function: { name: 'manage_comment', description: 'WRITE. Add a comment to the current selection.', parameters: { type: 'object', properties: { operation: { enum: ['add'] }, text: { type: 'string' }, explanation: { type: 'string' } }, required: ['text'] } } },
    { type: 'function', function: { name: 'manage_content_control', description: 'Add/list content controls. Added controls act as named, navigable anchors (use tag in #cite:cc:tag).', parameters: { type: 'object', properties: { operation: { enum: ['add', 'list'] }, tag: { type: 'string' }, title: { type: 'string' }, text: { type: 'string' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'eval_officejs', description: 'Execute arbitrary Office.js code in Word.run. Escape hatch. Code receives context and Word.', parameters: { type: 'object', properties: { code: { type: 'string' }, explanation: { type: 'string' } }, required: ['code'] } } }
  ];

  const SAMPLE_ARGS = {
    get_document_outline: { maxParagraphs: 400 },
    get_selection: {},
    get_paragraphs: { start: 0, count: 20 },
    get_tables: { maxTables: 20 },
    search_text: { query: 'keyword', matchCase: false },
    insert_text: { text: '这是新插入的一段文字。', location: 'End' },
    replace_text: { query: '旧文字', replacement: '新文字', replaceAll: true },
    apply_style: { target: 'selection', style: 'Heading1', font: { bold: true } },
    set_paragraph_format: { target: 'selection', alignment: 'Justified', lineSpacing: 18 },
    insert_heading: { text: '小结', level: 1, location: 'End' },
    insert_table: { rows: [['姓名', '分数'], ['张三', 90], ['李四', 85]], location: 'End', headerRow: true },
    insert_page_break: { location: 'End' },
    insert_image: { base64: '<BASE64_IMAGE>', location: 'End' },
    manage_comment: { operation: 'add', text: '请复核此处。' },
    manage_content_control: { operation: 'add', tag: 'summary', title: '总结', text: '总结内容' },
    eval_officejs: { code: "const sel = context.document.getSelection();\nsel.load('text');\nawait context.sync();\nreturn sel.text;" }
  };
  function defaultArgsForTool(name) { return App.pretty(SAMPLE_ARGS[name] || {}); }

  App.HOSTS.word = {
    hostType: 'word',
    available: true,
    metadataLabel: 'Document outline',
    systemPrompt: SYSTEM_PROMPT,
    toolDefinitions: TOOL_DEFINITIONS,
    toolExecutors: TOOL_EXECUTORS,
    defaultArgsForTool,
    evalToolName: 'eval_officejs',
    getMetadata: getDocumentMetadata,
    navigateCitation,
    follow: maybeFollow,
    i18n: {
      zh: {
        brand: 'dpoqb in Word', brandFooter: 'dpoqb in Word · Plain Edition',
        title: '准备好处理你的 Word 文档', subtitle: '你可以让我撰写、润色、排版或检索文档内容',
        input: '告诉我你想如何处理这份文档…',
        chart: '智能写作生成', chartDesc: '根据要求生成或续写段落',
        fix: '文档润色校对', fixDesc: '检查并修正措辞、语法与格式',
        analyze: '文档结构解析', analyzeDesc: '提炼大纲、生成摘要与结论',
        chartPrompt: '请根据当前文档主题，在文末续写一段合适的内容',
        fixPrompt: '帮我校对全文，修正语法和措辞问题并给出修改建议',
        analyzePrompt: '帮我读取文档大纲并生成一段整体摘要',
        demo: '当前不在 Word/Office 环境中，Word 工具只能在插件侧边栏里运行。',
        confirmEval: 'AI 请求执行以下 Office.js 代码，是否允许？'
      },
      en: {
        brand: 'dpoqb in Word', brandFooter: 'dpoqb in Word · Plain Edition',
        title: 'Ready to work with your Word document', subtitle: 'Ask me to write, edit, format, or search the document',
        input: 'Tell me what to do with this document…',
        chart: 'Writing Generation', chartDesc: 'Draft or continue paragraphs',
        fix: 'Proofread & Polish', fixDesc: 'Fix grammar, wording and formatting',
        analyze: 'Document Analysis', analyzeDesc: 'Extract outline and summarize',
        chartPrompt: 'Continue the document with a fitting new paragraph at the end',
        fixPrompt: 'Proofread the whole document and fix grammar and wording issues',
        analyzePrompt: 'Read the document outline and write an overall summary',
        demo: 'Not currently running inside Word/Office. Word tools only work in the add-in task pane.',
        confirmEval: 'The AI requests to run the following Office.js code. Allow it?'
      }
    }
  };
})();

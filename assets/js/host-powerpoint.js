(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const requireOffice = () => App.requireOffice();
  function clampText(s, n) { s = String(s == null ? '' : s); return s.length > n ? s.slice(0, n) + '…' : s; }

  // 读取形状文本：不同平台 shape.textFrame 兼容性不一，统一防御性处理。
  function shapeText(shape) {
    try {
      const tf = shape.textFrame;
      if (tf && tf.textRange && typeof tf.textRange.text === 'string') return tf.textRange.text;
    } catch {}
    return '';
  }

  async function getPresentationOutline(args = {}) {
    requireOffice();
    const maxSlides = Math.max(1, Number(args.maxSlides || 100));
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const items = slides.items.slice(0, maxSlides);
      items.forEach(s => { s.load('id'); s.shapes.load('items'); });
      await context.sync();
      items.forEach(s => s.shapes.items.forEach(sh => { try { sh.textFrame.textRange.load('text'); } catch {} }));
      await context.sync();
      const out = items.map((s, idx) => {
        const texts = s.shapes.items.map(shapeText).map(x => x.trim()).filter(Boolean);
        return { index: idx, id: s.id, title: clampText(texts[0] || '', 120), texts: texts.map(x => clampText(x, 200)) };
      });
      return { success: true, slideCount: slides.items.length, returned: out.length, hasMore: slides.items.length > out.length, slides: out };
    });
  }

  async function getSlide(args) {
    requireOffice();
    const index = Number(args.index);
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      slide.shapes.load('items');
      await context.sync();
      slide.shapes.items.forEach(sh => { sh.load('id,name,type,left,top,width,height'); try { sh.textFrame.textRange.load('text'); } catch {} });
      await context.sync();
      const shapes = slide.shapes.items.map(sh => ({ id: sh.id, name: sh.name, type: String(sh.type || ''), text: clampText(shapeText(sh), 300), left: sh.left, top: sh.top, width: sh.width, height: sh.height }));
      return { success: true, index, id: slide.id, shapes };
    });
  }

  async function getSelectedSlides() {
    requireOffice();
    return new Promise(resolve => {
      try {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const slides = (asyncResult.value && asyncResult.value.slides) || [];
            resolve({ success: true, slides: slides.map(s => ({ id: s.id, title: s.title, index: s.index })) });
          } else {
            resolve({ success: false, error: asyncResult.error ? asyncResult.error.message : 'getSelectedDataAsync failed' });
          }
        });
      } catch (e) { resolve({ success: false, error: e.message || String(e) }); }
    });
  }

  async function addSlide(args = {}) {
    requireOffice();
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.add(); // 默认版式
      slides.load('items');
      await context.sync();
      const newIndex = slides.items.length - 1;
      const slide = slides.items[newIndex];
      slide.load('id');
      await context.sync();
      return { success: true, index: newIndex, id: slide.id, _navTarget: { slideId: slide.id } };
    });
  }

  async function deleteSlide(args) {
    requireOffice();
    const index = Number(args.index);
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.delete();
      await context.sync();
      return { success: true, deletedIndex: index };
    });
  }

  async function duplicateSlide(args) {
    requireOffice();
    const index = Number(args.index);
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      if (typeof slide.duplicate !== 'function') throw new Error('Slide.duplicate is not available in this PowerPoint host; use eval_officejs');
      slide.duplicate();
      await context.sync();
      return { success: true, sourceIndex: index };
    });
  }

  async function setSlideNotes(args) {
    requireOffice();
    const index = Number(args.index);
    const text = String(args.text || '');
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      let notes;
      try { notes = slide.notesSlide; notes.shapes.load('items'); await context.sync(); }
      catch (e) { throw new Error('Notes slide API not available in this PowerPoint host; use eval_officejs'); }
      // 用第一个含 textFrame 的形状作为备注占位符
      let target = null;
      for (const sh of notes.shapes.items) { try { sh.textFrame.textRange.load('text'); target = target || sh; } catch {} }
      await context.sync();
      if (!target) throw new Error('No notes placeholder found on this slide');
      target.textFrame.textRange.text = text;
      await context.sync();
      return { success: true, index, _navTarget: { slideId: slide.id } };
    });
  }

  async function insertTextbox(args) {
    requireOffice();
    const { index, text = '', left = 50, top = 50, width = 400, height = 100 } = args;
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[Number(index)];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      if (!slide.shapes || typeof slide.shapes.addTextBox !== 'function') throw new Error('shapes.addTextBox not available in this PowerPoint host; use eval_officejs');
      slide.shapes.addTextBox(text, { left, top, width, height });
      await context.sync();
      return { success: true, index: Number(index), _navTarget: { slideId: slide.id } };
    });
  }

  async function setText(args) {
    requireOffice();
    const { index, shapeId, text = '' } = args;
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[Number(index)];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      slide.shapes.load('items');
      await context.sync();
      slide.shapes.items.forEach(sh => sh.load('id'));
      await context.sync();
      const shape = shapeId ? slide.shapes.items.find(sh => sh.id === shapeId) : slide.shapes.items[0];
      if (!shape) throw new Error('Target shape not found');
      shape.textFrame.textRange.text = text;
      await context.sync();
      return { success: true, index: Number(index), _navTarget: { slideId: slide.id } };
    });
  }

  async function insertImage(args) {
    requireOffice();
    const { base64 } = args;
    if (!base64) throw new Error('base64 image data is required');
    return new Promise(resolve => {
      try {
        Office.context.document.setSelectedDataAsync(base64, { coercionType: Office.CoercionType.Image }, r => {
          if (r.status === Office.AsyncResultStatus.Succeeded) resolve({ success: true });
          else resolve({ success: false, error: r.error ? r.error.message : 'setSelectedDataAsync failed' });
        });
      } catch (e) { resolve({ success: false, error: e.message || String(e) }); }
    });
  }

  async function gotoSlide(args) {
    requireOffice();
    const index = Number(args.index);
    const slideId = await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const slide = slides.items[index];
      if (!slide) throw new Error(`Slide ${index} not found`);
      slide.load('id');
      await context.sync();
      return slide.id;
    });
    const navigated = await selectSlideById(slideId);
    return { success: navigated, index, id: slideId, navigated };
  }

  function selectSlideById(slideId) {
    return new Promise(resolve => {
      try {
        const doc = Office.context.document;
        if (typeof doc.goToByIdAsync !== 'function') return resolve(false);
        doc.goToByIdAsync(slideId, Office.GoToType.Slide, r => resolve(r.status === Office.AsyncResultStatus.Succeeded));
      } catch { resolve(false); }
    });
  }

  async function evalOfficeJs(args) {
    requireOffice();
    const code = args.code || '';
    return PowerPoint.run(async context => {
      const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
      const fn = new AsyncFunction('context', 'PowerPoint', code);
      const result = await fn(context, PowerPoint);
      return { success: true, result: result ?? null };
    });
  }

  async function navigateCitation(ref) {
    requireOffice();
    // 格式: "s:<index>" 幻灯片序号 | "id:<slideId>" 幻灯片 id | 纯数字回退为序号
    if (/^id:/.test(ref)) { const navigated = await selectSlideById(ref.slice(3)); return { success: navigated, navigated }; }
    const index = Number(/^s:/.test(ref) ? ref.slice(2) : ref);
    return gotoSlide({ index });
  }

  async function maybeFollow(result) {
    if (!App.state.settings.followMode || !result) return;
    const nav = result._navTarget;
    if (nav && nav.slideId) { await selectSlideById(nav.slideId).catch(console.warn); }
  }

  const TOOL_EXECUTORS = {
    get_presentation_outline: getPresentationOutline,
    get_slide: getSlide,
    get_selected_slides: getSelectedSlides,
    add_slide: addSlide,
    delete_slide: deleteSlide,
    duplicate_slide: duplicateSlide,
    set_slide_notes: setSlideNotes,
    insert_textbox: insertTextbox,
    set_text: setText,
    insert_image: insertImage,
    goto_slide: gotoSlide,
    eval_officejs: evalOfficeJs
  };

  async function getPresentationMetadata() {
    requireOffice();
    return PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      const items = slides.items.slice(0, 60);
      items.forEach(s => { s.load('id'); s.shapes.load('items'); });
      await context.sync();
      items.forEach(s => s.shapes.items.forEach(sh => { try { sh.textFrame.textRange.load('text'); } catch {} }));
      await context.sync();
      const out = items.map((s, idx) => {
        const texts = s.shapes.items.map(shapeText).map(x => x.trim()).filter(Boolean);
        return { index: idx, id: s.id, title: clampText(texts[0] || '', 100), refId: `s:${idx}` };
      });
      return { success: true, presentationId: App.state.workbookId || 'presentation', slideCount: slides.items.length, slides: out };
    });
  }

  const SYSTEM_PROMPT = `You are an AI assistant integrated into Microsoft PowerPoint with access to read and modify the presentation.

Available tools:
READ:
- get_presentation_outline: List slides with their titles and text
- get_slide: Read all shapes and text on a specific slide (by index)
- get_selected_slides: Get the currently selected slides

WRITE:
- add_slide: Add a new slide
- delete_slide: Delete a slide by index
- duplicate_slide: Duplicate a slide by index
- set_slide_notes: Set the speaker notes of a slide
- insert_textbox: Add a text box to a slide (position/size in points)
- set_text: Replace the text of a shape on a slide
- insert_image: Insert a base64 image into the current selection
- goto_slide: Navigate to a slide by index
- eval_officejs: Execute PowerPoint.run code when the listed tools are not enough

Note: The PowerPoint Office.js API is thinner than Word/Excel. For advanced layout, precise styling, or operations not covered above, fall back to eval_officejs (and OOXML where needed). Some shape/notes operations vary by platform (web vs desktop); tools defend against missing APIs and will tell you to use eval_officejs when something is unavailable.

Citations: Use markdown links with #cite: hash to reference slides. Clicking navigates there.
- Slide by index: [agenda](#cite:s:1)
Example: [see slide 2](#cite:s:1)

Read the presentation before making large edits. Be concise.`;

  const TOOL_DEFINITIONS = [
    { type: 'function', function: { name: 'get_presentation_outline', description: 'List slides with titles and text content.', parameters: { type: 'object', properties: { maxSlides: { type: 'number' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'get_slide', description: 'Read all shapes and text on a slide by index.', parameters: { type: 'object', properties: { index: { type: 'number' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'get_selected_slides', description: 'Get the currently selected slide(s).', parameters: { type: 'object', properties: { explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'add_slide', description: 'WRITE. Add a new slide at the end.', parameters: { type: 'object', properties: { explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'delete_slide', description: 'WRITE. Delete a slide by index.', parameters: { type: 'object', properties: { index: { type: 'number' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'duplicate_slide', description: 'WRITE. Duplicate a slide by index.', parameters: { type: 'object', properties: { index: { type: 'number' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'set_slide_notes', description: 'WRITE. Set the speaker notes text of a slide.', parameters: { type: 'object', properties: { index: { type: 'number' }, text: { type: 'string' }, explanation: { type: 'string' } }, required: ['index', 'text'] } } },
    { type: 'function', function: { name: 'insert_textbox', description: 'WRITE. Add a text box to a slide. Position/size in points.', parameters: { type: 'object', properties: { index: { type: 'number' }, text: { type: 'string' }, left: { type: 'number' }, top: { type: 'number' }, width: { type: 'number' }, height: { type: 'number' }, explanation: { type: 'string' } }, required: ['index', 'text'] } } },
    { type: 'function', function: { name: 'set_text', description: 'WRITE. Replace the text of a shape on a slide (first shape if shapeId omitted).', parameters: { type: 'object', properties: { index: { type: 'number' }, shapeId: { type: 'string' }, text: { type: 'string' }, explanation: { type: 'string' } }, required: ['index', 'text'] } } },
    { type: 'function', function: { name: 'insert_image', description: 'WRITE. Insert a base64 image (no data: prefix) into the current selection.', parameters: { type: 'object', properties: { base64: { type: 'string' }, explanation: { type: 'string' } }, required: ['base64'] } } },
    { type: 'function', function: { name: 'goto_slide', description: 'Navigate to a slide by index.', parameters: { type: 'object', properties: { index: { type: 'number' }, explanation: { type: 'string' } }, required: ['index'] } } },
    { type: 'function', function: { name: 'eval_officejs', description: 'Execute arbitrary Office.js code in PowerPoint.run. Escape hatch for unsupported operations. Code receives context and PowerPoint.', parameters: { type: 'object', properties: { code: { type: 'string' }, explanation: { type: 'string' } }, required: ['code'] } } }
  ];

  const SAMPLE_ARGS = {
    get_presentation_outline: { maxSlides: 100 },
    get_slide: { index: 0 },
    get_selected_slides: {},
    add_slide: {},
    delete_slide: { index: 1 },
    duplicate_slide: { index: 0 },
    set_slide_notes: { index: 0, text: '这一页的讲稿备注。' },
    insert_textbox: { index: 0, text: '标题文字', left: 50, top: 50, width: 400, height: 80 },
    set_text: { index: 0, text: '替换后的文字' },
    insert_image: { base64: '<BASE64_IMAGE>' },
    goto_slide: { index: 2 },
    eval_officejs: { code: "const slides = context.presentation.slides;\nslides.load('items');\nawait context.sync();\nreturn slides.items.length;" }
  };
  function defaultArgsForTool(name) { return App.pretty(SAMPLE_ARGS[name] || {}); }

  App.HOSTS.powerpoint = {
    hostType: 'powerpoint',
    available: true,
    metadataLabel: 'Presentation outline',
    systemPrompt: SYSTEM_PROMPT,
    toolDefinitions: TOOL_DEFINITIONS,
    toolExecutors: TOOL_EXECUTORS,
    defaultArgsForTool,
    evalToolName: 'eval_officejs',
    getMetadata: getPresentationMetadata,
    navigateCitation,
    follow: maybeFollow,
    i18n: {
      zh: {
        brand: 'dpoqb in PowerPoint', brandFooter: 'dpoqb in PowerPoint · Plain Edition',
        title: '准备好处理你的演示文稿', subtitle: '你可以让我生成幻灯片、撰写内容或整理结构',
        input: '告诉我你想如何处理这份演示文稿…',
        chart: '幻灯片生成', chartDesc: '按主题快速生成幻灯片与文本',
        fix: '内容润色优化', fixDesc: '精炼要点、统一措辞与结构',
        analyze: '演示结构解析', analyzeDesc: '提炼大纲、生成讲稿备注',
        chartPrompt: '请根据当前演示文稿主题，新增一页合适的幻灯片',
        fixPrompt: '帮我精炼现有幻灯片的要点文字，使其更简洁有力',
        analyzePrompt: '帮我读取演示文稿大纲并总结整体结构',
        demo: '当前不在 PowerPoint/Office 环境中，PPT 工具只能在插件侧边栏里运行。',
        confirmEval: 'AI 请求执行以下 Office.js 代码，是否允许？'
      },
      en: {
        brand: 'dpoqb in PowerPoint', brandFooter: 'dpoqb in PowerPoint · Plain Edition',
        title: 'Ready to work with your presentation', subtitle: 'Ask me to create slides, write content, or organize structure',
        input: 'Tell me what to do with this presentation…',
        chart: 'Slide Generation', chartDesc: 'Generate slides and text by topic',
        fix: 'Content Polish', fixDesc: 'Refine bullets, unify wording & structure',
        analyze: 'Deck Analysis', analyzeDesc: 'Extract outline and write speaker notes',
        chartPrompt: 'Add a fitting new slide based on the presentation topic',
        fixPrompt: 'Refine the bullet text on the existing slides to be more concise',
        analyzePrompt: 'Read the presentation outline and summarize the overall structure',
        demo: 'Not currently running inside PowerPoint/Office. PowerPoint tools only work in the add-in task pane.',
        confirmEval: 'The AI requests to run the following Office.js code. Allow it?'
      }
    }
  };
})();

# dpoqb in Office - Plain HTML/JavaScript/CSS Rebuild

这是对原始 `office-addin.bigmodel.cn` 编译产物的可读版重构。核心页面已改为纯 `HTML + JavaScript + CSS`，没有 React/Vite 打包入口，便于直接阅读和二次修改。

现已支持 **Excel / Word / PowerPoint 三个宿主**：一套代码包按 `Office.context.host` 自适应，每个宿主一份 manifest（指向同一 `taskpane.html`）。

## 文件说明

- `taskpane.html`：插件侧边栏入口，按顺序加载下列脚本。
- `assets/js/`：核心逻辑，按职责拆分为多个文件，统一挂载到全局 `window.App` 命名空间：
  - `config.js`：常量、`PROVIDERS`、宿主中性的 `I18N`、`STORAGE_KEYS`。
  - `state.js`：全局 `state`、localStorage/会话存取、宿主无关的 `t()`/`hasOffice()` 等工具函数。
  - `markdown.js`：对话气泡的 Markdown 渲染（宿主无关）。
  - `host.js`：**宿主适配层**。宿主检测/选择、通用 Office 文档设置存取、`executeToolByName`/`maybeFollow`/`navigateCitation` 委托到当前 `App.host`。
  - `host-excel.js` / `host-word.js` / `host-powerpoint.js`：各宿主的提供者实现（工具集 `toolExecutors`、`toolDefinitions`、`systemPrompt`、`getMetadata`、`navigateCitation`、`follow`、宿主专属 i18n）。运行时只有一个被选中并挂到 `App.host`。
  - `api.js`：OpenAI-compatible 接口调用与 Agent Loop（宿主无关，走 `App.host.*`）。
  - `ui.js`：界面渲染与 DOM 事件绑定（宿主无关，文案/工具列表按 `App.host` 动态读取）。
  - `app.js`：入口，`Office.onReady` 后按宿主类型选定 `App.host` 并首屏渲染。
- `assets/css/app.css`：侧边栏样式。
- `commands.html` / `assets/js/commands.js`：Ribbon 命令文件（宿主无关）。
- `manifest.excel.prod.xml` / `manifest.excel.local.xml`：**Excel** 清单（`Host Name="Workbook"`）。
- `manifest.word.prod.xml` / `manifest.word.local.xml`：**Word** 清单（`Host Name="Document"`）。
- `manifest.ppt.prod.xml` / `manifest.ppt.local.xml`：**PowerPoint** 清单（`Host Name="Presentation"`）。
  - 三套 manifest 的 `<Id>` 各自唯一，`SourceLocation` 均指向同一 `taskpane.html`。`*.prod.xml` 指向 `https://dpoqb.top`，`*.local.xml` 指向 `http://localhost`。
- `dpoqb-in-excel.html`：原安装说明页。
- `assets/vendor/`：原项目引用/打包涉及的第三方或运行时资源，随包保留。

## 宿主适配层

`App.host` 是一个提供者对象，统一了各宿主的能力接口：

```
hostType, available, metadataLabel, systemPrompt,
toolDefinitions, toolExecutors, defaultArgsForTool(name),
evalToolName, getMetadata(), navigateCitation(ref), follow(result), i18n
```

`api.js`/`ui.js`/`app.js` 一律调用 `App.host.*` 或 `host.js` 提供的委托（`App.executeToolByName` 等），新增宿主只需新增一个 `host-*.js` 并注册到 `App.HOSTS`，无需改动通用层。

## 各宿主能力

### Excel（`Excel.run`）
- 读取：`get_cell_ranges` / `get_range_as_csv` / `search_data` / `get_all_objects`
- 写入：`set_cell_range` / `clear_cell_range` / `copy_to` / `modify_sheet_structure` / `modify_workbook_structure` / `resize_range` / `modify_object`
- 引用 `#cite:sheetId` 或 `#cite:sheetId!A1:B2`
- `eval_officejs` 兜底

### Word（`Word.run`）
- 读取：`get_document_outline` / `get_selection` / `get_paragraphs` / `get_tables` / `search_text`
- 写入：`insert_text` / `replace_text` / `apply_style` / `set_paragraph_format` / `insert_heading` / `insert_table` / `insert_page_break` / `insert_image` / `manage_comment` / `manage_content_control`
- 引用 `#cite:p:<段落序号>` 或 `#cite:cc:<内容控件tag>`
- `eval_officejs` 兜底

### PowerPoint（`PowerPoint.run` + Common API）
- 读取：`get_presentation_outline` / `get_slide` / `get_selected_slides`
- 写入：`add_slide` / `delete_slide` / `duplicate_slide` / `set_slide_notes` / `insert_textbox` / `set_text` / `insert_image` / `goto_slide`
- 引用 `#cite:s:<幻灯片序号>`
- `eval_officejs` 兜底。PowerPoint 的 Office.js API 比 Word/Excel 薄，部分形状/备注操作在 Web/桌面端兼容性不一，工具会防御性检测并在不可用时提示改用 `eval_officejs`。

## 通用能力

1. API 配置：dpoqb / OpenAI Compatible / DeepSeek / OpenRouter / Groq / Custom OpenAI Compatible。
2. 本地保存 API Key、模型、Base URL、语言、主题、会话。
3. Agent Loop：OpenAI-compatible `chat/completions` 工具调用循环。
4. “工具”页可手动选择当前宿主的工具并输入 JSON 参数直接执行。
5. `#cite:` 引用点击导航（格式由各宿主定义）。
6. 深色/浅色主题、中英文切换、跟随模式（写操作后视图跟随到改动位置）。

## 本地调试

Office Add-in 需要 HTTP 静态服务器。把对应宿主 `manifest.*.local.xml` 中的地址改为实际地址，并旁加载到对应的 Office 应用：

```bash
cd office-addin-plain
# 可用任意 HTTP 静态服务，端口与 manifest.excel.local.xml 保持一致
```

- Excel → `manifest.excel.local.xml`
- Word → `manifest.word.local.xml`
- PowerPoint → `manifest.ppt.local.xml`

将对应 manifest 放入 Office 的 Wef 目录进行旁加载。

## 注意

- 原包中的 `office.js` 来自 Microsoft 官方 CDN：`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`。这是 Office Add-in 官方要求的运行时脚本，未内置到本地包。
- 一个 manifest 只能绑定一个宿主，因此三宿主需各一份 manifest；但它们共用同一份 `taskpane.html` 与 `assets/`，运行时按 `Office.context.host` 自动选择工具集。
- 图表/透视表（Excel）、内容控件/批注（Word）、形状/备注（PowerPoint）等 API 在不同 Excel/Word/PowerPoint 版本与 Web/Desktop 环境中兼容性可能不同；如某个 Office.js API 在目标环境不可用，可通过 `eval_officejs` 工具补充实现。

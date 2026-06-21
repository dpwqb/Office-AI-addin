# dpoqb in Excel - Plain HTML/JavaScript/CSS Rebuild

这是对原始 `office-addin.bigmodel.cn` 编译产物的可读版重构。核心页面已改为纯 `HTML + JavaScript + CSS`，没有 React/Vite 打包入口，便于直接阅读和二次修改。

## 文件说明

- `taskpane.html`：Excel 插件侧边栏入口。
- `assets/js/app.js`：核心逻辑，包含对话 UI、API 配置、Agent Loop、Excel 工具函数、手动工具执行器。
- `assets/css/app.css`：侧边栏样式。
- `commands.html` / `assets/js/commands.js`：Ribbon 命令文件。
- `manifest.prod.xml`：保留原生产域名 `https://dpoqb.top` 的清单。
- `manifest.local.xml`：本地 HTTP 调试示例，默认指向 `http://localhost`。
- `dpoqb-in-excel.html`：原安装说明页，保留样式并将 manifest 下载链接改成本地相对路径。
- `assets/vendor/`：原项目引用/打包涉及的第三方或运行时资源，随包保留。
- `assets/original/`：原始编译后的 `taskpane.js/css` 作为对照参考，未被新版页面引用。

## 已重构/启用能力

1. Excel AI 侧边栏对话。
2. API 配置：dpoqb / OpenAI Compatible / DeepSeek / OpenRouter / Groq / Custom OpenAI Compatible。
3. 本地保存 API Key、模型、Base URL、语言、主题、会话。
4. Agent Loop：支持 OpenAI-compatible `chat/completions` 工具调用循环。
5. Excel 读取工具：
   - `get_cell_ranges`
   - `get_range_as_csv`
   - `search_data`
   - `get_all_objects`
6. Excel 写入/修改工具：
   - `set_cell_range`
   - `clear_cell_range`
   - `copy_to`
   - `modify_sheet_structure`
   - `modify_workbook_structure`
   - `resize_range`
   - `modify_object`
7. 隐藏/高级能力已暴露：
   - `eval_officejs` 可执行 Office.js 代码。
   - 新增“工具”页，可手动选择工具并输入 JSON 参数直接执行。
8. 支持 `#cite:sheetId!A1:B2` 单元格引用导航。
9. 支持深色/浅色主题、中英文切换、跟随模式。

## 本地调试

Office Add-in 需要 HTTP 静态服务器，把 `manifest.local.xml` 中的地址改为实际地址。

示例：

```bash
cd office-addin-plain
# 可用任意 HTTP 静态服务，端口与 manifest.local.xml 保持一致
```

然后将 `manifest.local.xml` 放入 Excel Wef 目录进行旁加载。

## 注意

- 原包中的 `office.js` 来自 Microsoft 官方 CDN：`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`。这是 Office Add-in 官方要求的运行时脚本，未内置到本地包。
- 原始 2.5MB `taskpane.js` 包含大量 SDK/运行时/模型适配代码。新版已按原功能重写核心流程，但没有逐字反编译所有第三方 SDK 内部实现。
- 图表与透视表 API 在不同 Excel 版本/Web/Desktop 环境中兼容性可能不同；如某个 Office.js API 在目标环境不可用，可通过已启用的 `eval_officejs` 工具补充实现。

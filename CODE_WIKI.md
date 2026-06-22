# WPS Office MCP - Code Wiki

> **项目口号**：让 AI 操控 WPS 就像老王操控键盘一样丝滑
> **版本**：1.0.0
> **运行环境**：Node.js ≥ 18.0.0，WPS Office（Mac / Windows / Linux）

---

## 1. 项目概览

这是一个基于 **MCP（Model Context Protocol）** 的 WPS Office 自动化工具集。它通过 `stdio` 协议将 AI 大模型（Claude 等）与本地 WPS 办公软件桥接起来，使 AI 能以自然语言的方式读取、编辑、格式化、制作 Excel 表格 / Word 文档 / PPT 演示文稿。

### 1.1 核心能力

| 应用 | 工具数量 | 主要能力 |
| --- | --- | --- |
| **Excel** | 82+ | 公式、数据、图表、透视表、工作表、格式化、行列操作、批注保护、高级数据处理 |
| **Word** | 32+ | 格式化、内容编辑、段落样式、目录生成、批注、书签、页眉页脚、校对 |
| **PPT** | 112+ | 幻灯片、形状、文本框、图片、图表、动画、切换、3D 效果、数据可视化组件、流程图、组织架构图、时间线、母版 |
| **Common** | 9+ | 跨应用数据缓存、PDF 转换、通用保存/打开、连接探测 |
| **内置工具** | 12 个 | `wps_check_connection`、`wps_get_active_document`、`wps_insert_text`、`wps_cache_data` 等 |

> **总计**：**247+ 个 MCP 工具**

### 1.2 架构分层（从上到下）

```
+--------------------- AI 应用层 ---------------------+
|  Claude Desktop / IDE AI Assistant (通过 MCP 对话)   |
+--------------------- MCP 服务层 ---------------------+
|  [mcp-server.ts] MCP Server (stdio transport)        |
|  └─ [tool-registry.ts] Tool 注册表 / 生命周期管理     |
|  └─ [logger.ts] 结构化日志                           |
+--------------------- 工具实现层 ---------------------+
|  src/tools/                                          |
|  ├─ excel/    ├─ word/    ├─ ppt/    ├─ common/      |
|  └─ (index.ts) 聚合导出                              |
+--------------------- WPS 客户端层 --------------------+
|  [wps-client.ts] 跨平台统一调用接口                   |
|  ├─ macOS / Linux → [mac-poll-server.ts] HTTP 轮询   |
|  │                        ↕ (http://127.0.0.1:58891) |
|  │                  [wps-claude-assistant/main.js]   |
|  │                   ↓ WPS JS 加载项 API             |
│  │                  Application.ActiveWorkbook / ... |
|  └─ Windows → PowerShell + WPS COM 自动化接口         |
+--------------------- 提示词层 ------------------------+
|  skills/wps-excel, wps-word, wps-ppt, wps-office     |
|   (SKILL.md，指导 AI 如何编排工具调用)                |
+------------------------------------------------------+
```

### 1.3 目录结构总览

```
/workspace
├── package.json                    # 顶层工程配置（scripts 代理到子工程）
├── INSTALL.md                      # 跨平台安装指南（macOS / Linux / Windows）
├── scripts/
│   ├── auto-install-mac.sh         # macOS 一键安装脚本
│   ├── auto-install.ps1            # Windows 一键安装脚本
│   ├── dev.ps1, install.ps1        # Windows 开发/安装辅助
│   └── install.sh                  # Linux 安装脚本
├── skills/
│   ├── wps-excel/SKILL.md          # Excel 提示词/编排规范
│   ├── wps-word/SKILL.md           # Word 提示词/编排规范
│   ├── wps-ppt/SKILL.md            # PPT 提示词/编排规范
│   └── wps-office/SKILL.md         # 通用 Office 提示词规范
├── wps-claude-assistant/           # macOS/Linux WPS 加载项（JS 插件）
│   ├── main.js                     # 轮询入口 + 命令分发（核心）
│   ├── manifest.xml                # WPS 加载项清单
│   ├── ribbon.xml                  # WPS 顶部菜单/Ribbon 定义
│   ├── index.html                  # 加载项宿主页面
│   ├── wps-auto.sh                 # 自动切换 WPS 应用脚本
│   ├── handlers/                   # 处理函数 README
│   └── utils/response.js           # 响应格式工具
├── wps-claude-addon/              # Windows 版 WPS 加载项（平行实现）
│   ├── js/main.js
│   ├── manifest.xml
│   └── ribbon.xml
└── wps-office-mcp/                 # MCP Server 主体（TypeScript）
    ├── package.json
    ├── tsconfig.json
    ├── jest.config.js
    ├── scripts/wps-com.ps1         # Windows COM 桥接脚本
    └── src/
        ├── index.ts                # 服务入口（main）
        ├── server/
        │   ├── mcp-server.ts       # MCP Server 核心实现
        │   └── tool-registry.ts    # Tool 注册/调度中心
        ├── client/
        │   ├── wps-client.ts       # 跨平台 WPS 客户端（invokeAction）
        │   ├── mac-poll-server.ts  # macOS HTTP 轮询服务器（:58891）
        │   └── wps-keepalive.ts    # WPS HTTP 服务保活（Windows）
        ├── tools/
        │   ├── index.ts            # 所有工具聚合入口
        │   ├── excel/              # Excel 82+ 个工具
        │   │   ├── index.ts, formula.ts, data.ts, pivot.ts,
        │   │   ├── chart.ts, sheet.ts, format.ts, workbook.ts,
        │   │   ├── data-advanced.ts, row-column.ts, comment-protect.ts
        │   ├── word/               # Word 32+ 个工具
        │   │   ├── index.ts, content.ts, document.ts,
        │   │   ├── format.ts, proofread.ts
        │   ├── ppt/                # PPT 112+ 个工具
        │   │   ├── index.ts, slide.ts, slide-ops.ts,
        │   │   ├── shape-basic.ts, textbox.ts, image.ts,
        │   │   ├── chart-flow.ts, data-viz.ts,
        │   │   ├── animation.ts, beautify-advanced.ts,
        │   │   ├── background.ts, misc.ts, presentation.ts, ...
        │   └── common/             # Common 9 个工具
        │       ├── index.ts, convert.ts, general.ts
        ├── types/
        │   ├── index.ts
        │   ├── tools.ts            # ToolDefinition / ToolHandler / ToolCallResult
        │   └── wps.ts              # WpsApiRequest / Response / AppType
        └── utils/
            ├── logger.ts           # Winston 日志（文件 + 可选 console）
            └── error.ts            # 业务错误封装
```

---

## 2. 核心模块详解

### 2.1 入口 `src/index.ts`

- **位置**：[wps-office-mcp/src/index.ts](file:///workspace/wps-office-mcp/src/index.ts)
- **作用**：
  1. 导出所有核心模块（`McpServer`、`ToolRegistry`、`WpsClient`、`logger`、类型）
  2. **启动 MCP Server**：调用 `createMcpServer()` 并建立 stdio transport
  3. **优雅关闭**：监听 `SIGINT` / `SIGTERM`，先停服务再退出
  4. **全局异常兜底**：`uncaughtException` / `unhandledRejection` → 日志 + `process.exit(1)`
- **关键约定**：通过 `require.main === module` 判断是否被直接调用；被 MCP 客户端加载时同样会走 `main()`。

### 2.2 MCP Server 核心 `src/server/mcp-server.ts`

- **位置**：[wps-office-mcp/src/server/mcp-server.ts](file:///workspace/wps-office-mcp/src/server/mcp-server.ts)
- **类**：`WpsMcpServer`（实例由 `createMcpServer()` 工厂创建）
- **核心实现**：
  - 基于 `@modelcontextprotocol/sdk/server` + `StdioServerTransport`
  - 支持 MCP 协议的两个核心请求类型：
    1. `ListToolsRequestSchema` / `tools/list`：返回注册表所有工具的 JSON Schema
    2. `CallToolRequestSchema` / `tools/call`：委托 `ToolRegistry.callTool()` 执行
  - 统一错误封装：业务 `McpError` 直接返回文本；其他错误映射为 `SdkMcpError(InternalError)`
- **跨应用数据缓存（P0 级特性）**：
  - MCP 本身每次调用是无状态的，且 macOS WPS 无法跨应用共享状态
  - `WpsMcpServer` 内部维护一个静态 `dataCache: Map<string, { data, timestamp, appType }>`
  - 通过 4 个内置工具暴露给 AI：
    - `wps_cache_data(key, data, appType)` — 写入缓存
    - `wps_get_cached_data(key)` — 读取缓存
    - `wps_list_cache()` — 列出所有键
    - `wps_clear_cache(key?)` — 清理缓存
  - **典型场景**：AI 从 Excel 读取数据 → 缓存到 MCP Server → 让 PPT 侧"获取缓存"创建图表幻灯片

### 2.3 Tool 注册表 `src/server/tool-registry.ts`

- **位置**：[wps-office-mcp/src/server/tool-registry.ts](file:///workspace/wps-office-mcp/src/server/tool-registry.ts)
- **单例模式**：`ToolRegistry.getInstance()` / 导出变量 `toolRegistry`
- **数据结构**：
  - `tools: Map<string, RegisteredTool>`
  - `categories: Map<ToolCategory, Set<string>>`
  - `ToolCategory ∈ { DOCUMENT, SPREADSHEET, PRESENTATION, COMMON }`
- **核心方法**：
  - `register(ToolDefinition, ToolHandler)`：新增单个工具，重复自动跳过
  - `registerAll([{ definition, handler }, ...])`：批量注册
  - `listTools(): { tools: ToolDefinition[] }`：MCP `tools/list` 使用
  - `callTool({ id, name, arguments })`：**核心调度入口**
    - 参数校验（必填项）
    - 执行 handler，捕获并封装错误
    - 记录执行耗时 + 成功/失败日志
    - 返回 `ToolCallResult = { id, success, content, error? }`
- **装饰器 `@RegisterTool(definition)`**：用于类方法上，自动把该方法作为 handler 注册
- **快捷函数**：`registerTool(def, handler)`

### 2.4 跨平台 WPS 客户端 `src/client/wps-client.ts`

- **位置**：[wps-office-mcp/src/client/wps-client.ts](file:///workspace/wps-office-mcp/src/client/wps-client.ts)
- **类**：`WpsClient`（单例 `wpsClient` 导出）
- **平台路由**（`execWpsAction` 内部）：
  - **Windows**：`execPowerShell()` → 调用 `scripts/wps-com.ps1`，`spawn powershell` 传递 JSON 参数，解析 stdout 为响应对象
  - **macOS / Linux**：`execMacPoll()` → `macPollServer.executeCommand(action, params)`，HTTP 轮询交给 WPS 加载项执行
- **对外方法（业务语义）**：
  - `checkConnection()` → `ping`
  - `getActiveWorkbook() / getActiveDocument() / getActivePresentation()`
  - `getCellValue(sheet, row, col)`、`setCellValue(sheet, row, col, value)`
  - `getRangeData(sheet, range)`、`setRangeData(sheet, range, data[][])`
  - `setFormula(sheet, row, col, formula)`
  - `insertText(text, position?)`、`getDocumentText()`
  - `addSlide(layout?)`、`createPresentation()`
  - `openFile(path)`、`saveFile()`、`saveFileAs(path)`
  - `executeMethod(method, params?, appType?)` — **万能入口**，所有高级工具走这里
  - `callApi(request)` — 兼容旧 API（带 method 映射表）

### 2.5 macOS / Linux HTTP 轮询服务器 `src/client/mac-poll-server.ts`

- **位置**：[wps-office-mcp/src/client/mac-poll-server.ts](file:///workspace/wps-office-mcp/src/client/mac-poll-server.ts)
- **架构背景**：WPS for Mac 的 JS 加载项运行在沙箱，**无法**直接启动 HTTP 服务器，因此反转角色：
  1. MCP Server 进程启动 `http://127.0.0.1:58891`
  2. WPS 加载项 `setInterval`（500ms）轮询 `GET /poll`
  3. 如果队列有命令 → 返回 `{ command: { action, params, requestId } }`
  4. 加载项通过 WPS 的 `Application` 对象执行后，`POST /result` 把结果回传
  5. `executeCommand()` 的 Promise 被 resolve，调用链继续向上
- **端口**：默认 `58891`，支持环境变量 `WPS_MCP_PORT` 覆盖
- **端点**：
  - `GET /poll` — 命令拉取
  - `POST /result` — 结果回传（body: JSON `{ requestId, result }`）
  - `GET /status` — 健康检查
  - `OPTIONS *` — CORS 预检
- **应用切换**：`MacPollServer.switchApp(app)` 调用 `wps-auto.sh`，解决"执行 Excel 命令后再执行 PPT 命令但 WPS 主应用仍然是 Excel"的问题
- **命令 → 应用映射表**：`COMMAND_APP_MAP`（300+ 键），声明每个 action 需要在哪种应用运行，供 `switchApp` 判断
- **超时机制**：每个命令排队时都会设置 `NodeJS.Timeout`，超时 reject

### 2.6 WPS 加载项（Mac / Linux）`wps-claude-assistant/`

- **位置**：[wps-claude-assistant/main.js](file:///workspace/wps-claude-assistant/main.js)
- **入口函数**：`OnAddinLoad(ribbonUI)` — WPS 启动加载项时调用
- **核心循环**：
  - `startPolling()` → `poll()` → `scheduleNextPoll()`（`setTimeout` 递归）
  - `handleCommand(cmd)`：巨型 switch，覆盖 **Excel / Word / PPT / 通用** 所有 action（300+ 个）
- **关键实现细节**：
  - `getAppType()`：通过 `Application.Name` + `ActiveWorkbook/ActiveDocument/ActivePresentation` 探测当前 WPS 应用
  - `colToLetter(col)`：数字列号转 A1 字母列号（Mac WPS 的 `Cells()` 不如 `Range("A1")` 稳定，所以内部统一转成 A1 格式）
  - **Mac 批量赋值限制**：`range.Value = data[][]` 在 Mac WPS 不可用 → `handleSetRangeData()` 手动循环单元格赋值
  - **PPT 布局映射**：`layoutMap = { title: 1, title_content: 2, blank: 12, two_column: 3, comparison: 4 }`
  - **配色方案**：`COLOR_SCHEMES = { business, tech, creative, minimal }`（用于 `autoBeautifySlide` 等）
- **安装要求**（来自 `INSTALL.md`）：
  - 目录名必须以 `_` 结尾（如 `claude-assistant_`）—— WPS 的规范
  - 需被 `publish.xml` 引用才能被加载识别

### 2.7 工具层（Excel / Word / PPT / Common）

**入口**：[wps-office-mcp/src/tools/index.ts](file:///workspace/wps-office-mcp/src/tools/index.ts)
导出聚合 `RegisteredTool[]`，由 `toolRegistry.registerAll()` 一次性注册。

#### 2.7.1 Excel 工具（82+）

| 子模块 | 位置 | 代表工具 |
| --- | --- | --- |
| **formula.ts** | [formula.ts](file:///workspace/wps-office-mcp/src/tools/excel/formula.ts) | `set_formula`, `generate_formula`, `diagnose_formula`, `set_array_formula`, `recalculate`, `auto_sum` |
| **data.ts** | [data.ts](file:///workspace/wps-office-mcp/src/tools/excel/data.ts) | `read_range`, `write_range`, `clean_data`, `remove_duplicates`, `sort_range`, `find_replace`, `insert_row`, `add_comment`, `protect_sheet`, `set_conditional_format`, `hide_column` |
| **pivot.ts** | [pivot.ts](file:///workspace/wps-office-mcp/src/tools/excel/pivot.ts) | `create_pivot_table`, `update_pivot_table` |
| **chart.ts** | [chart.ts](file:///workspace/wps-office-mcp/src/tools/excel/chart.ts) | `create_chart`, `update_chart`, `export_chart_as_image`, `export_range_as_image` |
| **sheet.ts** | [sheet.ts](file:///workspace/wps-office-mcp/src/tools/excel/sheet.ts) | `create_sheet`, `delete_sheet`, `rename_sheet`, `copy_sheet`, `get_sheet_list`, `switch_sheet`, `move_sheet`, `get_selection`, `delete_row`, `insert_column`, `delete_column`, `freeze_panes`, `auto_fill`, `set_named_range`, `set_zoom` |
| **format.ts** | [format.ts](file:///workspace/wps-office-mcp/src/tools/excel/format.ts) | `set_cell_format`, `set_cell_style`, `set_border`, `set_number_format`, `merge_cells`, `unmerge_cells`, `set_column_width`, `set_row_height`, `auto_fit_column`, `auto_fit_row`, `set_data_validation`, `hide_row` |
| **workbook.ts** | [workbook.ts](file:///workspace/wps-office-mcp/src/tools/excel/workbook.ts) | `open_workbook`, `get_open_workbooks`, `switch_workbook`, `close_workbook`, `create_workbook`, `get_cell_value`, `set_cell_value`, `get_formula`, `get_cell_info`, `clear_range` |
| **data-advanced.ts** | [data-advanced.ts](file:///workspace/wps-office-mcp/src/tools/excel/data-advanced.ts) | `auto_filter`, `copy_range`, `paste_range`, `fill_series`, `transpose`, `text_to_columns`, `subtotal` |
| **row-column.ts** | [row-column.ts](file:///workspace/wps-office-mcp/src/tools/excel/row-column.ts) | `insert_rows`, `insert_columns`, `delete_rows`, `delete_columns`, `hide_rows`, `show_rows`, `show_columns`, `group_rows` |
| **comment-protect.ts** | [comment-protect.ts](file:///workspace/wps-office-mcp/src/tools/excel/comment-protect.ts) | `delete_cell_comment`, `get_cell_comments`, `unprotect_sheet`, `lock_cells`, `set_array_formula`, `insert_excel_image`, `set_hyperlink` |

**典型 Tool 实现结构**（以 `setFormula` 为例）：

```typescript
const setFormulaDefinition: ToolDefinition = {
  name: 'wps_excel_set_formula',
  description: '在指定单元格设置Excel公式。公式必须以=开头。',
  category: ToolCategory.SPREADSHEET,
  inputSchema: {
    type: 'object',
    properties: {
      range:   { type: 'string',  description: '目标单元格地址，如 A1、B2:B10' },
      formula: { type: 'string',  description: '公式（必须以=开头）' },
      sheet:   { type: 'string',  description: '工作表名（可选）' },
    },
    required: ['range', 'formula'],
  },
};

const setFormulaHandler: ToolHandler = async (args) => {
  // 1. 业务参数校验（如 formula 必须以 = 开头）
  // 2. 调用 wpsClient.executeMethod('setFormula', params, WpsAppType.SPREADSHEET)
  // 3. 包装成 ToolCallResult 返回
  return { id: uuidv4(), success: true, content: [{ type: 'text', text: '...' }] };
};
```

#### 2.7.2 Word 工具（32+）

| 子模块 | 代表工具 |
| --- | --- |
| **format.ts** | `apply_style`, `set_font`, `generate_toc`, `insert_bookmark`, `set_page_setup` |
| **content.ts** | `insert_text`, `find_replace`, `insert_table`, `set_paragraph`, `get_active_document`, `insert_image`, `insert_page_break`, `set_font_style`, `insert_comment`, `set_text_color`, `get_paragraphs`, `find_in_document`, `smart_fill_field`, `replace_bookmark_content` |
| **document.ts** | `get_open_documents`, `switch_document`, `open_document`, `get_document_text`, `insert_header`, `insert_footer`, `generate_doc_toc`, `insert_section_break`, `set_line_spacing` |
| **proofread.ts** | `enable_track_changes`, `get_track_changes_status`, `replace_range`, `proofread_basic` |

#### 2.7.3 PPT 工具（112+）

PPT 是工具数量最多、功能最丰富的模块，覆盖：

| 子模块 | 能力 |
| --- | --- |
| **slide.ts** | 基础幻灯片增删改（`add_slide`、`delete_slide`、`duplicate_slide`、`get_slide_count`、`get_slide_info`、`switch_slide`、`set_slide_layout`、`get/set_slide_notes`） |
| **slide-ops.ts** | 高级幻灯片操作（`add_shape`、`set_shape_style`、`add_textbox`、`set_slide_title`、`insert_image`、`set_shape_text`、`set_animation`、`set_background`、`set_slide_size`、`set_transition`、`add_chart`、`set_shape_fill`、`add_speaker_notes`） |
| **shape-basic.ts** | 形状基础操作（定位、样式、阴影、边框、透明度、圆角、层级 Z-order） |
| **textbox.ts** | 文本框 CRUD + 样式编排（`delete_textbox`, `get_textboxes`, `set_textbox_style` 等） |
| **image.ts** | 图片插入/删除/样式（`insert_image`, `delete_image`, `set_image_style`, `export_slide_as_image`） |
| **chart-flow.ts** | 图表/流程图/组织架构图/时间线（`create_flow_chart`, `create_org_chart`, `create_timeline`） |
| **data-viz.ts** | **数据可视化组件**（进度条 `create_progress_bar`、仪表盘 `create_gauge`、迷你图 `create_mini_charts`、环形图 `create_donut_chart`） |
| **animation.ts** | 动画/切换/超链接（`add_animation`, `set_animation_order`, `set_transition`, `add_ppt_hyperlink`） |
| **background.ts** | 背景/主题/配色（`set_slide_background`, `set_background_gradient`, `apply_color_scheme`） |
| **beautify-advanced.ts** | **智能美化**（`auto_beautify_slide`, `beautify_all_slides`, `unify_font`, `smart_distribute`, `create_kpi_cards`, `create_styled_table`, `add_title_decoration`, `add_page_indicator`） |
| **presentation.ts** | 文稿生命周期（`create_presentation`, `open_presentation`, `close_presentation`, `switch_presentation`, `get_open_presentations`, `set_slide_theme`, `copy_slide`, `insert_slide_image`） |
| **misc.ts** | 3D 效果、放映、查找替换（`set_3d_rotation`, `set_3d_depth`, `set_3d_material`, `create_3d_text`, `start_slide_show`, `end_slide_show`, `find_ppt_text`, `replace_ppt_text`） |

#### 2.7.4 Common 工具（9+）

| 子模块 | 工具 |
| --- | --- |
| **convert.ts** | `convert_to_pdf`、`convert_format`、`get_app_type_by_extension`（根据扩展名识别 appType）、`get_format_code` |
| **general.ts** | 通用工具，连接 `wpsClient.executeMethod()` 的业务封装 |

### 2.8 类型定义 `src/types/`

- **tools.ts**：
  - `ToolParameterSchema` — MCP 参数 Schema（`type: string|number|boolean|object|array`）
  - `ToolInputSchema` — `{ type: 'object', properties, required? }`
  - `ToolDefinition` — `{ name, description, inputSchema, category? }`
  - `ToolCategory` — 枚举（DOCUMENT / SPREADSHEET / PRESENTATION / COMMON）
  - `ToolHandler = (args) => Promise<ToolCallResult>`
  - `ToolCallResult = { id, success, content, error? }`
  - `ToolContent = { type: 'text'|'image'|'resource', text?, data?, mimeType? }`
  - `RegisteredTool = { definition, handler }`
- **wps.ts**：
  - `WpsAppType`（SPREADSHEET / DOCUMENT / PRESENTATION）
  - `WpsApiRequest = { method: string, params?: Record<string, unknown> }`
  - `WpsApiResponse<T> = { success: boolean, data?: T, error?: string }`
  - `WpsClientStatus = { connected, lastHeartbeat?, error? }`
  - `DocumentInfo / WorkbookInfo / PresentationInfo`

### 2.9 日志与错误

- **logger.ts**：Winston 驱动
  - 日志位置：`~/.wps-office-mcp/logs/{combined,error}.log`
  - **注意**：默认不输出到 Console — 因为 MCP 通过 stdio 传输，Console 输出会污染协议流；显式设置环境变量 `MCP_CONSOLE_LOG=true` 才会开启 Console 输出
  - `createChildLogger(moduleName)` 生成带模块名的子 logger
  - `log.info/warn/error/debug` 是快捷封装
  - `logRequest / logResponse` 用于 WPS API 级别的请求响应日志
- **error.ts**：
  - `McpError` — 基类（`new McpError('message')`）
  - `ToolNotFoundError(name)` — 工具不存在
  - `ToolExecutionError(name, originalError)` — 工具执行失败
  - `InvalidParamsError(msg, context?)` — 参数错误
  - `errorUtils.wrap(error, context)` — 将任意 Error / 未知值包装成 McpError，统一处理链

---

## 3. 关键设计决策与扩展点

### 3.1 设计决策 1：反转轮询架构（macOS 核心方案）

**问题**：WPS Mac 加载项不能启动 HTTP 服务器，无法被外部主动调用。
**方案**：MCP Server 启动 HTTP 服务（`:58891`），WPS 加载项每 500ms `GET /poll` 拉取命令，执行后 `POST /result` 回传。
**影响**：
  - 跨应用状态通过 `WpsMcpServer.dataCache` 传递
  - 需要应用自动切换（`wps-auto.sh`）
  - 必须有超时机制，避免命令挂起

### 3.2 设计决策 2：统一注册表（ToolRegistry）

**问题**：工具数量 247+，需要可扩展的注册机制而不是在代码里写 if-else。
**方案**：`RegisterTool` 装饰器 + `registerAll([...])` 批量注册。每个子模块导出 `RegisteredTool[]`，入口 `tools/index.ts` 展开合并。
**扩展方式**：新增一个 Tool 只需 3 步：
  1. 在对应子模块里新增 `{ definition, handler }`
  2. 在子模块 `index.ts` 的导出数组里加入
  3. 在 `wps-claude-assistant/main.js` 的 `handleCommand` switch 中补一个 `case 'yourAction': result = handle...(params); break`（如果是新 action）

### 3.3 设计决策 3：严格 TypeScript（`tsconfig.json`）

- `strict: true` + `strictNullChecks: true` + `noImplicitAny: true`
- `noUnusedLocals / noUnusedParameters / noImplicitReturns`
- 目标 `ES2022`：匹配 Node 18 的 API 版本
- `declaration: true`：产出 `.d.ts` 供其他工程类型导入

### 3.4 设计决策 4：MCP stdio 传输

- Claude Desktop 通过 `claude mcp add wps-office node /abs/path/to/wps-office-mcp/dist/index.js` 注册
- Server 启动后通过 `process.stdin` 读取 MCP JSON-RPC，`process.stdout` 写回结果
- **重要**：任何 `console.log` 都会破坏协议 → 全程使用 `logger` 且默认禁用 Console 输出

### 3.5 扩展点

1. **新 Tool**：遵循 2.4 的 3 步流程
2. **跨应用数据**：通过 `wps_cache_data` / `wps_get_cached_data` 桥接任意应用边界
3. **新平台**：新增一种 `execXxx` 实现，在 `wps-client.ts` 的平台路由里加入分支
4. **新 Skill**：在 `skills/xxx/SKILL.md` 里写"AI 如何调用该工具"的提示词规范，然后在 Claude Code 的 Skills 目录做软链接/复制

---

## 4. 依赖关系（关键 npm 包）

| 包 | 作用 | 版本 |
| --- | --- | --- |
| `@modelcontextprotocol/sdk` | MCP Server / Client SDK（核心） | ^1.29.0 |
| `axios` | HTTP 客户端（保活、外部 API） | ^1.6.0 |
| `uuid` | 生成 Tool 调用 ID / 请求 ID | ^9.0.0 |
| `winston` | 结构化日志 | ^3.11.0 |
| `typescript` / `ts-node` / `ts-jest` | 开发/编译/测试链 | ^5.3.0 / ^10.9.0 / ^29.1.0 |
| `jest` | 单元测试 & 覆盖率 | ^29.7.0 |
| `rimraf` | `npm run clean` 清理 dist | ^5.0.0 |
| `@types/node` | Node 类型声明（`ES2022` lib） | ^20.10.0 |

---

## 5. 构建与运行方式

### 5.1 常用命令（项目根目录）

| 命令 | 说明 |
| --- | --- |
| `npm run build` | `cd wps-office-mcp && tsc` → 生成 `dist/` |
| `npm start` | `node dist/index.js`（需先 build） |
| `npm run install:all` | 在子工程里 `npm install` |
| `npm test` | 运行 Jest 测试套件 |
| `npm run clean` | 删除 `wps-office-mcp/dist` |

### 5.2 开发模式

```bash
cd wps-office-mcp
npm install
npm run dev    # ts-node src/index.ts
```

### 5.3 MCP Server 注册（一次性）

**macOS / Linux**：
```bash
claude mcp add wps-office node /abs/path/to/wps-office-mcp/dist/index.js
```

**Windows**：
将以下 JSON 写入 `%USERPROFILE%\.claude\settings.json` 的 `mcpServers` 段：
```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["C:\\abs\\path\\to\\wps-office-mcp\\dist\\index.js"]
    }
  }
}
```

### 5.4 Skills 注册（提示词层）

- **macOS / Linux**：软链接到 `~/.claude/skills/wps-excel` 等
- **Windows**：复制 `skills/wps-excel` 到 `%USERPROFILE%\.claude\skills\wps-excel`

### 5.5 WPS 加载项安装

- **macOS**：`~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/`（目录必须 `_` 结尾）
- **Linux**：`~/.local/share/Kingsoft/wps/jsaddons/claude-assistant_/`
- **Windows**：`%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_\`
- 同步更新对应目录的 `publish.xml`：
  ```xml
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <jsplugins>
    <jsplugin name="claude-assistant" type="wps,et,wpp" url="claude-assistant_/" enable="enable_dev"/>
  </jsplugins>
  ```

### 5.6 完整启动顺序（Linux / macOS 尤其重要）

1. 启动 **Claude Code**（它会注册 MCP Server，Claude Code 自身 spawn 该进程）
2. 启动 **WPS Office**（加载项随之加载，开始轮询 `http://127.0.0.1:58891`）
3. 在 Claude Code 中用自然语言表达需求（如"打开销售数据.xlsx，生成每月销售额图表，再做一份 5 页的 PPT 摘要"）

### 5.7 调试建议

- 日志目录：`~/.wps-office-mcp/logs/`
- 查看：`tail -f ~/.wps-office-mcp/logs/combined.log`
- 手动探测：`curl http://127.0.0.1:58891/status`（返回 `{ status: 'running', currentApp, hasPendingCommand }`）
- 设置 `DEBUG=true` 以获取更详细的 MCP Server 日志
- 若 AI 反馈工具调用失败，检查：
  1. WPS 是否加载并在前台显示"Claude助手"Ribbon
  2. 轮询端口 58891 是否被占用
  3. 工具 action 是否在 `COMMAND_APP_MAP` 里注册了正确的 appType

---

## 6. 常见流程（从 AI 到 WPS）

以"在 Excel 中设置公式并生成图表"为例：

```
[Claude Code 用户输入]
    "帮我在 Sheet1 的 C1 写 =SUM(A1:B1)，然后生成 A1:C10 的柱状图"
         │
         ▼
[Claude AI] 解析需求后调用 MCP tools/call
         │
         ▼
[wps-office-mcp/src/server/mcp-server.ts]
  CallToolRequestSchema
         │
         ▼
[ToolRegistry.callTool(name='wps_excel_set_formula', args={ range:'C1', formula:'=SUM(A1:B1)', sheet:'Sheet1' })]
         │
         ▼
[setFormulaHandler] → wpsClient.executeMethod('setFormula', params, SPREADSHEET)
         │
         ▼
[WpsClient.invokeAction('setFormula', params)] → execMacPoll / execPowerShell
         │
         ▼
[macPollServer.executeCommand] → 等待 WPS 加载项下一次 GET /poll 时拿到命令
         │
         ▼
[wps-claude-assistant/main.js]
  handleCommand({ action:'setFormula', params })
    → handleSetFormula(params)
       → Application.ActiveWorkbook.Sheets('Sheet1').Range('C1').Formula = '=SUM(A1:B1)'
         │
         ▼
  sendResult(requestId, { success:true, data:{ cell:'$C$1', calculatedValue:xxx } })
         │
         ▼
[macPollServer.handleResult] → resolve promise
         │
         ▼
[ToolRegistry.callTool] → return ToolCallResult(success:true, content: [...])
         │
         ▼
[MCP Server] → 写回 stdout（JSON-RPC）
         │
         ▼
[Claude AI] 收到成功后，继续调用 wps_excel_create_chart 生成图表
         │
         ▼
[用户在 WPS 中看到]：C1 出现公式结果，图表插入到工作表
```

---

## 7. 关键文件索引（速查）

| 功能 | 文件 |
| --- | --- |
| 项目入口 & npm scripts | [package.json](file:///workspace/package.json) |
| MCP Server npm 配置 | [wps-office-mcp/package.json](file:///workspace/wps-office-mcp/package.json) |
| TypeScript 编译配置 | [wps-office-mcp/tsconfig.json](file:///workspace/wps-office-mcp/tsconfig.json) |
| 跨平台安装指南 | [INSTALL.md](file:///workspace/INSTALL.md) |
| MCP Server 启动入口 | [wps-office-mcp/src/index.ts](file:///workspace/wps-office-mcp/src/index.ts) |
| MCP Server 核心 | [mcp-server.ts](file:///workspace/wps-office-mcp/src/server/mcp-server.ts) |
| Tool 注册/调度 | [tool-registry.ts](file:///workspace/wps-office-mcp/src/server/tool-registry.ts) |
| 跨平台 WPS Client | [wps-client.ts](file:///workspace/wps-office-mcp/src/client/wps-client.ts) |
| macOS 轮询服务器 | [mac-poll-server.ts](file:///workspace/wps-office-mcp/src/client/mac-poll-server.ts) |
| Windows 保活 | [wps-keepalive.ts](file:///workspace/wps-office-mcp/src/client/wps-keepalive.ts) |
| Tools 聚合入口 | [tools/index.ts](file:///workspace/wps-office-mcp/src/tools/index.ts) |
| Excel 工具入口 | [excel/index.ts](file:///workspace/wps-office-mcp/src/tools/excel/index.ts) |
| Word 工具入口 | [word/index.ts](file:///workspace/wps-office-mcp/src/tools/word/index.ts) |
| Mac 加载项主逻辑 | [wps-claude-assistant/main.js](file:///workspace/wps-claude-assistant/main.js) |
| 日志 | [logger.ts](file:///workspace/wps-office-mcp/src/utils/logger.ts) |
| 错误 | [utils/error.ts](file:///workspace/wps-office-mcp/src/utils/error.ts) |
| Tool 类型 | [types/tools.ts](file:///workspace/wps-office-mcp/src/types/tools.ts) |
| WPS API 类型 | [types/wps.ts](file:///workspace/wps-office-mcp/src/types/wps.ts) |

---

> **维护提示**：本 Wiki 与源码同步更新。当以下文件发生结构性变化时请更新本 Wiki：`src/server/*.ts`、`src/client/*.ts`、`src/tools/**/index.ts`、`wps-claude-assistant/main.js`、`package.json`、`tsconfig.json`。

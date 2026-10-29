> 📦 **已归档（Archived — 仅作参考）**
> 本仓库已收入归档仓 **dev-toolkit**（代码与完整 git 历史一并迁入，原仓冻结不再开发）。
> 历史分支见归档仓的 `archive/pomodmer` 分支；活跃维护请前往 **dev-toolkit**。
> _This repo is archived into dev-toolkit; development continues there. Do not commit here._

# Pomodmer —— Cordis 内核 + 本地 Web UI 的番茄钟

仓库：<https://gitee.com/siwpan/pomodmer>（公开，MIT）

> 由 **cordis-table-app v1.0.0** 改造（2026-09-29 回滚到 v1.0.0 后转向番茄钟）。
> 表格能力（Python + OpenPyXL）保留为**可选插件**，启动时自动探测启用。
> 规划与执行记录见 [`docs/POMODORO-PLAN.md`](./docs/POMODORO-PLAN.md)；
> **完整能力说明（API 全表 / 数据模型 / 24 个表格方法 / 已知限制）见 [`docs/CAPABILITIES.md`](./docs/CAPABILITIES.md)**。

一个跑在本机的番茄钟：**计时权威在 Node 服务端，浏览器只负责渲染**。
刷新页面、切到后台、关掉标签页都不会让计时出错；进程被强杀重启后也不会"凭空补时"。

## 许可证

MIT License（见 [LICENSE](./LICENSE)），版权归 `siwpan` 所有。

## 快速开始

```bash
npm install            # 只装 cordis / loader / logger-console（已装可跳过）
npm run pom            # 启动番茄钟，自动打开 http://127.0.0.1:5173/
npm run smoke          # 冒烟测试（内核 + 番茄钟端到端，零外部依赖）
npm run smoke:ui       # 真实浏览器界面验收（用系统 Edge/Chrome + CDP，零新增依赖；没浏览器自动跳过）
npm run check:xlsx     # 跨实现验收：用 LibreOffice 打开导出的 xlsx 核对图表/隐藏表/真日期（没装则跳过）
npm run verify         # 冒烟 + 界面验收
npm run verify:full    # 再加 LibreOffice 验收（Node / 浏览器 / LibreOffice 三重实现）
npm run bench:export   # 导出基线：3000 条会话的耗时/内存实测
```

桌面壳（可选，需要先 `npm install electron`）：

```bash
npm run electron       # Electron 桌面版：托盘 / 置顶 / Ctrl+Alt+Space 全局快捷键 / 原生通知
```

Windows 双击即用：`pomodmer.bat`（依赖缺失会自动 `npm install`）。

**xlsx 表格能力自动启用**：启动时探测本机 Python + openpyxl，探到就把表格/报表插件一并挂上
（`cordis.with-xlsx.json`），探不到就安静退回纯番茄钟，两种机器都不用改配置。
日志里会打印这次选了哪个清单以及原因。要强制启用/关闭：

```bash
POMODMER_CONFIG=cordis.with-xlsx.json npm run pom   # 强制启用（缺 Python 时报表会明确报错）
POMODMER_NO_XLSX=1 npm run pom                      # 主动关闭（不想让 Python 被拉起时）
POMODMER_TABLE_TIMEOUT_MS=60000 npm run pom         # 表格调用超时（默认 15s，慢机器/超大工作簿可调大）
node src/pom.mjs --config cordis.table.json         # 只挂表格服务，不挂番茄钟
```

启动参数：

```bash
node src/pom.mjs --port 8080        # 换端口（或环境变量 POMODMER_PORT）
node src/pom.mjs --no-open          # 不自动开浏览器（或 POMODMER_OPEN=0）
```

## 架构

```
浏览器（零框架，web/）                       只渲染：按 endAt 插值画进度环与倒计时
   │  POST /api/...（控制）      ▲  SSE /api/events（每秒 tick + 阶段事件）
   ▼                             │
Cordis 宿主（Node.js 22，零新增 npm 依赖）
  ├─ ctx.store    src/plugins/store.mjs     JSON 原子落盘（data/store.json）
  ├─ ctx.timer    src/plugins/timer.mjs     番茄钟状态机 + 权威计时
  └─ ctx.http     src/plugins/http.mjs      node:http 静态托管 + REST + SSE
  （可选）ctx.table  src/plugins/table.mjs ──JSON-RPC stdio──▶ python/rpc_server.py（OpenPyXL）
  （可选）ctx.report src/plugins/report.mjs  周报 xlsx（依赖 ctx.table）
  （可选）Electron 壳  electron/main.cjs ──spawn──▶ 上面的宿主，BrowserWindow 加载同一个 web/
```

- `cordis.json` 是插件清单（声明式入口）；`cordis.table.json` 只挂表格插件。
- 启动时按「显式 `--config` > `POMODMER_CONFIG` > 自动探测」选清单：探到 Python + openpyxl 就用
  `cordis.with-xlsx.json`（番茄钟 + 表格 + 报表），否则用 `cordis.json`。决策与原因每次都打进日志。
- 三个服务的加载顺序由 `inject` 依赖图推导：`store → timer → http`，无需手动编排。

## 复用既有表格能力（Python + OpenPyXL）

这套能力继承自 `cordis-table-app` v1.0.0，它不是"周报专用代码"，而是一个**通用表格服务**：
`ctx.table` 把表格操作经 JSON-RPC（stdio 行分隔）转发给常驻的 Python worker，宿主自己完全不碰 Excel。

| 分类 | 方法 | 说明 |
|---|---|---|
| 工作簿 | `createWorkbook(path, sheet)` | 新建工作簿并指定首个工作表名 |
| 工作表 | `addSheet(path, name)` | 追加工作表 |
| 写入 | `writeRows(path, sheet, startRow, startCol, rows)` | 批量写二维数组（空值写 `null` 即空单元格） |
| 写入 | `writeRows(..., { colTypes })` | 按列声明类型 `'datetime'`/`'date'`/`'number'`/`'text'`：把 ISO 串还原成 Excel **原生**类型，日期才能排序/筛选/做日期计算 |
| 写入 | `writeCell(path, sheet, row, col, value)` | 单格写 |
| 写入 | `setCellText(path, sheet, row, col, text)` | 只改文本（保留样式与合并） |
| 样式 | `styleHeader(path, sheet, row, ncols, opts)` | 表头底色/加粗/居中/边框（`opts.fill` / `opts.font_color`） |
| 样式 | `setCellFill(path, sheet, row, col, color)` | 单元格底色（ARGB 十六进制） |
| 样式 | `setBorders(path, sheet, r1, c1, r2, c2, opts)` | 区域边框 |
| 样式 | `setComment(path, sheet, row, col, text)` | 批注 |
| 结构 | `mergeCells(path, sheet, r1, c1, r2, c2)` | 合并单元格区域 |
| 样式 | `setColumnWidth(path, sheet, col, width)` | 列宽 |
| 样式 | `setNumberFormat(path, sheet, items)` | 批量数字格式：`items: [{ range: 'A2:A8', format: 'yyyy-mm-dd' }]` |
| 样式 | `freezePanes(path, sheet, cell)` | 冻结窗格（`'A2'` = 表头常驻） |
| 读取 | `readSheet(path, sheet)` | 读回 `{ rows, max_row, max_col }`（日期以 ISO 串返回） |
| 读取 | `sheetMeta(path, sheet, row)` | 冻结窗格 / 数字格式 / 值类型 / 合并区 / 批注 —— 用来验证导出文件的**真实形态** |
| 落盘 | `save(path)` / `close(path)` | 保存（`close` 顺带移出 worker 的内存缓存） |
| 自检 | `ping()` | 返回 Python / openpyxl 版本 |

> 同一个 worker 里还有 `readMap` / `renderMapPng` / `renderToken` / `genDungeon` /
> `exportDonjon` / `importDonjon`（权威源的地图与 VTT 能力），pomodmer 用不到，
> 但它们随三件套一起来 —— 需要时直接可用，不必再开一份 worker。

三种用法，同一份 worker 支撑：

1. **UI / HTTP 出报表**（四条链路，分工明确）：
   - `GET /api/export?format=xlsx&scope=weekly&days=7` → `report.weekly({ days })`：**周报**，最近 N 天，
     四张表：概览 / 每日 / 会话 / 任务。UI 的「导出周报 xlsx」按钮走这条，直接下载。
   - `scope=month[&month=YYYY-MM]` → `report.month()`：**月报**，整月 + 一张「对比」表（本月 vs 上月的
     完成番茄 / 专注时长 / 活跃天数 / 作废 / 计划外事件 / 会话条数，含变化与变化率）。
   - `scope=range&from=&to=` → `report.range()`：**自定义区间**（含首尾两天），UI 里给两个日期框，
     空了先自动填最近 7 天、再点一次才导出。
   - `scope=all` → `report.full()`：**全量导出**，从第一条会话到今天，五张表（多一张「计划外事件」）。
   - 每条都以附件直接下载，响应头带 `X-Pomodmer-Info` 摘要（scope / days / 表数 / 会话数 / 完成番茄）；
     需要**磁盘路径**而不是浏览器下载时（脚本、定时任务），仍可用 `POST /api/report`。
   - 导出的 xlsx 里还有两张**隐藏表**：`_meta`（时区偏移等上下文）与 `_epochs`（逐行绝对毫秒）。
     它们是"跨时区回灌仍逐毫秒精确"的依据（见下方回灌说明）。
   - 「每日」表右侧（F 列起）内嵌两张**原生图表**：柱状图（每日完成番茄）+ 折线图（每日专注时长），
     引用的是同一张表的数据区。CSV/TSV 没有图表概念（纯数据出口）。
   - **回灌**：`POST /api/import/xlsx?mode=preview|apply` → `report.importXlsx(file, { mode })`。
     导出 → 在 Excel 里改 → 导回，是这条能力的完整闭环。语义是**去重追加**：只新增 + 有限更新
     （任务的预估/完成状态），**从不删除**已有记录；先 `preview` 拿到"新增/更新/跳过"的计划，确认后再 `apply`。
     四个不变量在冒烟里都有断言：预览不写库；同一份表重复导入全部跳过（幂等）；
     **原样导出的文件再导回一条都不新增**（xlsx 只有秒精度，按毫秒比对会把历史翻倍）；
     **跨时区导入逐毫秒等于导出时的绝对时刻**（靠隐藏表 `_meta`/`_epochs`；被编辑过的行按文件记录的时区偏移换算，精度到秒）。
   - 分工：`GET /api/export`（JSON）用来**回灌备份**（能被 `/api/import` 再吃回去），
     xlsx 三条链路用来**给人看、二次分析、改完再导回**（真实日期、数字格式、冻结表头，
     可以直接排序/筛选/透视）。
2. **命令行：任意 JSON → xlsx**（与番茄钟无关的通用小工具，换任何项目都能直接用）：

   ```bash
   npm run gen -- --json sample.json --out output/out.xlsx --sheet Data --header
   npm run demo        # 端到端演示：扫描 Cordis 包清单生成 xlsx 并回读校验
   ```

3. **写自己的插件复用 `ctx.table`**（不要绕开它直接调 Python）：

   ```js
   import { Service } from 'cordis'
   export class MyExport extends Service {
     static provide = 'myexport'
     static inject = ['store', 'timer']      // 只把**必需**依赖写进 inject
     async dump() {
       const t = this.ctx.get('table')        // 可选依赖：用 ctx.get，不要写进 inject
       if (!t) throw new Error('未启用表格能力：装好 Python + openpyxl 后重启即自动启用')
       const file = 'output/dump.xlsx'
       await t.createWorkbook(file, '数据')
       await t.writeRows(file, '数据', 1, 1, [['列A', '列B'], [1, 2]])
       await t.styleHeader(file, '数据', 1, 2)
       await t.save(file)
       return { file }
     }
   }
   ```

   > 可选依赖**写进 `inject` 会让插件永远停在 PENDING**（README 尾部「设计要点」有实测记录），
   > 所以允许降级的能力一律 `ctx.get('table')` 运行时取。

**移植到别的项目**：`src/plugins/table.mjs` + `src/core/bridge.mjs` + `python/rpc_server.py` 三件套，
清单里加一条 `{ "id": "table", "name": "./src/plugins/table.mjs" }` 即可。
解释器不硬编码任何盘符（项目 `.venv` → `py -3` → PATH → 用户级/系统级/各盘符常见落点，可用 `config.pythonBin` / `PYTHON_BIN` 覆盖），换机换系统都不用改代码。

### 这三件套不在本仓库维护（单一权威源 + 来源锁）

权威源是 [cordis-table-app](https://gitee.com/siwpan/cordis-table-app)（表格 + 地图/VTT 能力的超集）。
pomodmer 是从那边**同步过来的消费者**：

```bash
npm run sync:kit     # 从权威源同步三件套，并更新 tools/table-kit.lock.json
npm run check:kit    # 体检：锁 ↔ 本地（权威源在场时再比对源），不一致即报错
```

为什么不让两边各存一份：这套代码曾经在两个仓库里各自演化，结果同一个
**"日期单元格杀死整个 worker"** 的缺陷被两边各自踩了一遍、各修了一遍，UTF-8 那个坑同样是两边各修一次。
复制粘贴治不了这个病，能治的是"单一权威源 + 可验证的来源锁"：

- `tools/table-kit.json` 声明同步哪些文件、权威源在哪；
- `tools/table-kit.lock.json` 记录**来源仓库 + 提交号 + 每个文件的 sha256**；
- 哈希按**行尾归一**后的内容计算：`core.autocrlf` 在不同机器上把同一提交检出成 CRLF / LF 不会被误判成漂移，
  但真正改内容（多一行、改一字）立刻报红；
- 手改这三件套却不走同步 → `npm run check:kit` 和冒烟第 13 节立刻报红；同步工具默认拒绝覆盖未登记的本地改动。

三条纪律（都是踩过的坑）：

- **stdio 只走协议**：stdout 只允许一行一个 JSON-RPC 对象，被调库的任何 `print()` 都会污染通道 ——
  worker 用 `contextlib.redirect_stdout(sys.stderr)` 兜住，别去掉。
- **两端必须同为 UTF-8**：Windows 中文环境里 Python 默认按 gbk/cp936 解 stdio，UTF-8 中文会先变乱码、
  再退化成孤立代理字符，一路潜伏到 openpyxl 落盘才以 `surrogates not allowed` 炸出来。
  桥侧注入 `PYTHONIOENCODING=utf-8`，worker 侧 `_force_utf8_stdio()` 兜底，两道都要。
- **读回来的值必须能过 JSON，且响应序列化要在 try 内**：Excel 里的日期是 `datetime`、高精度数是
  `Decimal`、二进制是 `bytes`，`json.dumps` 一个都不认。worker 用 `_json_default` 统一转成
  ISO 串/数字/base64，并且把序列化也放进 `try` —— 否则**一个日期单元格就能让 `json.dumps` 抛错、
  掀翻主循环带走整个 worker**，此后所有调用都变成 `python worker exited`（这个坑真实发生过，别改回去）。
- **工作簿按 `path` 缓存在 Python 内存**：写完必须 `save`（或 `close`）才落盘；同一 path 复用同一份工作簿。

## 番茄工作法规则（落地实现）

| 规则 | 实现 |
|---|---|
| 专注 25 分钟 / 短休 5 分钟 / 每 4 个番茄长休 15–30 分钟 | `settings.focusMs / shortMs / longMs / longEvery`，均可在 UI 设置面板改 |
| 番茄不可分割，提前结束不算完成 | `skip` 记录 `completed=false, reason='skipped'`；`round` 不增加 |
| 中断即作废 | `void` 记录 `reason='void'`，当前番茄归零 |
| 系统休眠/进程挂起不静默补时 | tick 空档 > 10s 判定为挂起 → 当前番茄按作废处理（`reason='suspend'`） |
| 退出期间到点不算完成 | 重启恢复时记录 `reason='shutdown'` 且 `completed=false` |

## 计时为什么准

1. **不累加 tick 次数**：进度永远现算 `remaining = endAt - now`，`endAt` 是绝对时刻。
2. **不用 `setInterval`**：递归 `setTimeout` 且对齐到整秒边界（目标时间固定），不累积漂移。
3. **时钟分工**：`Date.now()`（墙钟）负责开始/结束时刻与落盘；
   `performance.now()`（单调钟）只用于调度耗时观测。
4. **前端不推进时间**：SSE 每秒下发 `endAt` 与服务端 `now`，浏览器按本地插值渲染。

## 功能（P0 + P1）

**P0**：番茄钟核心 —— 专注/短休/长休自动流转、开始/暂停/继续/重置/跳过/作废、
可配置时长与长休间隔、自动开休息、提示音与滴答声、桌面通知、亮/暗主题。

**P1**：任务清单（番茄归属，实际/预估对比）、统计面板（今日/连续天数/最近 7 天柱状图/
12 周热力图/每任务进度条）、中断记录（内部/外部 + 计划外事件列表）、数据导出/导入 JSON。

**P2**：xlsx 导出四条链路 —— 周报（最近 N 天）/ **月报**（整月 + 「本月 vs 上月」对比）/
**自定义区间**（任意 from~to）/ **全量导出**（从第一条会话到今天）；日期是 Excel 原生日期、带数字格式、
表头冻结、**「每日」表内嵌柱状图与折线图**（openpyxl 原生图表对象，趋势一眼可见）；
**xlsx 回灌**（导出 → 在 Excel 里改 → 导回，先预览再确认，去重追加不删记录，
且**跨时区导入逐毫秒精确**）；**CSV/TSV 出口**（纯 JS，没装 Python 也能导原始数据）；
白噪音与棕噪音（Web Audio 实时合成）、
迷你窗（独立小窗只显示倒计时）、Windows 原生气泡通知（best-effort）、`pomodmer.bat` 自愈式启动器。

**P3（可选）**：Electron 桌面壳 —— 托盘常驻、窗口置顶、`Ctrl+Alt+Space` 全局快捷键、
原生阶段通知；装不了 Electron 也不影响浏览器模式。

## HTTP API

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/state` | 全量快照（settings / tasks / runtime / 今日统计） |
| GET | `/api/events` | SSE：`hello` 快照 + 每秒 `tick` + `phase` / `session` / `settings` / `interrupt` / `imported` 事件 |
| POST | `/api/timer/start` `pause` `resume` `toggle` `reset` `skip` `void` | 控制 |
| PATCH | `/api/settings` | 改时长与开关（`focusMs` / `theme` 等，非法值被拒） |
| GET / POST | `/api/tasks` | 任务列表 / 新增 |
| PATCH / DELETE | `/api/tasks/:id` | 改任务 / 删任务 |
| PATCH | `/api/tasks/current` | 设当前任务（番茄完成后自动计到它头上） |
| GET | `/api/sessions?limit=100` | 历史会话 |
| GET | `/api/stats?days=7&weeks=12` | 统计：today / streak / recent / heat / byTask / total |
| GET / POST | `/api/unplanned` | 计划外事件列表 / 记录（`kind`: internal\|external） |
| DELETE | `/api/unplanned/:id` | 删除一条中断记录 |
| GET / POST | `/api/report` | 周报能力探测 / 生成周报 xlsx（未启用时 GET 给提示、POST 返回 501） |
| GET | `/api/export` | 导出全部数据（JSON 附件；**用来回灌备份**，可被 `/api/import` 再导入） |
| GET | `/api/export?format=csv\|tsv&sheet=会话` | **CSV / TSV 出口（纯 JS，不需要 Python）**：`sheet=` 概览/每日/会话/任务/计划外事件；可选 `scope=`（同 xlsx）、`bom=0`（给 pandas/grep 用）；默认带 UTF-8 BOM 便于 Excel 双击打开 |
| GET | `/api/export?format=xlsx` | xlsx 导出（附件下载）。`scope=all`（默认，5 张可见表全量）/ `weekly&days=7`（周报 4 张）/ `month[&month=YYYY-MM]`（月报 6 张，含「本月 vs 上月」对比）/ `range&from=&to=`（自定义区间）；入参非法 → 400；响应头 `X-Pomodmer-Info` 带回摘要 |
| POST | `/api/import` | 导入数据（校验失败返回 400；导入后停在暂停态） |
| POST | `/api/import/xlsx?mode=preview\|apply` | 从 xlsx **回灌**：`preview` 只算不写（先看会新增/更新/跳过多少），`apply` 才落库；去重追加，**永不删已有记录** |

> 路由顺序坑：`/api/tasks/current` 必须排在 `/api/tasks/:id` 之前，否则会被当成 `id='current'` 的任务。

> 上面是速查表；每个接口的字段、响应头与错误语义，以及**已知限制清单**（例如
> `bridge.call()` 目前没有超时）见 [`docs/CAPABILITIES.md`](./docs/CAPABILITIES.md)。

## 数据

`data/store.json`（已 gitignore，属个人专注记录）：

```jsonc
{
  "settings": { "focusMs": 1500000, "shortMs": 300000, "longMs": 900000, "longEvery": 4, ... },
  "tasks":    [ { "id", "title", "estimate", "spent", "done", "createdAt" } ],
  "sessions": [ { "id", "type", "taskId", "plannedMs", "startAt", "endAt", "actualMs", "completed", "interrupts", "reason" } ],
  "unplanned":[ { "id", "kind": "internal|external", "note", "at", "phase", "taskId" } ],
  "runtime":  { "phase", "running", "endAt", "remainingMs", "startedAt", "plannedMs", "round", "currentTaskId", "interrupts" }
}
```

写入是**原子**的（写 `.tmp` 再 `rename`）；数据文件损坏时自动备份为 `.broken-<ts>` 并回退默认配置。

## 目录

```
pomodmer/
├─ LICENSE / README.md / docs/
│  ├─ CAPABILITIES.md       完整能力说明（API 全表 / 数据模型 / 表格 24 方法 / 已知限制）
│  └─ POMODORO-PLAN.md      规划、决策与执行记录（含下一阶段候选 P5–P8）
├─ cordis.json              插件清单（默认：store + timer + http）
├─ cordis.table.json        插件清单（可选：仅 table，xlsx 能力）
├─ cordis.with-xlsx.json    插件清单（可选：番茄钟 + table + report，启用周报）
├─ tools/                   表格三件套同步（sync-table-kit.mjs / table-kit.json / table-kit.lock.json）
│                           + ui-smoke.mjs（真实浏览器界面验收）+ bench-export.mjs（导出基线）
├─ pomodmer.bat             自愈式启动器（Windows 双击即用）
├─ electron/                main.cjs 主进程 / start.mjs 启动器 / make-icon.mjs + icon.png
├─ web/                     index.html / app.js / style.css（零框架前端）
├─ python/rpc_server.py     OpenPyXL 工作者（可选能力用）
└─ src/
   ├─ pom.mjs               应用入口
   ├─ boot.mjs              宿主启动/关停（入口/CLI/测试共用）
   ├─ core/bridge.mjs       JsonRpcBridge（stdio 行分隔 JSON-RPC）
   ├─ plugins/store.mjs     StoreService（ctx.store）
   ├─ plugins/timer.mjs     TimerService（ctx.timer）
   ├─ plugins/http.mjs      HttpService（ctx.http）
   ├─ plugins/table.mjs     TableService（ctx.table，可选；由权威源同步，勿手改）
   ├─ demo.mjs / cli.mjs    xlsx 演示与命令行工具（走 cordis.table.json）
   ├─ _smoke.mjs            Cordis 内核冒烟
   └─ _smoke_pom.mjs        番茄钟端到端冒烟（真实 HTTP）
```

## 设计要点（Cordis 4.0.0-rc.10 实测）

- `ctx.plugin()` / `loader.create()` 的效果是**异步执行**的，注册后必须 `await`，服务才会挂上。
- 服务的初始化用 `[Service.init]()`，返回清理函数（或 Promise of 清理函数），
  fiber dispose 时自动执行 —— 定时器、HTTP server、SSE 连接都在这里回收。
- Cordis **没有 `ctx.stop()`**；根 fiber 的 dispose 语义是"重启"。正确关停 =
  逆序 `await entry.fiber.dispose()` 顶层入口（`boot.mjs` 的 `stopHost` 已封装）。
- 服务间只通过 `ctx.<key>` 查找，不 import 具体实现；依赖用 `static inject = [...]` 声明。
  **可选依赖不要写进 `inject`**（否则插件永远 PENDING），用 `ctx.get('table')` 在用的时候取。

## Electron 踩坑（实测）

- 环境里若有 `ELECTRON_RUN_AS_NODE=1`，`electron.exe` 会退化成纯 Node：
  `require('electron')` 返回的是**二进制路径字符串**，症状是 `app.whenReady()` 报 undefined。
  该变量在启动前就被读取，main 进程里改来不及 → 用 `electron/start.mjs` 在 spawn 前删掉它。
- Electron 主进程**不支持** `import { app } from 'electron'` 的命名导出（44.4.5 实测），
  本项目 `package.json` 是 `type: module`，所以主进程写成 `main.cjs`（CommonJS）。
- 国内安装必须配镜像：`ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/`（约 100MB）。

## 路线图

**功能路线（P0–P3，全部完成）**：

- **P0（已完成，tag v2.0.0）**：Cordis 宿主 + store/timer/http 三插件 + 本地 Web UI + 端到端冒烟。
- **P1（已完成，tag v2.1.0）**：任务清单 + 番茄归属、统计与热力图、中断记录、亮/暗主题、快捷键、导入导出。
- **P2（已完成，tag v2.2.0）**：周报 xlsx（复用 ctx.table）、白噪音/棕噪音、迷你窗、原生通知、`pomodmer.bat` 启动器。
- **P3（已完成，tag v2.2.0）**：Electron 桌面壳（托盘 / 置顶 / `Ctrl+Alt+Space` / 原生通知），仅 devDependencies，装不了也不影响浏览器模式。

**表格能力专项（P1–P4，全部完成）**：xlsx 数据可用性（真实日期/数字格式/冻结表头 + 修 worker 崩溃）→
三件套统一到权威源（来源锁 + 同步工具）→ 全量导出 → xlsx 回灌（预览 + 去重追加）→ 导出面做齐。
逐步记录见 [`docs/POMODORO-PLAN.md`](./docs/POMODORO-PLAN.md)。

**下一阶段候选（P5–P8，已立项未开工）**：压力与边界加固（worker 超时与重启、大数据基线、时区）→
月度/自定义区间报表 → CSV/TSV 出口（不依赖 Python）→ 报表内嵌图表（需改权威源并同步）。
排序依据、改动面与验收要点见规划文档 §9。

详见 [`docs/POMODORO-PLAN.md`](./docs/POMODORO-PLAN.md)。

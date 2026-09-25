# Cordis 表格小应用（Cordis 内核 + OpenPyXL）

基于 **Cordis（Node.js 插件框架）作为应用内核**，
表格能力交给 **Python + OpenPyXL**，两者通过 **JSON-RPC stdio 桥** 连接。

> 即 `cordis-host-python-bridge` 模式：Cordis 宿主负责编排/插件化，Python 进程负责实际
> 的 Excel 读写；宿主不碰 Excel，Python 不碰编排。

## 许可证

本项目以 **MIT License** 开源（见 [LICENSE](./LICENSE)），版权归 `siwpan` 所有。
可自由使用、修改、分发，仅需保留版权与许可声明。

## 架构

```
Cordis 宿主 (Node.js)
  ├─ Context                      src/boot.mjs       进程根上下文
  ├─ @cordisjs/plugin-loader      声明式插件宿主（读 cordis.json 的 entries）
  ├─ ConsoleExporter              日志导出器（否则 loader 失败静默）
  └─ TableService extends Service 挂到 ctx.table     src/plugins/table.mjs
        └─ JsonRpcBridge (stdio)  src/core/bridge.mjs
              └─ python/rpc_server.py   OpenPyXL 工作者（零依赖 stdlib 协议层）
```

- `cordis.json` 是插件清单（唯一条目入口）：`{ "entries": [ { id, name, config } ] }`，
  其中 `name` 是相对项目根（baseUrl）的插件模块路径。
- 通过 loader 加载的插件 = `TableService`（构造器插件），它 `provide('table')`，对外暴露
  `ctx.table` 的一系列表格方法（createWorkbook / writeRows / styleHeader / readSheet / save ...）。
- 这些方法经 `JsonRpcBridge` 以"一行一 JSON-RPC"转发给 Python worker，由 OpenPyXL 执行。

## 源码来源

Cordis 内核由 `npm install` 安装（`cordis` / `@cordisjs/plugin-loader` / `@cordisjs/plugin-logger-console`），
与 `package.json` 中声明的版本一致。`vendor/cordis-main/` 仅用于在本机保留一份框架源码镜像，
**已被 `.gitignore` 排除，不随仓库分发**（第三方内核源码不入仓，见仓库卫生门禁）。

## 运行

```bash
# 进入本项目根目录（任意路径均可，无盘符硬编码）
npm install            # 装 cordis / loader / logger-console（已装可跳过）

# 演示：生成 output/cordis-packages.xlsx（含两张表）
npm start

# 通用生成：JSON -> xlsx
npm run gen -- --json sample.json --out output/sample.xlsx --sheet Data

# 跑团地图：手写地图 -> 渲染 VTT PNG（端到端演示）
npm run map:demo
# 跑团地图：程序化生成地牢 -> 渲染 PNG
npm run map:gen
# 跑团地图：命令行生成/渲染（从参数）
npm run map:cli -- gen --cols 30 --rows 20 --rooms 6 --level 3 --png output/map.png
```

> **Python 解释器自动发现**（无需硬编码路径，换机即用）：
> 1. `cordis.json` 里 table 条目的 `config.pythonBin`；
> 2. 环境变量 `PYTHON_BIN`；
> 3. 扫描系统 `PATH` + 常见安装目录（Windows `C:/Python3XX`、`%LOCALAPPDATA%/Programs/Python/...`，
>    macOS/Linux `/usr/bin/python3`、`/opt/homebrew/bin/python3` 等），挑第一个能 `import openpyxl` 的；
> 4. 兜底回退到 `PATH` 里的 `python3` / `python`。
>
> 推荐直接 `pip install openpyxl`（≥3.1），然后让程序自动发现即可。

## 跑团地图制作（地图编辑 + VTT 渲染 + 程序化生成）

把表格内核延伸到跑团地图：**xlsx 是可编辑「源真相」，PNG 是「VTT 导出物」**。

- **Map 表**：每单元格 = 1 格（5 英尺）。填充色 = 地形，边框 = 墙，单元格文本 = 房间标签，批注 = DM 备注（遭遇/陷阱/战利品）。
- **数据表**：`Rooms`(id,x,y,w,h,name,desc,encounter,traps,treasure)、`Encounters`、`Loot`——地图与战前准备一起生成（Dungeon Flow 式）。
- **渲染**：`renderMapPng` 把网格光栅化为 PNG，默认 **100px/格（Foundry VTT）**；`--px 70` 即得 Roll20 规格；可选烘焙网格。另可 `renderToken` 导出 400×400 透明占位 token。

地形调色板（ARGB，可改）：

| 名称   | 色值       | 含义         |
|--------|------------|--------------|
| STONE  | `FF3A3A3A` | 石墙 / 室外   |
| FLOOR  | `FFF2E2C0` | 地板         |
| CORR   | `FFD9C9A0` | 走廊         |
| DOOR   | `FF8B5A2B` | 门           |
| WATER  | `FF4F81BD` | 水池         |

> 网格基准：D&D 5e 中 1 格 = 5 英尺。VTT 像素标准：Foundry 默认 100px/格（最小 50，常见 70/100/140/200），Roll20 固定 70px，Fantasy Grounds 50px。渲染尺寸必须 = 列×px × 行×px（**整数格、图边与网格齐平**），Foundry 才能对齐。

可用方法（`ctx.table`，均经 JSON-RPC 转发给 Python worker）：`setCellFill / mergeCells / setCellText / setComment / setBorders / readMap / renderMapPng / renderToken / genDungeon`。

程序化生成 `genDungeon({cols,rows,rooms,party_level,seed})`：随机铺房间+走廊+门+水池，并填好 `Rooms/Encounters/Loot` 数据表；房间键/遭遇/陷阱/战利品均为**泛型文本**（系统无关 + 轻量 5e CR），不引用任何官方资料，规避版权。

> 注意：openpyxl 的 `merge_cells` 会清除非锚点单元格的填充，故带独立填充（如房间内的水池）的区域**不要整块合并**——房间名写在左上角单元格即可。

## 设计要点（与 Cordis rc.10 实测一致）

- `ctx.plugin()` / `loader.create()` 的效果是**异步执行**的，注册后必须 `await` 一次，服务才会挂上。
- Cordis **没有 `ctx.stop()`**；根 fiber 的 dispose 语义是"重启"。正确关停 = 逆序 `await entry.fiber.dispose()`
  顶层入口，触发各服务的清理函数（此处 = 杀掉 Python worker），再兜底强杀。
- Python worker 的 stdout **只走 JSON-RPC**，所有日志/异常栈走 stderr，避免污染协议通道。
- 每个 RPC 方法 try/except 兜住，返回 JSON-RPC error，保证子进程不死。

## 目录

```
cordis-table-app/
├─ LICENSE                   MIT 许可证
├─ cordis.json              插件清单（声明式入口）
├─ python/rpc_server.py     OpenPyXL 工作者（JSON-RPC over stdio）
├─ src/
│  ├─ boot.mjs              宿主启动/关停（入口共用）
│  ├─ core/bridge.mjs       JsonRpcBridge（stdio 行分隔 JSON-RPC）
│  ├─ plugins/table.mjs     TableService（挂 ctx.table，含 Python 自发现）
│  ├─ demo.mjs              端到端演示（包清单表 + 演示表）
│  ├─ cli.mjs               通用 JSON -> xlsx 工具
│  ├─ map_demo.mjs          跑团地图手写演示（地图 + 渲染）
│  ├─ map_gen.mjs           程序化地牢生成演示
│  ├─ map_cli.mjs           地图生成 / 渲染命令行
│  └─ _smoke.mjs            冒烟测试（Cordis 内核可用性）
├─ vendor/cordis-main/      Cordis 内核源码镜像（可选，gitignore 排除）
└─ output/                 生成的 xlsx 产物（gitignore 排除）
```

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
# 跑团地图：导出 DonJon 风格 JSON
npm run map:export -- --xlsx output/map.xlsx --out output/map.donjon.json --name my-dungeon
# 跑团地图：导入 DonJon 风格 JSON -> 新 xlsx（可顺带渲染 PNG）
npm run map:import -- --json output/map.donjon.json --xlsx output/imported.xlsx --png output/imported.png
# 跑团地图：DonJon 互通往返校验（生成->导出->导入->比对 600 格）
npm run map:dj
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

地形调色板（ARGB，可改；前 5 种为 v1.1.0 原有，其余为 v1.2.0 新增）：

| 名称    | 色值       | 含义                          |
|---------|------------|-------------------------------|
| STONE   | `FF3A3A3A` | 石墙 / 未开凿                  |
| FLOOR   | `FFF2E2C0` | 房间地板                      |
| CORR    | `FFD9C9A0` | 走廊                          |
| DOOR    | `FF8B5A2B` | 门                            |
| WATER   | `FF4F81BD` | 水池                          |
| LAVA    | `FFB3471F` | 熔岩（危害地形）              |
| GRASS   | `FF7CA64B` | 苔藓 / 草地面（房间变体）      |
| SAND    | `FFE6D98A` | 沙地（房间变体）              |
| WOOD    | `FF8B6B3A` | 木桥 / 木板（房间变体）        |
| RUBBLE  | `FF8A8A8A` | 碎石地面（房间变体）          |
| STAIR   | `FF9ACD32` | 楼梯（上 / 下）               |
| SECRET  | `FF6A5ACD` | 密门                          |
| TRAP    | `FFB22222` | 陷阱标记（手动地图用）        |

> 房间生成时会在房间内随机铺 1~2 格 `grass/sand/wood/rubble` 变异地面、随机 1~2 处楼梯、以及 1 处密门，增加视觉层次。water/lava/grass/sand/wood/rubble/stair/trap 在 DonJon 无对应位，导出时用 `_cta.tiles` 精确还原（见下节）。

> 网格基准：D&D 5e 中 1 格 = 5 英尺。VTT 像素标准：Foundry 默认 100px/格（最小 50，常见 70/100/140/200），Roll20 固定 70px，Fantasy Grounds 50px。渲染尺寸必须 = 列×px × 行×px（**整数格、图边与网格齐平**），Foundry 才能对齐。

可用方法（`ctx.table`，均经 JSON-RPC 转发给 Python worker）：`setCellFill / mergeCells / setCellText / setComment / setBorders / readMap / renderMapPng / renderToken / genDungeon / exportDonjon / importDonjon`。

程序化生成 `genDungeon({cols,rows,rooms,party_level,seed})`：随机铺房间+走廊+门+水池+楼梯+密门，房间内随机变异地面（草/沙/木/碎石），并填好 `Rooms/Encounters/Loot` 数据表；房间按 **22 种主题模板**（大厅/守卫室/祭坛/宝库/王座厅/图书馆/锻炉/陵墓…）生成，含偏置的遭遇/陷阱/战利品；房间键/遭遇/陷阱/战利品均为**泛型文本**（系统无关 + 轻量 5e CR），不引用任何官方资料，规避版权。

> 注意：openpyxl 的 `merge_cells` 会清除非锚点单元格的填充，故带独立填充（如房间内的水池）的区域**不要整块合并**——房间名写在左上角单元格即可。

### DonJon 风格 JSON 互通（导入 / 导出）

让生成的地图能在 **其他 VTT 工具间互通**：可导出为 DonJon（donjon.bin.sh）风格的 JSON，也能导入一份 DonJon 风格 JSON 还原为可编辑的 xlsx 地图。

- **导出** `exportDonjon(path, out, {sheet,name,px,seed,level})`：把 Map 表 + 数据表写成 DonJon 风格 JSON。
- **导入** `importDonjon(path, jsonPath, {sheet})`：把 DonJon 风格 JSON 还原为 Map 表 + `Rooms/Encounters/Loot` 数据表 + 房间标签 / 批注 / 外墙。

JSON 结构（对齐 donjon.bin.sh 真实导出）：

| 字段     | 含义                                                                                  |
|----------|---------------------------------------------------------------------------------------|
| `name` / `creator` | 地图名 / 生成器标识                                                          |
| `settings` | `{name, cell_size, n_cols, n_rows, seed, level}`（`cell_size` 默认 100px/格）        |
| `cells`  | `cells[行][列]` 整数矩阵；每位是 DonJon cell 标志位                                              |
| `rooms`  | `[{id, col, row, width, height, name, contents{summary, detail{monster[],trap[],hidden_treasure[],room_features}}, doors{north/south/east/west[]}}]` |
| `stairs` / `egress` | 楼梯 / 外围出口（本工具生成时通常为空）                                           |
| `_cta`   | **cordis-table-app 私有扩展**（其他工具忽略未知键）：`palette`(地形→ARGB) + `terrain_to_donjon`(地形→DonJon 位) + `tiles`([r,c,key] 精确覆盖) |

DonJon cell 标志位（来自 donjon.bin.sh `dungeon.pl` 源码）：`0x01` 墙 / `0x02` 房间 / `0x04` 走廊 / `0x10` 外墙 / `0x20000` 门 / `0x100000` 密门 / `0x400000` 下楼梯 / `0x800000` 上楼梯。

> **保真策略**：`water / lava / grass / sand / wood / rubble / stair / trap` 这些地形在 DonJon 没有对应位，导出时归为"房间(0x02)"，并用 `_cta.tiles` 记录它们在矩阵中的精确位置；导入时先按 DonJon 标准位解码（`stone/door/corr/floor/secret/stair` 精确还原，`water` 等退化为 floor），再用 `_cta.tiles` 覆盖回精确颜色。**我们工具间往返 100% 无损**；其他 DonJon 工具导入时只忽略 `_cta`，把 `water` 等当作普通地板——这是合理降级，不破坏任何字段。

> **互通验证**：`npm run map:dj` 跑"生成 → 导出 → 导入 → 比对"，校验 600 格填充与房间数完全往返一致；另有"无 `_cta` 的外来 DonJon JSON"导入用例，确认标准位解码不崩、不丢房间。

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
│  ├─ map_cli.mjs           地图生成 / 渲染 / 导出 / 导入命令行
│  ├─ map_dj.mjs            DonJon 风格 JSON 互通往返校验
│  └─ _smoke.mjs            冒烟测试（Cordis 内核可用性）
├─ vendor/cordis-main/      Cordis 内核源码镜像（可选，gitignore 排除）
└─ output/                 生成的 xlsx 产物（gitignore 排除）
```

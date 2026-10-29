// TableService —— Cordis 服务，挂到 ctx.table
// 它是"构造器插件"：被 Cordis 实例化后通过 Service 基类提供 ctx.table；
// 内部 spawn 一个 Python OpenPyXL worker（JsonRpcBridge），所有表格操作经 JSON-RPC 转发。
import { Service } from 'cordis'
import { JsonRpcBridge, pythonCommandParts } from '../core/bridge.mjs'
import { fileURLToPath } from 'node:url'
import { dirname, resolve, sep } from 'node:path'
import { existsSync } from 'node:fs'
import { execSync } from 'node:child_process'

const __dirname = dirname(fileURLToPath(import.meta.url))
const SCRIPT = resolve(__dirname, '../../python/rpc_server.py')

// 自发现 Python 解释器（全程不硬编码任何本机盘符）：
//   1) cordis.json 里 table 条目的 config.pythonBin（显式优先）
//   2) 环境变量 PYTHON_BIN
//   3) `py` 官方启动器 / PATH 里的 python（最常见、也最符合系统约定）
//   4) **遍历所有就绪盘符**探测常见安装落点（Python3XX 根目录 + 用户级安装目录）
// 换机 / 换系统 / 装到任意盘都能直接用，无需改代码。
//
// 2026-09-25 修正：原实现把候选表写死成 C:/D: 上的几个标准落点，注释却声称
// "不硬编码任何本机盘符"——**声称与实现对不上**：装到 F: 盘或用户级目录就一个都不命中，
// 表现为"明明装了 Python 却找不到"。死根路径门禁报出 6 处即此因。
// 现改为按盘符枚举 + 用户级目录探测，不再出现任何写死的绝对路径。
const POSIX_COMMON = [
  '/usr/bin/python3',
  '/usr/local/bin/python3',
  '/opt/homebrew/bin/python3',
  '/usr/bin/python',
]
const PY_VERSIONS_MAC = ['314', '313', '312', '311']

function pathCandidates() {
  const out = []
  const pathEnv = process.env.PATH || ''
  const sepCh = process.platform === 'win32' ? ';' : ':'
  const names = process.platform === 'win32'
    ? ['python.exe', 'python3.exe', 'pythonw.exe']
    : ['python3', 'python']
  for (const dir of pathEnv.split(sepCh)) {
    if (!dir) continue
    for (const n of names) {
      const p = resolve(dir, n)
      try { if (existsSync(p)) out.push(p) } catch { /* ignore */ }
    }
  }
  return out
}

// 枚举 Windows 上"当前就绪"的盘符，用于在各盘根目录探测 Python 安装。
// 不写死任何盘符：只探测**实际存在**的盘，装机盘符变了也能命中。
function driveLetters() {
  if (process.platform !== 'win32') return []
  const out = []
  for (let c = 65; c <= 90; c++) {          // A..Z
    const d = String.fromCharCode(c)
    try {
      // 探测根目录是否可访问（不存在的盘会抛错）
      if (existsSync(d + ':/')) out.push(d)
    } catch { /* ignore */ }
  }
  return out
}

// 项目级落点：放在系统级之前。理由：项目自带解释器（.venv）最能保证依赖齐备；
// 而系统 PATH 里可能有一个装了 openpyxl 但版本/环境有问题的解释器（本机实测踩中）。
function projectCandidates() {
  if (process.platform !== 'win32') return []
  return [
    resolve(__dirname, '../../.venv/Scripts/python.exe'),
    resolve(__dirname, '../../venv/Scripts/python.exe'),
  ]
}

function commonCandidates() {
  if (process.platform !== 'win32') {
    const extra = []
    for (const v of PY_VERSIONS_MAC) {
      extra.push('/Library/Frameworks/Python.framework/Versions/3.' + v.slice(1) + '/bin/python3')
    }
    return [...POSIX_COMMON, ...extra]
  }
  const out = []
  const localApp = process.env.LOCALAPPDATA || ''
  const progFiles = [process.env.ProgramFiles, process.env['ProgramFiles(x86)']].filter(Boolean)
  // ⓪ 项目级虚拟环境已由 projectCandidates() 单独前插，此处不再重复
  // ① 用户级安装（官方安装器默认落点，与盘符无关）
  const localRoot = localApp ? resolve(localApp, 'Programs/Python') : ''
  // ② 系统级安装落点
  const roots = [localRoot, ...progFiles.map((p) => resolve(p, 'Python'))].filter(Boolean)
  for (const root of roots) {
    for (const v of ['314', '313', '312', '311', '310']) {
      out.push(resolve(root, 'Python' + v, 'python.exe'))
    }
    out.push(resolve(root, 'python.exe'))
  }
  // ③ 遍历所有就绪盘符的根目录（覆盖装在非系统盘的情况，如 /Python314）
  for (const drive of driveLetters()) {
    for (const v of ['314', '313', '312', '311', '310']) {
      out.push(drive + ':/Python' + v + '/python.exe')
    }
    out.push(drive + ':/Python/python.exe')
  }
  return out
}

function canImportOpenpyxl(bin, timeoutMs = 8000) {
  // bin 可能是"带参数的命令"（Windows 官方启动器 py -3）：整串加引号会被 cmd 当成
  // 一个不存在的程序名，必然探活失败（这正是这一条以前从未生效的原因）。
  // 统一走 pythonCommandParts 拆词：字面路径原样、命令行拆出程序与前置参数。
  const [cmd, ...preArgs] = pythonCommandParts(bin)
  const line = [cmd, ...preArgs].map((s) => '"' + s + '"').join(' ')
  try {
    execSync(line + ' -c "import openpyxl"', { stdio: 'ignore', timeout: timeoutMs })
    return true
  } catch {
    return false
  }
}

// 候选解释器（去重后保持优先级）：
//   ① 项目自带 .venv/venv —— 依赖最可控（本机实测：PATH 里的解释器可能版本/环境有问题）
//   ② py -3 官方启动器
//   ③ PATH 里的 python
//   ④ 用户级 / 系统级 / 各盘符下的常见安装落点
function candidateList(config) {
  const raw = []
  if (config && config.pythonBin) raw.push(config.pythonBin)
  if (process.env.PYTHON_BIN) raw.push(process.env.PYTHON_BIN)
  raw.push(...projectCandidates(), 'py -3', ...pathCandidates(), ...commonCandidates(), 'python3', 'python')
  const seen = new Set()
  return raw.filter((c) => {
    if (!c || seen.has(c)) return false
    seen.add(c)
    return true
  })
}

// 探测结果 memo：一个进程内环境不会变，而每次探测都要 spawn 解释器（本机约 0.3s）。
// 没有 memo 的话，"宿主启动时决定清单"和 TableService 构造会各探一遍，纯属白等。
let probeMemo = { key: null, result: null }

/**
 * 探测本机表格能力：返回 { bin, openpyxl, anyBin }
 *   - bin     能 `import openpyxl` 的解释器；null 表示没有
 *   - anyBin  第一个"看起来存在"的解释器；用来区分"根本没装 Python"与"装了但缺 openpyxl"
 *              （这两种情况给用户的提示完全不同，不能混为一谈）
 *
 * 显式指定（config.pythonBin / PYTHON_BIN）视为权威：只探它、不做跨解释器回退。
 * 否则用户明明指定了某个解释器，却被悄悄换成另一个，排查起来极其费解。
 *
 * opts.timeoutMs 单个候选的探测上限；opts.budgetMs 整轮探测总预算（0=不限）。
 * 宿主启动时的自动探测用较小预算：个别候选（如 Microsoft Store 的 python 占位程序）卡住时，
 * 不能让"锦上添花"的能力探测把启动拖成好几秒 —— 探不到就安静退回纯表格之外的形态。
 */
export function probeTableWorker(config, opts = {}) {
  const explicit = (config && config.pythonBin) || process.env.PYTHON_BIN || null
  const key = explicit || '(auto)'
  if (probeMemo.key === key && probeMemo.result && !opts.fresh) return probeMemo.result

  const timeoutMs = opts.timeoutMs || 8000
  const budgetMs = opts.budgetMs || 0
  let result

  if (explicit) {
    result = { bin: explicit, openpyxl: canImportOpenpyxl(explicit, timeoutMs), anyBin: explicit }
  } else {
    const startedAt = Date.now()
    let anyBin = null
    let bin = null
    for (const c of candidateList(config)) {
      if (budgetMs && Date.now() - startedAt > budgetMs) break
      const bare = !c.includes(sep)
      // 带参数的启动器（py -3）与裸命令名由操作系统解析，绝对路径才做存在性检查
      if (!bare && !existsSync(c)) continue
      // anyBin 只认"确认存在"的绝对路径：否则在根本没装 Python 的机器上，
      // 列表尾部的裸名 python3 会让提示误报成"装了 Python 但缺 openpyxl"，指错方向。
      if (!anyBin && !bare) anyBin = c
      if (canImportOpenpyxl(c, timeoutMs)) {
        bin = c
        break
      }
    }
    result = { bin, openpyxl: !!bin, anyBin }
  }

  probeMemo = { key, result }
  return result
}

// TableService 用的解释器解析：优先能 `import openpyxl` 的；一个都没有时回退到首个
// "看起来存在"的解释器或 'python3'，让"缺依赖"由 Python 自己报出来（错误信息更明确），
// 而不是在这里变成一句"找不到 Python"。
export function resolvePythonBin(config) {
  const p = probeTableWorker(config)
  return p.openpyxl ? p.bin : p.anyBin || 'python3'
}

export class TableService extends Service {
  static provide = 'table'

  constructor(ctx, config) {
    super(ctx, 'table')
    this.config = config || {}
    const pythonBin = resolvePythonBin(this.config)
    // 调用超时：配置项优先，其次环境变量，最后用桥的默认值（15s）。
    // 超大工作簿导出确实可能慢，慢机器可以调大：POMODMER_TABLE_TIMEOUT_MS=60000
    const timeoutMs = Number(this.config.timeoutMs) || Number(process.env.POMODMER_TABLE_TIMEOUT_MS) || undefined
    this.bridge = new JsonRpcBridge(pythonBin, SCRIPT, { timeoutMs })
  }

  // worker pid 是**派生状态**，不缓存：worker 会因为崩溃/超时被换掉，
  // 缓存一份必然过期（踩过：崩溃自愈后 pid 还指着死进程，日志与排查全被带偏）。
  get pid() {
    return this.bridge && this.bridge.child ? this.bridge.child.pid : null
  }

  // Service.init：返回清理函数（fiber dispose 时执行 -> 杀掉 python worker）
  [Service.init]() {
    const bridge = this.bridge
    bridge.start()
    bridge.onExit = (code) => {
      this.ctx?.logger?.debug?.('python worker exited (code ' + code + ')')
    }
    // 超时/写入失败会自动重启：日志里必须看得见，否则"偶发失败一次、重试就好了"永远查不出根因
    bridge.onRestart = (reason, count) => {
      this.ctx?.logger?.warn?.('table: python worker 重启（%s），累计 %d 次', reason, count)
    }
    return () => {
      bridge.stop()
    }
  }

  // ---- RPC 代理方法（与 python/rpc_server.py 的 method 一一对应）----
  ping() {
    return this.bridge.call('ping', {})
  }
  createWorkbook(path, sheet) {
    return this.bridge.call('create_workbook', { path, sheet })
  }
  // opts.hidden：建隐藏表（给机器读的元数据表用，例如导出时记下的时区偏移与逐行 epoch）
  addSheet(path, name, opts = {}) {
    return this.bridge.call('add_sheet', { path, name, ...opts })
  }
  // opts.colTypes：按列声明类型（'text' | 'number' | 'datetime' | 'date' | 'time'），
  // 让 worker 把 ISO 字符串还原成 Excel 原生类型 —— 日期要能排序/筛选/算日期就必须走这条。
  // JS 侧写 camelCase，过桥时统一翻成 worker 的 snake_case（协议那侧只有 snake_case）。
  writeRows(path, sheet, startRow, startCol, rows, opts = {}) {
    const { colTypes, ...rest } = opts
    return this.bridge.call('write_rows', {
      path, sheet, start_row: startRow, start_col: startCol, rows, col_types: colTypes, ...rest,
    })
  }
  writeCell(path, sheet, row, col, value) {
    return this.bridge.call('write_cell', { path, sheet, row, col, value })
  }
  styleHeader(path, sheet, row, ncols, opts = {}) {
    return this.bridge.call('style_header', { path, sheet, row, ncols, ...opts })
  }
  setColumnWidth(path, sheet, col, width) {
    return this.bridge.call('set_column_width', { path, sheet, col, width })
  }
  // items: [{ range: 'A2:A8', format: 'yyyy-mm-dd' }, ...]
  setNumberFormat(path, sheet, items) {
    return this.bridge.call('set_number_format', { path, sheet, items })
  }
  freezePanes(path, sheet, cell) {
    return this.bridge.call('freeze_panes', { path, sheet, cell })
  }
  // 图表：把趋势长在表里（openpyxl 原生图表对象，不是图片）。
  // spec 用显式行列号而不是 A1 字符串（解析带中文/空格的表名引用是经典坑），形如：
  //   { type: 'col'|'bar'|'line'|'pie', title, anchor: 'F2', data_sheet, height, width,
  //     cats: { col, min_row, max_row },
  //     series: [ { col, min_row, max_row, title_from_data: true } ] }
  addChart(path, sheet, spec = {}) {
    return this.bridge.call('add_chart', { path, sheet, ...spec })
  }
  // 读工作表元信息（冻结窗格 / 数字格式 / 值类型 / 图表）：给"验证导出文件真实形态"用
  sheetMeta(path, sheet, row) {
    return this.bridge.call('sheet_meta', { path, sheet, row })
  }
  readSheet(path, sheet) {
    return this.bridge.call('read_sheet', { path, sheet })
  }
  save(path) {
    return this.bridge.call('save', { path })
  }
  close(path) {
    return this.bridge.call('close', { path })
  }

  // ---- 地图相关方法（与 python/rpc_server.py 的 method 对应，共用同一 worker）----
  setCellFill(path, sheet, row, col, color) {
    return this.bridge.call('set_cell_fill', { path, sheet, row, col, color })
  }
  mergeCells(path, sheet, r1, c1, r2, c2) {
    return this.bridge.call('merge_cells', { path, sheet, r1, c1, r2, c2 })
  }
  setCellText(path, sheet, row, col, text) {
    return this.bridge.call('set_cell_text', { path, sheet, row, col, text })
  }
  setComment(path, sheet, row, col, text) {
    return this.bridge.call('set_comment', { path, sheet, row, col, text })
  }
  setBorders(path, sheet, r1, c1, r2, c2, opts = {}) {
    return this.bridge.call('set_borders', { path, sheet, r1, c1, r2, c2, ...opts })
  }
  readMap(path, sheet) {
    return this.bridge.call('read_map', { path, sheet })
  }
  renderMapPng(path, sheet, out, px = 100, grid = false) {
    return this.bridge.call('render_map_png', { path, sheet, out, px, grid })
  }
  renderToken(text, out, size = 400, color = 'FF4F81BD') {
    return this.bridge.call('render_token', { text, out, size, color })
  }
  genDungeon(path, opts = {}) {
    return this.bridge.call('gen_dungeon', { path, ...opts })
  }
  // ---- DonJon 风格 JSON 互通 ----
  // 导出：把 Map 表 + 数据表 -> DonJon 风格 JSON 文件（out）
  exportDonjon(path, out, opts = {}) {
    return this.bridge.call('export_donjon', { path, out, ...opts })
  }
  // 导入：把 DonJon 风格 JSON 文件（json）-> 新 Map 表 + 数据表
  importDonjon(path, jsonPath, opts = {}) {
    return this.bridge.call('import_donjon', { path, json: jsonPath, ...opts })
  }
}

export default TableService

// TableService —— Cordis 服务，挂到 ctx.table
// 它是"构造器插件"：被 Cordis 实例化后通过 Service 基类提供 ctx.table；
// 内部 spawn 一个 Python OpenPyXL worker（JsonRpcBridge），所有表格操作经 JSON-RPC 转发。
import { Service } from 'cordis'
import { JsonRpcBridge } from '../core/bridge.mjs'
import { fileURLToPath } from 'node:url'
import { dirname, resolve, sep } from 'node:path'
import { existsSync } from 'node:fs'
import { execSync } from 'node:child_process'

const __dirname = dirname(fileURLToPath(import.meta.url))
const SCRIPT = resolve(__dirname, '../../python/rpc_server.py')

// 自发现 Python 解释器（全程不硬编码任何本机盘符）：
//   1) cordis.json 里 table 条目的 config.pythonBin（显式优先）
//   2) 环境变量 PYTHON_BIN
//   3) 扫描 PATH + 常见安装目录，挑第一个能 `import openpyxl` 的
//   4) 兜底回退到 PATH 里的 python3 / python（让其自行报缺依赖，错误信息更明确）
// 换机 / 换系统 / 装到任意盘都能直接用，无需改代码。
const WINDOWS_COMMON = [
  'C:/Python314/python.exe',
  'C:/Python313/python.exe',
  'C:/Python312/python.exe',
  'C:/Python311/python.exe',
  'D:/Python314/python.exe',
  'D:/Python313/python.exe',
]
const POSIX_COMMON = [
  '/usr/bin/python3',
  '/usr/local/bin/python3',
  '/opt/homebrew/bin/python3',
  '/usr/bin/python',
]

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

function commonCandidates() {
  if (process.platform === 'win32') {
    const localApp = process.env.LOCALAPPDATA || ''
    const extra = []
    if (localApp) {
      // Microsoft Store / 官方安装器常见落点
      extra.push(resolve(localApp, 'Programs/Python/Python314/python.exe'))
      extra.push(resolve(localApp, 'Programs/Python/Python313/python.exe'))
      extra.push(resolve(localApp, 'Programs/Python/Python312/python.exe'))
    }
    return [...WINDOWS_COMMON, ...extra]
  }
  return POSIX_COMMON
}

function canImportOpenpyxl(bin) {
  try {
    execSync('"' + bin + '" -c "import openpyxl"', { stdio: 'ignore', timeout: 8000 })
    return true
  } catch {
    return false
  }
}

export function resolvePythonBin(config) {
  if (config && config.pythonBin) return config.pythonBin
  if (process.env.PYTHON_BIN) return process.env.PYTHON_BIN

  const candidates = [...pathCandidates(), ...commonCandidates(), 'python3', 'python']
  const seen = new Set()
  let firstExisting = null
  for (const c of candidates) {
    if (!c || seen.has(c)) continue
    seen.add(c)
    // 裸命令名（python3/python）由操作系统 PATH 解析，直接探活
    if (c.includes(sep) && !existsSync(c)) continue
    if (canImportOpenpyxl(c)) return c
    if (!firstExisting) firstExisting = c
  }
  // 没有任何一个带 openpyxl：回退到首个探活对象（让它去报缺依赖，错误信息更明确）
  return firstExisting || 'python3'
}

export class TableService extends Service {
  static provide = 'table'

  constructor(ctx, config) {
    super(ctx, 'table')
    this.config = config || {}
    const pythonBin = resolvePythonBin(this.config)
    this.bridge = new JsonRpcBridge(pythonBin, SCRIPT)
    this.pid = null
  }

  // Service.init：返回清理函数（fiber dispose 时执行 -> 杀掉 python worker）
  [Service.init]() {
    const bridge = this.bridge
    const ready = bridge.start()
    ready.then((pid) => {
      this.pid = pid
    })
    bridge.onExit = (code) => {
      this.ctx?.logger?.debug?.('python worker exited (code ' + code + ')')
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
  addSheet(path, name) {
    return this.bridge.call('add_sheet', { path, name })
  }
  writeRows(path, sheet, startRow, startCol, rows) {
    return this.bridge.call('write_rows', { path, sheet, start_row: startRow, start_col: startCol, rows })
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
}

export default TableService

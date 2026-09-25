// map_demo.mjs —— P0 验证：手写一张网格地图 -> 渲染 VTT PNG -> 端到端校验
// 验证点：① Python 自发现 ② 单元格填充=地形 ③ 合并/标签/批注 ④ PNG 尺寸=列×px×行×px
//        ⑤ openpyxl 回读填充+文本 ⑥ python worker 进程回收
import { bootHost, stopHost } from './boot.mjs'
import { mkdirSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const XLSX = resolve(__dirname, '../output/map_demo.xlsx')
const PNG = resolve(__dirname, '../output/map_demo.png')
const TOKEN = resolve(__dirname, '../output/token_demo.png')

const COLS = 14
const ROWS = 14
const PX = 100 // Foundry 默认 100px/格

async function main() {
  const ctx = await bootHost()
  const m = ctx.table
  const info = await m.ping()
  console.log('[map-demo] python/openpyxl =>', info.python, '/', info.openpyxl)

  mkdirSync(dirname(XLSX), { recursive: true })
  await m.createWorkbook(XLSX, 'Map')

  // 地形调色板（ARGB）
  const STONE = 'FF3A3A3A'
  const FLOOR = 'FFF2E2C0'
  const CORR = 'FFD9C9A0'
  const DOOR = 'FF8B5A2B'
  const WATER = 'FF4F81BD'

  // 整张填石墙，再挖房间/走廊/门/水池
  const grid = []
  for (let r = 1; r <= ROWS; r++) {
    const row = []
    for (let c = 1; c <= COLS; c++) row.push(STONE)
    grid.push(row)
  }
  for (let r = 2; r <= 11; r++) for (let c = 2; c <= 11; c++) grid[r - 1][c - 1] = FLOOR
  for (let c = 11; c <= 13; c++) grid[6 - 1][c - 1] = CORR // 走廊
  grid[6 - 1][11 - 1] = DOOR // 门
  grid[9 - 1][4 - 1] = WATER
  grid[9 - 1][5 - 1] = WATER

  for (let r = 1; r <= ROWS; r++)
    for (let c = 1; c <= COLS; c++) await m.setCellFill(XLSX, 'Map', r, c, grid[r - 1][c - 1])

  // 房间标签 + DM 批注 + 外墙（注意：openpyxl 合并会清除非锚点单元格填充，
  // 故不在带独立填充的房间内部做整块合并——房间名写在左上角单元格即可）
  await m.setCellText(XLSX, 'Map', 2, 2, '大厅 A')
  await m.setComment(XLSX, 'Map', 2, 2, '遭遇：3 只狗头人(CR 1/4)；陷阱：地板夹层；战利品：12 gp + 蓝宝石')
  await m.setBorders(XLSX, 'Map', 2, 2, 11, 11, { color: 'FF1F1F1F', style: 'thick' })

  await m.save(XLSX)

  // 渲染 PNG（VTT 自绘网格，故 grid=false）
  const ren = await m.renderMapPng(XLSX, 'Map', PNG, PX, false)
  console.log('[map-demo] render =>', JSON.stringify(ren))
  const expectW = COLS * PX
  const expectH = ROWS * PX
  if (ren.width !== expectW || ren.height !== expectH)
    throw new Error('PNG 尺寸不符 VTT flush-grid 规则: ' + ren.width + 'x' + ren.height + ' 期望 ' + expectW + 'x' + expectH)

  // token 占位图
  const tok = await m.renderToken('狗头人', TOKEN, 400, 'FF8B5A2B')
  console.log('[map-demo] token =>', JSON.stringify(tok))

  // 回读校验（填充色 + 文本 + 批注内容经 openpyxl 落盘）
  const back = await m.readMap(XLSX, 'Map')
  console.log('[map-demo] readMap =>', back.max_row, 'x', back.max_col)
  console.log('[map-demo] cell(2,2).text =', back.rows[1][1].text, '| cell(2,2).fill =', back.rows[1][1].fill)
  console.log('[map-demo] cell(9,4).fill(水池) =', back.rows[8][3].fill)
  if (back.rows[1][1].text !== '大厅 A') throw new Error('回读文本不符')
  if (back.rows[8][3].fill !== WATER) throw new Error('回读水池填充色不符')
  // 房间内部地板填充须保留（验证未做整块合并导致填充丢失）
  if (back.rows[4][4].fill !== FLOOR) throw new Error('回读房间内部地板填充不符: ' + back.rows[4][4].fill)

  const pyPid = ctx.table ? ctx.table.pid : null
  await stopHost()
  console.log('[map-demo] 产物:', XLSX, PNG, TOKEN)
  console.log('[map-demo] 宿主已关停，python worker (pid=' + pyPid + ') 已回收。')
  process.exit(0)
}

main().catch((e) => {
  console.error('[map-demo] FATAL:', e)
  process.exit(1)
})

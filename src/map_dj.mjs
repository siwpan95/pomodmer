// map_dj.mjs —— DonJon 风格 JSON 互通往返校验
// 流程：gen 地牢 -> 导出 DonJon JSON -> 导入为另一份 xlsx -> 校验单元格填充往返一致 + 房间数一致 + PNG 整数格
import { bootHost, stopHost } from './boot.mjs'
import { mkdirSync, readFileSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const XLSX_A = resolve(__dirname, '../output/dj_a.xlsx')   // 原始生成
const XLSX_B = resolve(__dirname, '../output/dj_b.xlsx')   // 导入还原
const JSONF = resolve(__dirname, '../output/dj.donjon.json')
const PNG_A = resolve(__dirname, '../output/dj_a.png')
const PNG_B = resolve(__dirname, '../output/dj_b.png')

const COLS = 30
const ROWS = 20
const PX = 100

function mapFills(rm) {
  // readMap 返回 rows: [{fill,text}]；压成 "row,col->fill" 便于比对
  const m = {}
  for (let r = 0; r < rm.rows.length; r++) {
    for (let c = 0; c < rm.rows[r].length; c++) {
      const cell = rm.rows[r][c]
      m[r + ',' + c] = cell.fill || 'NONE'
    }
  }
  return m
}

async function main() {
  const ctx = await bootHost()
  const m = ctx.table
  const info = await m.ping()
  console.log('[map-dj] python/openpyxl =>', info.python, '/', info.openpyxl)

  mkdirSync(dirname(XLSX_A), { recursive: true })

  // 1) 生成
  await m.createWorkbook(XLSX_A, 'Map')
  const gen = await m.genDungeon(XLSX_A, {
    sheet: 'Map', cols: COLS, rows: ROWS, rooms: 6, party_level: 3, seed: 777,
  })
  console.log('[map-dj] gen =>', JSON.stringify(gen))
  await m.save(XLSX_A)

  // 2) 导出 DonJon JSON
  const exp = await m.exportDonjon(XLSX_A, JSONF, { sheet: 'Map', name: 'dj-test', px: PX, seed: 777, level: 3 })
  console.log('[map-dj] export =>', JSON.stringify(exp))
  const dj = JSON.parse(readFileSync(JSONF, 'utf8'))
  if (!Array.isArray(dj.cells) || !dj.cells.length) throw new Error('导出 JSON 缺少 cells')
  if (!dj._cta || !dj._cta.palette) throw new Error('导出 JSON 缺少 _cta.palette')
  console.log('[map-dj] dj name=%s cols=%d rows=%d rooms=%d paletteKeys=%s',
    dj.name, dj.settings.n_cols, dj.settings.n_rows, dj.rooms.length, Object.keys(dj._cta.palette).join(','))

  // 3) 导入为另一份 xlsx
  const imp = await m.importDonjon(XLSX_B, JSONF, { sheet: 'Map' })
  console.log('[map-dj] import =>', JSON.stringify(imp))
  await m.save(XLSX_B)

  // 4) 渲染两套 PNG
  const renA = await m.renderMapPng(XLSX_A, 'Map', PNG_A, PX, false)
  const renB = await m.renderMapPng(XLSX_B, 'Map', PNG_B, PX, false)
  console.log('[map-dj] renderA', JSON.stringify(renA), '| renderB', JSON.stringify(renB))
  if (renA.width !== COLS * PX || renB.width !== COLS * PX) throw new Error('PNG 宽非整数格')
  if (renA.height !== ROWS * PX || renB.height !== ROWS * PX) throw new Error('PNG 高非整数格')

  // 5) 单元格填充往返一致性
  const rmA = await m.readMap(XLSX_A, 'Map')
  const rmB = await m.readMap(XLSX_B, 'Map')
  const fa = mapFills(rmA)
  const fb = mapFills(rmB)
  const keys = new Set([...Object.keys(fa), ...Object.keys(fb)])
  let mism = 0
  for (const k of keys) {
    if (fa[k] !== fb[k]) {
      mism++
      if (mism <= 5) console.log('  填充差异 @%s: A=%s B=%s', k, fa[k], fb[k])
    }
  }
  if (mism > 0) throw new Error('单元格填充往返不一致，差异格数=' + mism)
  console.log('[map-dj] 单元格填充往返一致 ✔ (比对 %d 格)', keys.size)

  // 6) 房间数一致
  const roomsA = await m.readSheet(XLSX_A, 'Rooms')
  const roomsB = await m.readSheet(XLSX_B, 'Rooms')
  const nA = roomsA.max_row - 1 // 去表头
  const nB = roomsB.max_row - 1
  console.log('[map-dj] Rooms 数据行: A=%d B=%d', nA, nB)
  if (nA !== nB) throw new Error('房间数往返不一致: ' + nA + ' vs ' + nB)
  if (nA < 1) throw new Error('未生成任何房间')

  await stopHost()
  console.log('[map-dj] 往返校验全部通过。产物:', XLSX_A, XLSX_B, JSONF, PNG_A, PNG_B)
  process.exit(0)
}

main().catch((e) => {
  console.error('[map-dj] FATAL:', e)
  process.exit(1)
})

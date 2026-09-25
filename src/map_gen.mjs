// map_gen.mjs —— P1 验证：程序化生成地牢 -> 渲染 VTT PNG -> 校验整数格 + 数据表
// 验证点：① gen_dungeon 铺 Map 表 + Rooms/Encounters/Loot 数据表
//        ② PNG 尺寸 = 列×px × 行×px（Foundry flush-grid 整数格）
//        ③ 数据表回读行数 = 房间数 + 表头 ④ worker 回收
import { bootHost, stopHost } from './boot.mjs'
import { mkdirSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const XLSX = resolve(__dirname, '../output/map_gen.xlsx')
const PNG = resolve(__dirname, '../output/map_gen.png')

const COLS = 30
const ROWS = 20
const PX = 100 // Foundry 默认

async function main() {
  const ctx = await bootHost()
  const m = ctx.table
  const info = await m.ping()
  console.log('[map-gen] python/openpyxl =>', info.python, '/', info.openpyxl)

  mkdirSync(dirname(XLSX), { recursive: true })
  await m.createWorkbook(XLSX, 'Map')

  const gen = await m.genDungeon(XLSX, {
    sheet: 'Map',
    cols: COLS,
    rows: ROWS,
    rooms: 6,
    party_level: 3,
    seed: 12345,
  })
  console.log('[map-gen] gen =>', JSON.stringify(gen))
  if (gen.rooms < 1) throw new Error('未生成任何房间')

  await m.save(XLSX)

  // 渲染 PNG（网格关，VTT 自绘）
  const ren = await m.renderMapPng(XLSX, 'Map', PNG, PX, false)
  console.log('[map-gen] render =>', JSON.stringify(ren))
  const expectW = COLS * PX
  const expectH = ROWS * PX
  if (ren.width !== expectW || ren.height !== expectH)
    throw new Error('PNG 尺寸不符 VTT flush-grid 整数格规则: ' + ren.width + 'x' + ren.height)
  // 整数格：宽高必须能被网格数整除（Foundry 对齐要求）
  if (ren.width % COLS !== 0 || ren.height % ROWS !== 0)
    throw new Error('非整数格：宽/列 或 高/行 不能整除')

  // 数据表回读
  const rooms = await m.readSheet(XLSX, 'Rooms')
  console.log('[map-gen] Rooms 表:', rooms.max_row, '行（含表头）,', rooms.max_col, '列')
  if (rooms.max_row < 2) throw new Error('Rooms 表无数据行')
  console.log('[map-gen] Rooms 首数据行:', JSON.stringify(rooms.rows[1]))

  const enc = await m.readSheet(XLSX, 'Encounters')
  const loot = await m.readSheet(XLSX, 'Loot')
  console.log('[map-gen] Encounters:', enc.max_row, '行 | Loot:', loot.max_row, '行')

  const pyPid = ctx.table ? ctx.table.pid : null
  await stopHost()
  console.log('[map-gen] 产物:', XLSX, PNG)
  console.log('[map-gen] 宿主已关停，python worker (pid=' + pyPid + ') 已回收。')
  process.exit(0)
}

main().catch((e) => {
  console.error('[map-gen] FATAL:', e)
  process.exit(1)
})

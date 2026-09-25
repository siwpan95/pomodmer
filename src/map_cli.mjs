// map_cli.mjs —— 跑团地图命令行：gen / render（从参数驱动）
// 用法：
//   node src/map_cli.mjs gen   --xlsx output/map.xlsx --cols 30 --rows 20 --rooms 6 --level 3 --seed 12345 --png output/map.png [--px 100] [--grid]
//   node src/map_cli.mjs render --xlsx output/map.xlsx --sheet Map --out output/map.png [--px 100] [--grid]
import { bootHost, stopHost } from './boot.mjs'
import { mkdirSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { parseArgs } from 'node:util'

const __dirname = dirname(fileURLToPath(import.meta.url))

const { positionals, values } = parseArgs({
  allowPositionals: true,
  options: {
    cols: { type: 'string' },
    rows: { type: 'string' },
    rooms: { type: 'string' },
    level: { type: 'string' },
    seed: { type: 'string' },
    sheet: { type: 'string', default: 'Map' },
    xlsx: { type: 'string' },
    png: { type: 'string' },
    out: { type: 'string' },
    json: { type: 'string' },
    name: { type: 'string' },
    px: { type: 'string', default: '100' },
    grid: { type: 'boolean', default: false },
  },
})

const cmd = positionals[0]
const num = (v, d) => (v == null ? d : parseInt(v, 10))

async function main() {
  const ctx = await bootHost()
  const m = ctx.table
  const xlsx = resolve(values.xlsx || 'output/map.xlsx')
  mkdirSync(dirname(xlsx), { recursive: true })

  if (cmd === 'gen') {
    await m.createWorkbook(xlsx, values.sheet)
    const gen = await m.genDungeon(xlsx, {
      sheet: values.sheet,
      cols: num(values.cols, 30),
      rows: num(values.rows, 20),
      rooms: num(values.rooms, 6),
      party_level: num(values.level, 3),
      seed: values.seed ? parseInt(values.seed, 10) : undefined,
    })
    await m.save(xlsx)
    if (values.png) {
      const p = resolve(values.png)
      const ren = await m.renderMapPng(xlsx, values.sheet, p, num(values.px, 100), values.grid)
      console.log('[map-cli] gen', JSON.stringify(gen), '| render', JSON.stringify(ren))
    } else {
      console.log('[map-cli] gen', JSON.stringify(gen))
    }
  } else if (cmd === 'render') {
    const out = resolve(values.out || 'output/map.png')
    const ren = await m.renderMapPng(xlsx, values.sheet, out, num(values.px, 100), values.grid)
    console.log('[map-cli] render', JSON.stringify(ren))
  } else if (cmd === 'export') {
    const out = resolve(values.out || 'output/map.donjon.json')
    const exp = await m.exportDonjon(xlsx, out, {
      sheet: values.sheet,
      name: values.name,
      px: num(values.px, 100),
      seed: values.seed ? parseInt(values.seed, 10) : 0,
      level: num(values.level, 3),
    })
    console.log('[map-cli] export', JSON.stringify(exp))
  } else if (cmd === 'import') {
    const jsonPath = resolve(values.json || 'output/map.donjon.json')
    const imp = await m.importDonjon(xlsx, jsonPath, { sheet: values.sheet })
    await m.save(xlsx)
    console.log('[map-cli] import', JSON.stringify(imp))
    if (values.png) {
      const p = resolve(values.png)
      const ren = await m.renderMapPng(xlsx, values.sheet, p, num(values.px, 100), values.grid)
      console.log('[map-cli] import+render', JSON.stringify(ren))
    }
  } else {
    console.error(
      '用法: node src/map_cli.mjs gen|render|export|import [--xlsx ... --json ... --out ... --cols .. --rows .. --rooms .. --level .. --seed .. --name .. --png .. --sheet .. --px .. --grid]'
    )
    process.exit(2)
  }

  await stopHost()
  process.exit(0)
}

main().catch((e) => {
  console.error('[map-cli] FATAL:', e)
  process.exit(1)
})

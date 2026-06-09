import * as XLSX from 'xlsx'

export interface ResultRow {
  no: number
  code: string
  description: string
  po: string
  qty: number
  containers: Record<string, number>
  left: number
}

export interface ProcessResult {
  rows: ResultRow[]
  containerNames: string[]
}

function findHeaderRowIndex(rows: unknown[][], keyword: string): number {
  for (let i = 0; i < rows.length; i++) {
    if ((rows[i] as unknown[]).some(cell => typeof cell === 'string' && cell.includes(keyword))) {
      return i
    }
  }
  return -1
}

function findColIndex(row: unknown[], keyword: string): number {
  return row.findIndex(cell => typeof cell === 'string' && cell.toUpperCase().includes(keyword.toUpperCase()))
}

function parseCISheet(sheet: XLSX.WorkSheet): { code: string; description: string; qty: number; marks: string }[] {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })

  const headerIdx = findHeaderRowIndex(rows, 'Fixture code')
  if (headerIdx === -1) throw new Error('ไม่พบ header "Fixture code" ในชีท CI')

  const header = rows[headerIdx] as unknown[]
  const codeCol = findColIndex(header, 'Fixture')
  const descCol = findColIndex(header, 'DESCRIPTION')
  const qtyCol = findColIndex(header, 'QTY')
  const marksCol = findColIndex(header, 'MARKS')

  if (codeCol === -1) throw new Error('ไม่พบคอลัมน์ Fixture code ในชีท CI')
  if (qtyCol === -1) throw new Error('ไม่พบคอลัมน์ QTY ในชีท CI')
  if (marksCol === -1) throw new Error('ไม่พบคอลัมน์ MARKS ในชีท CI')

  const items: { code: string; description: string; qty: number; marks: string }[] = []

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    const code = row[codeCol]

    if (!code || typeof code !== 'string' || code.trim() === '') continue
    if (code.toUpperCase().includes('TOTAL') || code.toUpperCase().includes('COUNTRY')) break

    const qty = typeof row[qtyCol] === 'number' ? (row[qtyCol] as number) : parseFloat(String(row[qtyCol])) || 0
    const marks = String(row[marksCol] || '').replace(/\n/g, ' ').trim()
    const desc = descCol !== -1 ? String(row[descCol] || '').trim() : ''

    if (qty > 0) {
      items.push({ code: code.trim(), description: desc, qty, marks })
    }
  }

  return items
}

function parseContainerSheet(sheet: XLSX.WorkSheet): Record<string, number> {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })

  const headerIdx = findHeaderRowIndex(rows, 'Fixture code')
  if (headerIdx === -1) return {}

  const header = rows[headerIdx] as unknown[]
  const codeCol = findColIndex(header, 'Fixture')

  // QTY might not be in header row — check current and next row
  let qtyCol = findColIndex(header, 'QTY')
  if (qtyCol === -1 && headerIdx + 1 < rows.length) {
    qtyCol = findColIndex(rows[headerIdx + 1] as unknown[], 'QTY')
  }

  if (codeCol === -1 || qtyCol === -1) return {}

  const result: Record<string, number> = {}

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    const code = row[codeCol]

    if (!code || typeof code !== 'string' || code.trim() === '') continue
    if (code.toUpperCase().includes('COUNTRY') || code.toUpperCase().includes('SHIPPING')) break

    const qty = typeof row[qtyCol] === 'number' ? (row[qtyCol] as number) : parseFloat(String(row[qtyCol])) || 0

    if (qty > 0) {
      result[code.trim()] = (result[code.trim()] || 0) + qty
    }
  }

  return result
}

export function processExcel(buffer: ArrayBuffer): ProcessResult {
  const workbook = XLSX.read(buffer, { type: 'array' })

  const ciSheetName = workbook.SheetNames.find(n => n.toUpperCase() === 'CI')
  const plSheetName = workbook.SheetNames.find(n => n.toUpperCase() === 'PL')

  if (!ciSheetName) throw new Error('ไม่พบชีท CI ในไฟล์ Excel')

  const containerNames = workbook.SheetNames.filter(n => n !== ciSheetName && n !== plSheetName)

  const ciItems = parseCISheet(workbook.Sheets[ciSheetName])

  const containerMaps: Record<string, Record<string, number>> = {}
  for (const name of containerNames) {
    containerMaps[name] = parseContainerSheet(workbook.Sheets[name])
  }

  // FIFO: group CI rows by code, track remaining qty per row
  const ciByCode: Record<string, { rowIdx: number; remaining: number }[]> = {}
  ciItems.forEach((item, idx) => {
    if (!ciByCode[item.code]) ciByCode[item.code] = []
    ciByCode[item.code].push({ rowIdx: idx, remaining: item.qty })
  })

  // Initialize container assignments per CI row
  const rowContainers: Record<number, Record<string, number>> = {}
  ciItems.forEach((_, idx) => {
    rowContainers[idx] = {}
    for (const name of containerNames) rowContainers[idx][name] = 0
  })

  // For each container (in sheet order), distribute to CI rows FIFO
  for (const containerName of containerNames) {
    for (const [code, containerQty] of Object.entries(containerMaps[containerName])) {
      const ciRows = ciByCode[code]
      if (!ciRows) continue
      let remaining = containerQty
      for (const ciRow of ciRows) {
        if (remaining <= 0) break
        const take = Math.min(remaining, ciRow.remaining)
        rowContainers[ciRow.rowIdx][containerName] += take
        ciRow.remaining -= take
        remaining -= take
      }
    }
  }

  const rows: ResultRow[] = ciItems.map((item, idx) => {
    const containers = rowContainers[idx]
    const totalInContainers = Object.values(containers).reduce((s, v) => s + v, 0)
    return {
      no: idx + 1,
      code: item.code,
      description: item.description,
      po: item.marks,
      qty: item.qty,
      containers,
      left: item.qty - totalInContainers,
    }
  })

  return { rows, containerNames }
}

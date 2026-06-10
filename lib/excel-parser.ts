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
  invoiceNo: string
  totalAmount: number
  currency: string
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

function extractTotalAndCurrency(sheet: XLSX.WorkSheet): { totalAmount: number; currency: string } {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })

  let currency = ''
  let totalAmount = 0

  // Scan header rows for currency code
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    for (const cell of rows[i] as unknown[]) {
      const s = String(cell || '')
      if (/\bCNY\b|\bRMB\b|人民币/.test(s)) { currency = 'CNY'; break }
      if (/\bUSD\b|\bUS\$/.test(s)) { currency = 'USD'; break }
      if (/\bEUR\b/.test(s)) { currency = 'EUR'; break }
      if (/\bJPY\b/.test(s)) { currency = 'JPY'; break }
      if (/\bGBP\b/.test(s)) { currency = 'GBP'; break }
      if (/\bTHB\b/.test(s)) { currency = 'THB'; break }
    }
    if (currency) break
  }

  // Find TOTAL row (search from bottom up)
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i] as unknown[]
    const joined = row.map(c => String(c || '').toUpperCase()).join(' ')
    if (!joined.includes('TOTAL') && !joined.includes('合计') && !joined.includes('AMOUNT IN')) continue

    // Get last numeric value in row
    for (let j = row.length - 1; j >= 0; j--) {
      const cell = row[j]
      const cellStr = String(cell || '')

      if (!currency) {
        if (/¥|￥/.test(cellStr)) currency = 'CNY'
        else if (/\$/.test(cellStr)) currency = 'USD'
        else if (/€/.test(cellStr)) currency = 'EUR'
        else if (/£/.test(cellStr)) currency = 'GBP'
      }

      const num = typeof cell === 'number' ? cell : parseFloat(cellStr.replace(/[^0-9.]/g, ''))
      if (!isNaN(num) && num > 10) {
        totalAmount = num
        break
      }
    }
    if (totalAmount > 0) break
  }

  return { totalAmount, currency }
}

function extractInvoiceNo(sheet: XLSX.WorkSheet): string {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const row = rows[i] as unknown[]
    for (let j = 0; j < row.length; j++) {
      const cell = String(row[j] || '')
      if (cell.toLowerCase().includes('invoice no')) {
        const next = String(row[j + 1] || '').trim()
        if (next) return next
      }
    }
  }
  return ''
}

export function processExcel(buffer: ArrayBuffer): ProcessResult {
  const workbook = XLSX.read(buffer, { type: 'array' })

  const ciSheetName = workbook.SheetNames.find(n => n.toUpperCase() === 'CI')
  const plSheetName = workbook.SheetNames.find(n => n.toUpperCase() === 'PL')

  if (!ciSheetName) throw new Error('ไม่พบชีท CI ในไฟล์ Excel')

  const containerNames = workbook.SheetNames.filter(n => {
    if (n === ciSheetName || n === plSheetName) return false
    const upper = n.toUpperCase()
    // exclude common non-container sheets
    if (['SUMMARY', 'COVER', 'INDEX', 'SHEET'].some(k => upper.includes(k))) return false
    return true
  })

  const invoiceNo = extractInvoiceNo(workbook.Sheets[ciSheetName])
  const { totalAmount, currency } = extractTotalAndCurrency(workbook.Sheets[ciSheetName])
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

  return { rows, containerNames, invoiceNo, totalAmount, currency }
}

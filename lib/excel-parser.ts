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

function findHeaderRowIndex(rows: unknown[][], ...keywords: string[]): number {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    for (const kw of keywords) {
      if (row.some(cell => typeof cell === 'string' && cell.toLowerCase().includes(kw.toLowerCase()))) {
        return i
      }
    }
  }
  return -1
}

function findColIndex(row: unknown[], ...keywords: string[]): number {
  for (const kw of keywords) {
    const idx = row.findIndex(cell => typeof cell === 'string' && cell.toUpperCase().includes(kw.toUpperCase()))
    if (idx !== -1) return idx
  }
  return -1
}

// Exact match (used for short column names like "PO" to avoid false positives)
function findExactColIndex(row: unknown[], keyword: string): number {
  return row.findIndex(cell => typeof cell === 'string' && cell.trim().toUpperCase() === keyword.toUpperCase())
}

function parseCISheet(sheet: XLSX.WorkSheet): { code: string; description: string; qty: number; marks: string }[] {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })

  // Support YONGGUAN ("Fixture code"), LITELON ("Product Code", "Item no"), DC30 ("Item code") formats
  const headerIdx = findHeaderRowIndex(rows, 'Fixture code', 'Product Code', 'Item no', 'Item code')
  if (headerIdx === -1) throw new Error('ไม่พบ header ในชีท CI (Fixture code / Product Code / Item no / Item code)')

  const header = rows[headerIdx] as unknown[]
  const codeCol = findColIndex(header, 'Fixture', 'Product Code', 'Item no', 'Code')
  const qtyCol = findColIndex(header, 'QTY', 'Qty', 'Quantity')

  // PO: exact "PO" (LITELON) → exact "Remarks" (DC30) → "Marks"/"Shipping Marks" (YONGGUAN)
  // "Remarks" must be checked before substring 'MARKS' because "Remarks" contains "marks"
  let marksCol = findExactColIndex(header, 'PO')
  if (marksCol === -1) marksCol = findExactColIndex(header, 'Remarks')
  if (marksCol === -1) marksCol = findColIndex(header, 'MARKS', 'Shipping Marks', 'Shipping')

  const descCol = findColIndex(header, 'DESCRIPTION', 'Description')

  if (codeCol === -1) throw new Error('ไม่พบคอลัมน์ Code ในชีท CI')
  if (qtyCol === -1) throw new Error('ไม่พบคอลัมน์ QTY ในชีท CI')

  const items: { code: string; description: string; qty: number; marks: string }[] = []

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    const rawCode = row[codeCol]
    if (!rawCode || typeof rawCode !== 'string' || rawCode.trim() === '') continue
    const code = rawCode.trim()
    if (code.toUpperCase().includes('TOTAL') || code.toUpperCase().includes('COUNTRY')) break

    const qty = typeof row[qtyCol] === 'number' ? (row[qtyCol] as number) : parseFloat(String(row[qtyCol])) || 0
    const marks = marksCol !== -1 ? String(row[marksCol] || '').replace(/\n/g, ' ').trim() : ''
    const desc = descCol !== -1 ? String(row[descCol] || '').trim() : ''

    if (qty > 0) {
      items.push({ code, description: desc, qty, marks })
    }
  }

  return items
}

// YONGGUAN format: each container is a separate sheet
function parseContainerSheet(sheet: XLSX.WorkSheet): Record<string, number> {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })

  const headerIdx = findHeaderRowIndex(rows, 'Fixture code', 'Product Code', 'Item no')
  if (headerIdx === -1) return {}

  const header = rows[headerIdx] as unknown[]
  const codeCol = findColIndex(header, 'Fixture', 'Product Code', 'Item no')

  let qtyCol = findColIndex(header, 'QTY', 'Qty', 'Quantity')
  if (qtyCol === -1 && headerIdx + 1 < rows.length) {
    qtyCol = findColIndex(rows[headerIdx + 1] as unknown[], 'QTY', 'Qty', 'Quantity')
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

// Matches both "#1-Cntr-FCIU6571413/..." and "Cntr-TIIU044096/..." formats
const CNTR_RE = /(?:#\d+-)?Cntr-([A-Z][A-Z0-9]+)/i
// Matches ISO container/seal format "CSNU2005017/CX478086" — container number is before "/"
const CNTR_ISO_RE = /([A-Z]{4}\d{7})\/[A-Z0-9]+/i

// Check if a sheet contains inline container markers (LITELON or DC30 style)
function hasCombinedPLMarkers(sheet: XLSX.WorkSheet): boolean {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })
  for (let i = 0; i < Math.min(rows.length, 80); i++) {
    for (const cell of rows[i] as unknown[]) {
      const s = String(cell || '')
      if (CNTR_RE.test(s) || CNTR_ISO_RE.test(s)) return true
    }
  }
  return false
}

// LITELON format: single PL sheet with inline container markers (e.g. "#1-Cntr-FCIU6571413/..." or "Cntr-TIIU044096/...")
function parseCombinedPLSheet(sheet: XLSX.WorkSheet): { containerNames: string[]; containerMaps: Record<string, Record<string, number>> } {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })

  const containerMaps: Record<string, Record<string, number>> = {}
  const containerNames: string[] = []
  let currentContainer = ''

  // Detect code and qty column positions from header row
  let codeCol = 1  // default: column B
  let qtyCol = 2   // default: column C

  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const lower = (rows[i] as unknown[]).map(c => String(c || '').toLowerCase())
    const itemIdx = lower.findIndex(c => c.includes('item no') || c.includes('product code') || c.includes('item code'))
    if (itemIdx !== -1) {
      codeCol = itemIdx
      for (let j = itemIdx + 1; j < lower.length; j++) {
        if (lower[j].includes('qty') || lower[j].includes('pcs') || lower[j].includes('quantity')) {
          qtyCol = j
          break
        }
      }
      break
    }
  }

  for (const row of rows as unknown[][]) {
    // Scan every cell for container marker (LITELON or DC30 ISO format)
    let isMarker = false
    for (const cell of row) {
      const s = String(cell || '')
      const match = s.match(CNTR_RE) || s.match(CNTR_ISO_RE)
      if (match) {
        currentContainer = match[1].toUpperCase()
        if (!containerMaps[currentContainer]) {
          containerMaps[currentContainer] = {}
          containerNames.push(currentContainer)
        }
        isMarker = true
        break
      }
    }
    if (isMarker || !currentContainer) continue

    const code = String(row[codeCol] || '').trim()
    if (!code) continue
    if (/item no|product code|item code/i.test(code)) continue  // header row
    if (!/^[A-Z0-9]/i.test(code)) continue             // must start alphanumeric
    if (code.toUpperCase().includes('TOTAL')) continue

    const qtyRaw = row[qtyCol]
    const qty = typeof qtyRaw === 'number' ? (qtyRaw as number) : parseFloat(String(qtyRaw || '')) || 0
    if (qty > 0) {
      containerMaps[currentContainer][code] = (containerMaps[currentContainer][code] || 0) + qty
    }
  }

  return { containerNames, containerMaps }
}

function extractTotalAndCurrency(sheet: XLSX.WorkSheet): { totalAmount: number; currency: string } {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })

  let currency = ''
  let totalAmount = 0

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

  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i] as unknown[]
    const joined = row.map(c => String(c || '').toUpperCase()).join(' ')
    if (!joined.includes('TOTAL') && !joined.includes('合计') && !joined.includes('AMOUNT IN')) continue

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
      // Match "invoice no" and LITELON typo "invoive no"
      if (/invoice\s*no|invoive\s*no/i.test(cell)) {
        const match = cell.match(/invo\w+\s*no\.?:?\s*(.+)/i)
        if (match && match[1].trim()) return match[1].trim()
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
  if (!ciSheetName) throw new Error('ไม่พบชีท CI ในไฟล์ Excel')

  // Separate container sheets (YONGGUAN style): non-CI, non-standard sheets
  const EXCLUDE = ['CI', 'PL', 'SUMMARY', 'COVER', 'INDEX', 'SHEET']
  const separateContainerSheets = workbook.SheetNames.filter(n => {
    const upper = n.toUpperCase()
    if (EXCLUDE.some(k => upper === k || upper.includes(k))) return false
    return true
  })

  const invoiceNo = extractInvoiceNo(workbook.Sheets[ciSheetName])
  const { totalAmount, currency } = extractTotalAndCurrency(workbook.Sheets[ciSheetName])
  const ciItems = parseCISheet(workbook.Sheets[ciSheetName])

  let containerNames: string[]
  let containerMaps: Record<string, Record<string, number>>

  if (separateContainerSheets.length > 0) {
    // YONGGUAN format: separate sheet per container
    containerNames = separateContainerSheets
    containerMaps = {}
    for (const name of containerNames) {
      containerMaps[name] = parseContainerSheet(workbook.Sheets[name])
    }
  } else {
    // LITELON format: find any sheet that has inline container markers (#N-Cntr-...)
    // Sheet may be named "PL", "Sheet1", or anything else
    const combinedSheet = workbook.SheetNames
      .filter(n => n !== ciSheetName)
      .find(n => hasCombinedPLMarkers(workbook.Sheets[n]))

    if (combinedSheet) {
      const parsed = parseCombinedPLSheet(workbook.Sheets[combinedSheet])
      containerNames = parsed.containerNames
      containerMaps = parsed.containerMaps
    } else {
      containerNames = []
      containerMaps = {}
    }
  }

  // FIFO matching
  const ciByCode: Record<string, { rowIdx: number; remaining: number }[]> = {}
  ciItems.forEach((item, idx) => {
    if (!ciByCode[item.code]) ciByCode[item.code] = []
    ciByCode[item.code].push({ rowIdx: idx, remaining: item.qty })
  })

  const rowContainers: Record<number, Record<string, number>> = {}
  ciItems.forEach((_, idx) => {
    rowContainers[idx] = {}
    for (const name of containerNames) rowContainers[idx][name] = 0
  })

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

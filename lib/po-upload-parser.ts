import * as XLSX from 'xlsx'

export interface POUploadRow {
  item_code: string
  description: string
  qty: number
  unit_price: number
  total: number
}

export interface POUploadResult {
  rows: POUploadRow[]
  currency: 'CNY' | 'USD'
  total_amount: number
}

export function parsePOUploadExcel(buffer: ArrayBuffer): POUploadResult {
  const wb = XLSX.read(buffer, { type: 'array' })
  // Pick first visible sheet
  const visibleName =
    wb.SheetNames.find((_, i) => !wb.Workbook?.Sheets?.[i]?.Hidden) ??
    wb.SheetNames[0]
  const ws = wb.Sheets[visibleName]
  const raw = XLSX.utils.sheet_to_json<(string | number)[]>(ws, {
    header: 1,
    defval: '',
  }) as (string | number)[][]

  // Find header row — must contain "item" and "code" somewhere
  let headerRow = -1
  let itemCol = -1, descCol = -1, qtyCol = -1, priceCol = -1
  let currency: 'CNY' | 'USD' = 'CNY'

  for (let i = 0; i < Math.min(raw.length, 12); i++) {
    const row = raw[i]
    let found = false
    for (let j = 0; j < row.length; j++) {
      const h = String(row[j]).toLowerCase().trim()
      if (h.includes('item') && h.includes('code')) { itemCol = j; found = true }
      if (h.includes('desc') || h.includes('สินค้า') || h.includes('name')) descCol = j
      if (h === 'qty' || h === 'quantity' || h === 'จำนวน') qtyCol = j
      if (h.includes('unit') && h.includes('price')) {
        priceCol = j
        const raw_h = String(row[j]).toUpperCase()
        if (raw_h.includes('USD')) currency = 'USD'
      }
    }
    if (found && priceCol >= 0) { headerRow = i; break }
  }

  if (headerRow < 0) {
    throw new Error(
      'ไม่พบ header — ต้องมีคอลัมน์ "Item Code" และ "Unit Price" ในไฟล์'
    )
  }

  // Fallbacks if columns not found
  if (descCol < 0) descCol = itemCol + 1
  if (qtyCol < 0) qtyCol = priceCol - 1

  const rows: POUploadRow[] = []

  for (let i = headerRow + 1; i < raw.length; i++) {
    const row = raw[i]
    const itemCode = String(row[itemCol] ?? '').trim()

    // Stop at empty item_code or TOTAL row
    if (!itemCode) continue
    if (itemCode.toUpperCase() === 'TOTAL' || itemCode.toUpperCase().startsWith('REMARK')) break

    const qty = parseFloat(String(row[qtyCol] ?? '')) || 0
    const unitPrice = parseFloat(String(row[priceCol] ?? '')) || 0
    if (unitPrice === 0 && qty === 0) continue

    rows.push({
      item_code: itemCode,
      description: String(row[descCol] ?? '').trim(),
      qty,
      unit_price: unitPrice,
      total: qty * unitPrice,
    })
  }

  if (rows.length === 0) throw new Error('ไม่พบข้อมูลสินค้าในไฟล์')

  const total_amount = rows.reduce((s, r) => s + r.total, 0)
  return { rows, currency, total_amount }
}

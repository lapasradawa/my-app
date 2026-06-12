import * as XLSX from 'xlsx'

export interface POItem {
  item_code: string
  description: string
  fob_price: number
  currency: string
  document_no: string
  document_date: string
}

export function parsePO(buffer: ArrayBuffer): { items: POItem[]; currency: string } {
  const workbook = XLSX.read(buffer, { type: 'array' })
  const sheetName = workbook.SheetNames[0]
  const sheet = workbook.Sheets[sheetName]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' })

  let document_no = ''
  let document_date = ''

  for (let i = 0; i < Math.min(rows.length, 25); i++) {
    const row = rows[i] as unknown[]
    for (let j = 0; j < row.length - 1; j++) {
      const s = String(row[j] || '').toLowerCase()
      if (/document\s*no/.test(s) && !document_no) document_no = String(row[j + 1] || '').trim()
      if (/document\s*date/.test(s) && !document_date) document_date = String(row[j + 1] || '').trim()
    }
  }

  let headerIdx = -1
  let itemCodeCol = -1
  let descCol = -1
  let priceCol = -1
  let currency = 'CNY'

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    const found = row.findIndex(c => /item\s*code/i.test(String(c || '')))
    if (found !== -1) {
      headerIdx = i
      itemCodeCol = found
      for (let j = 0; j < row.length; j++) {
        const cell = String(row[j] || '')
        if (/description/i.test(cell) && descCol === -1) descCol = j
        if (/unit\s*price/i.test(cell) && priceCol === -1) {
          priceCol = j
          const m = cell.match(/\(([A-Z]{2,3})\/PC\)/i) || cell.match(/\(([A-Z]{2,3})\)/i)
          if (m) currency = m[1].toUpperCase()
        }
      }
      break
    }
  }

  if (headerIdx === -1) throw new Error('ไม่พบ header "Item Code" ในไฟล์ PO')
  if (priceCol === -1) throw new Error('ไม่พบคอลัมน์ UNIT PRICE ในไฟล์ PO')

  const items: POItem[] = []
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    const code = String(row[itemCodeCol] || '').trim()
    if (!code) continue
    if (/^total/i.test(code)) break
    if (!/^[A-Z0-9]/i.test(code)) continue
    const rawPrice = row[priceCol]
    const price = typeof rawPrice === 'number' ? rawPrice : parseFloat(String(rawPrice || ''))
    if (isNaN(price) || price <= 0) continue
    const desc = descCol !== -1 ? String(row[descCol] || '').trim() : ''
    items.push({ item_code: code, description: desc, fob_price: price, currency, document_no, document_date })
  }

  return { items, currency }
}

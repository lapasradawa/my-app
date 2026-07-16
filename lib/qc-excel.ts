import ExcelJS from 'exceljs'

export interface QCItem {
  item_code: string
  product_description: string
  qty: number
  unit_price: number
  total: number
  qty_defective: number
  remark: string
}

export interface QCReportData {
  report_no: string
  supplier_company: string
  destuffing_date: string
  issue_found_date: string
  invoice_no: string
  po_no: string
  attachment_desc: string
  description: string
  items: QCItem[]
  corrective_actions: string[]
  corrective_action_comment: string
  root_cause: string
  preventive_action: string
  verification_accepted: boolean | null
  verification_comment: string
}

const BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' }, bottom: { style: 'thin' },
  left: { style: 'thin' }, right: { style: 'thin' },
}
const THICK: Partial<ExcelJS.Borders> = {
  top: { style: 'medium' }, bottom: { style: 'medium' },
  left: { style: 'medium' }, right: { style: 'medium' },
}

function applyBorder(ws: ExcelJS.Worksheet, r: number, c: number, border = BORDER) {
  const cell = ws.getCell(r, c)
  cell.border = border
}
function applyRangeBorder(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  for (let r = r1; r <= r2; r++)
    for (let c = c1; c <= c2; c++)
      applyBorder(ws, r, c)
}

function cell(ws: ExcelJS.Worksheet, r: number, c: number, value: ExcelJS.CellValue, opts?: {
  bold?: boolean; size?: number; color?: string; bg?: string; align?: ExcelJS.Alignment['horizontal']; wrap?: boolean; valign?: ExcelJS.Alignment['vertical']
}) {
  const cl = ws.getCell(r, c)
  cl.value = value
  cl.font = { name: 'Arial', size: opts?.size ?? 9, bold: opts?.bold ?? false, color: { argb: opts?.color ?? 'FF000000' } }
  if (opts?.bg) cl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.bg } }
  cl.alignment = { horizontal: opts?.align ?? 'left', vertical: opts?.valign ?? 'middle', wrapText: opts?.wrap ?? false }
}

function merge(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  ws.mergeCells(r1, c1, r2, c2)
}

export async function exportQCReportExcel(data: QCReportData) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('QC Report', {
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } },
  })

  // Columns: A=1 B=2 C=3 D=4 E=5 F=6 G=7 H=8
  ws.columns = [
    { width: 6 },   // A - NO.
    { width: 22 },  // B - ITEM CODE
    { width: 30 },  // C - PRODUCT DESCRIPTION
    { width: 10 },  // D - QTY
    { width: 12 },  // E - UNIT PRICE
    { width: 14 },  // F - TOTAL
    { width: 10 },  // G - QTY DEFECTIVE
    { width: 20 },  // H - REMARK
  ]

  // ── Row 1: Logo + Title ──────────────────────────────────────────
  ws.getRow(1).height = 30
  merge(ws, 1, 1, 2, 1)
  cell(ws, 1, 1, 'rbs', { bold: true, size: 18, color: 'FF004080' })

  merge(ws, 1, 2, 2, 6)
  cell(ws, 1, 2, 'QUALITY CLAIM REPORT', { bold: true, size: 14, align: 'center' })
  applyRangeBorder(ws, 1, 2, 2, 6)

  merge(ws, 1, 7, 1, 8)
  cell(ws, 1, 7, 'Report No.', { bold: true, size: 9, bg: 'FFD9D9D9' })
  applyRangeBorder(ws, 1, 7, 1, 8)

  ws.getRow(2).height = 16
  merge(ws, 2, 7, 2, 8)
  cell(ws, 2, 7, data.report_no, { bold: true, size: 10, align: 'center' })
  applyRangeBorder(ws, 2, 7, 2, 8)

  // ── Rows 3-6: Meta fields ────────────────────────────────────────
  const metaRows = [
    ['Subject :', 'QUALITY CLAIM', 'Supplier company :', data.supplier_company, 'Destuffing Date :', data.destuffing_date],
    ['Customer Company', 'RETAIL BUSINESS SOLUTION CO., LTD', null, null, 'Issue Found Date :', data.issue_found_date],
    ['Address :', '387 SUKHONTHASAWAT RD., LADPRAO, LADPRAO, BANGKOK, THAILAND 10230', 'Invoice :', data.invoice_no, 'Attachment :', data.attachment_desc || 'PO, Photos'],
    [null, null, 'PO No. :', data.po_no, null, null],
  ]

  metaRows.forEach((row, ri) => {
    const r = 3 + ri
    ws.getRow(r).height = 14
    // Col A-B: left meta
    if (row[0] !== null) { cell(ws, r, 1, row[0], { bold: true, size: 8 }); applyBorder(ws, r, 1) }
    if (row[1] !== null) { merge(ws, r, 2, r, 3); cell(ws, r, 2, row[1], { size: 8, wrap: true }); applyRangeBorder(ws, r, 2, r, 3) }
    // Col C-E: middle meta
    if (row[2] !== null) { cell(ws, r, 4, row[2], { bold: true, size: 8 }); applyBorder(ws, r, 4) }
    if (row[3] !== null) { cell(ws, r, 5, row[3], { size: 8 }); applyBorder(ws, r, 5) }
    // Col F-H: right meta
    if (row[4] !== null) { cell(ws, r, 6, row[4], { bold: true, size: 8 }); applyBorder(ws, r, 6) }
    if (row[5] !== null) { merge(ws, r, 7, r, 8); cell(ws, r, 7, row[5], { size: 8 }); applyRangeBorder(ws, r, 7, r, 8) }
  })

  // ── Part 1: ISSUE ────────────────────────────────────────────────
  let r = 8
  ws.getRow(r).height = 14
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, 'Part 1 : ISSUE', { bold: true, size: 9, bg: 'FFD9D9D9' })
  applyRangeBorder(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 12
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, 'Description :', { bold: true, size: 8 })
  applyRangeBorder(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = Math.max(14, Math.ceil((data.description || '').length / 100) * 12)
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, data.description || '', { size: 8, wrap: true })
  applyRangeBorder(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 14
  const headers = ['NO.', 'ITEM CODE', 'PRODUCT DESCRIPTION', 'QTY\n(PCS)', 'UNIT PRICE\n(CNY)', 'TOTAL\n(CNY)', 'QTY DEFECTIVE\n(PCS)', 'REMARK']
  headers.forEach((h, ci) => {
    cell(ws, r, ci + 1, h, { bold: true, size: 8, align: 'center', bg: 'FFD9D9D9', wrap: true })
    applyBorder(ws, r, ci + 1)
  })

  let totalQty = 0, totalAmount = 0, totalDefective = 0
  data.items.forEach((item, i) => {
    r++; ws.getRow(r).height = 14
    const vals = [i + 1, item.item_code, item.product_description, item.qty, item.unit_price, item.total, item.qty_defective, item.remark]
    vals.forEach((v, ci) => {
      cell(ws, r, ci + 1, v, { size: 8, align: ci >= 3 && ci <= 6 ? 'right' : 'left' })
      applyBorder(ws, r, ci + 1)
    })
    totalQty += item.qty || 0
    totalAmount += item.total || 0
    totalDefective += item.qty_defective || 0
  })

  r++; ws.getRow(r).height = 13
  merge(ws, r, 1, r, 3)
  cell(ws, r, 1, 'Total', { bold: true, size: 8, align: 'center', bg: 'FFD9D9D9' })
  applyRangeBorder(ws, r, 1, r, 3)
  cell(ws, r, 4, totalQty, { bold: true, size: 8, align: 'right', bg: 'FFD9D9D9' }); applyBorder(ws, r, 4)
  cell(ws, r, 5, '', { bg: 'FFD9D9D9' }); applyBorder(ws, r, 5)
  cell(ws, r, 6, totalAmount, { bold: true, size: 8, align: 'right', bg: 'FFD9D9D9' }); applyBorder(ws, r, 6)
  cell(ws, r, 7, totalDefective, { bold: true, size: 8, align: 'right', bg: 'FFD9D9D9' }); applyBorder(ws, r, 7)
  cell(ws, r, 8, '', { bg: 'FFD9D9D9' }); applyBorder(ws, r, 8)

  // ── Part 2: CORRECTIVE ACTION ────────────────────────────────────
  r += 2; ws.getRow(r).height = 14
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, 'Part 2 : CORRECTIVE ACTION', { bold: true, size: 9, bg: 'FFD9D9D9' })
  applyRangeBorder(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 12
  merge(ws, r, 1, r, 3)
  cell(ws, r, 1, 'CORRECTIVE ACTION', { bold: true, size: 8, bg: 'FFD9D9D9' }); applyRangeBorder(ws, r, 1, r, 3)
  merge(ws, r, 4, r, 8)
  cell(ws, r, 4, 'DESCRIPTION / COMMENT', { bold: true, size: 8, bg: 'FFD9D9D9' }); applyRangeBorder(ws, r, 4, r, 8)

  const caOptions = ['REPLACEMENT IN NEXT SHIPMENT', 'CREDIT NOTE / REFUND', 'REWORK / REPAIR', 'OTHER']
  const startCA = r + 1
  caOptions.forEach((opt, i) => {
    r++; ws.getRow(r).height = 13
    const checked = data.corrective_actions.includes(opt)
    merge(ws, r, 1, r, 3)
    cell(ws, r, 1, `${checked ? '☑' : '☐'}  ${opt}`, { size: 8 }); applyRangeBorder(ws, r, 1, r, 3)
  })
  merge(ws, startCA, 4, r, 8)
  cell(ws, startCA, 4, data.corrective_action_comment || '', { size: 8, wrap: true, valign: 'top' })
  applyRangeBorder(ws, startCA, 4, r, 8)

  // ── Part 3: PREVENTIVE ACTION ────────────────────────────────────
  r += 2; ws.getRow(r).height = 14
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, 'Part 3 : PREVENTIVE ACTION (FILLED BY SUPPLIER)', { bold: true, size: 9, bg: 'FFD9D9D9' })
  applyRangeBorder(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 12
  merge(ws, r, 1, r, 4)
  cell(ws, r, 1, 'ROOT CAUSE :', { bold: true, size: 8, bg: 'FFD9D9D9' }); applyRangeBorder(ws, r, 1, r, 4)
  merge(ws, r, 5, r, 8)
  cell(ws, r, 5, 'ACTION :', { bold: true, size: 8, bg: 'FFD9D9D9' }); applyRangeBorder(ws, r, 5, r, 8)

  r++; ws.getRow(r).height = 50
  merge(ws, r, 1, r, 4)
  cell(ws, r, 1, data.root_cause || '', { size: 8, wrap: true, valign: 'top' }); applyRangeBorder(ws, r, 1, r, 4)
  merge(ws, r, 5, r, 8)
  cell(ws, r, 5, data.preventive_action || '', { size: 8, wrap: true, valign: 'top' }); applyRangeBorder(ws, r, 5, r, 8)

  // ── Part 4: VERIFICATION ─────────────────────────────────────────
  r += 2; ws.getRow(r).height = 14
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, 'Part 4 : VERIFICATION STATUS (FILLED BY RBS)', { bold: true, size: 9, bg: 'FFD9D9D9' })
  applyRangeBorder(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 13
  merge(ws, r, 1, r, 3)
  cell(ws, r, 1, `${data.verification_accepted === true ? '☑' : '☐'}  ACCEPTED`, { size: 9 }); applyRangeBorder(ws, r, 1, r, 3)
  merge(ws, r, 4, r, 8)
  cell(ws, r, 4, `${data.verification_accepted === false ? '☑' : '☐'}  NOT ACCEPTED     COMMENT : ${data.verification_comment || ''}`, { size: 8 }); applyRangeBorder(ws, r, 4, r, 8)

  r += 2; ws.getRow(r).height = 40
  merge(ws, r, 1, r, 4)
  cell(ws, r, 1, 'FOLLOW-UP INSPECTOR\n\n\n(Mr. Weerapong Choungkrai)\nDATE…………/…………/…………', { size: 8, align: 'center', wrap: true }); applyRangeBorder(ws, r, 1, r, 4)
  merge(ws, r, 5, r, 8)
  cell(ws, r, 5, 'FOLLOW-UP APPROVAL\n\n\n(Mr. Noppharat Sriwichai)\nDATE…………/…………/…………', { size: 8, align: 'center', wrap: true }); applyRangeBorder(ws, r, 5, r, 8)

  // ── Generate file ─────────────────────────────────────────────────
  const buf = await wb.xlsx.writeBuffer()
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url; a.download = `QC-Report-${data.report_no}.xlsx`; a.click()
  URL.revokeObjectURL(url)
}

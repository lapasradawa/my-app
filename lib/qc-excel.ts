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

function applyBorder(ws: ExcelJS.Worksheet, r: number, c: number, border = BORDER) {
  ws.getCell(r, c).border = border
}
function applyRangeBorder(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  for (let r = r1; r <= r2; r++)
    for (let c = c1; c <= c2; c++)
      applyBorder(ws, r, c)
}

function cell(ws: ExcelJS.Worksheet, r: number, c: number, value: ExcelJS.CellValue, opts?: {
  bold?: boolean; size?: number; color?: string; bg?: string
  align?: ExcelJS.Alignment['horizontal']; wrap?: boolean; valign?: ExcelJS.Alignment['vertical']
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

  // Load logo image
  let logoId: number | null = null
  try {
    const res = await fetch('/rbs-logo.png')
    const buf = await res.arrayBuffer()
    logoId = wb.addImage({ buffer: buf, extension: 'png' })
  } catch { /* logo optional */ }

  const ws = wb.addWorksheet('QC Report', {
    pageSetup: {
      paperSize: 9, orientation: 'portrait', fitToPage: true,
      fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    },
  })

  // Columns: A=1 B=2 C=3 D=4 E=5 F=6 G=7 H=8
  ws.columns = [
    { width: 6 },   // A - NO.
    { width: 16 },  // B - ITEM CODE / label
    { width: 28 },  // C - PRODUCT DESCRIPTION / value
    { width: 10 },  // D - QTY / label
    { width: 14 },  // E - UNIT PRICE / value
    { width: 14 },  // F - TOTAL / label
    { width: 12 },  // G - QTY DEFECTIVE / value
    { width: 18 },  // H - REMARK
  ]

  // ── Row 1-2: Logo + Title ─────────────────────────────────────────
  ws.getRow(1).height = 22
  ws.getRow(2).height = 18

  // Logo (col A, rows 1-2)
  merge(ws, 1, 1, 2, 1)
  if (logoId !== null) {
    ws.addImage(logoId, { tl: { col: 0, row: 0 }, br: { col: 1, row: 2 }, editAs: 'oneCell' })
  } else {
    cell(ws, 1, 1, 'rbs', { bold: true, size: 16, color: 'FF004080', align: 'center' })
  }

  // Title (cols B-F, rows 1-2)
  merge(ws, 1, 2, 2, 6)
  cell(ws, 1, 2, 'QUALITY CLAIM REPORT', { bold: true, size: 14, align: 'center' })
  applyRangeBorder(ws, 1, 2, 2, 6)

  // Report No. box (cols G-H)
  merge(ws, 1, 7, 1, 8)
  cell(ws, 1, 7, 'Report No.', { bold: true, size: 9, bg: 'FFD9D9D9', align: 'center' })
  applyRangeBorder(ws, 1, 7, 1, 8)

  merge(ws, 2, 7, 2, 8)
  cell(ws, 2, 7, data.report_no, { bold: true, size: 10, align: 'center' })
  applyRangeBorder(ws, 2, 7, 2, 8)

  // ── Rows 3-6: Meta fields ─────────────────────────────────────────
  // Layout (changed): Invoice/PO moved UP to rows 3-4 (right side)
  // Row 3: Subject        | Supplier Company            | Destuffing Date
  // Row 4: Customer Co.   | [Supplier continued]        | Issue Found Date
  // Row 5: Address        | Invoice                     | Attachment
  // Row 6: [Addr cont.]   | PO No.                      | [blank]

  const ROW_H = 16

  // Row 3
  ws.getRow(3).height = ROW_H
  cell(ws, 3, 1, 'Subject :', { bold: true, size: 8 }); applyBorder(ws, 3, 1)
  merge(ws, 3, 2, 3, 3)
  cell(ws, 3, 2, 'QUALITY CLAIM', { size: 8, wrap: true }); applyRangeBorder(ws, 3, 2, 3, 3)
  cell(ws, 3, 4, 'Supplier company :', { bold: true, size: 8 }); applyBorder(ws, 3, 4)
  merge(ws, 3, 5, 4, 5)
  cell(ws, 3, 5, data.supplier_company, { size: 8, wrap: true, valign: 'top' }); applyRangeBorder(ws, 3, 5, 4, 5)
  cell(ws, 3, 6, 'Destuffing Date :', { bold: true, size: 8 }); applyBorder(ws, 3, 6)
  merge(ws, 3, 7, 3, 8)
  cell(ws, 3, 7, data.destuffing_date, { size: 8 }); applyRangeBorder(ws, 3, 7, 3, 8)

  // Row 4
  ws.getRow(4).height = ROW_H
  cell(ws, 4, 1, 'Customer Company :', { bold: true, size: 8, wrap: true }); applyBorder(ws, 4, 1)
  merge(ws, 4, 2, 4, 3)
  cell(ws, 4, 2, 'RETAIL BUSINESS SOLUTION CO., LTD', { size: 8, wrap: true }); applyRangeBorder(ws, 4, 2, 4, 3)
  cell(ws, 4, 4, '', { size: 8 }); applyBorder(ws, 4, 4)
  // col 5 already merged above
  cell(ws, 4, 6, 'Issue Found Date :', { bold: true, size: 8 }); applyBorder(ws, 4, 6)
  merge(ws, 4, 7, 4, 8)
  cell(ws, 4, 7, data.issue_found_date, { size: 8 }); applyRangeBorder(ws, 4, 7, 4, 8)

  // Row 5
  ws.getRow(5).height = ROW_H
  cell(ws, 5, 1, 'Address :', { bold: true, size: 8 }); applyBorder(ws, 5, 1)
  merge(ws, 5, 2, 6, 3)
  cell(ws, 5, 2, '387 SUKHONTHASAWAT RD., LADPRAO, LADPRAO, BANGKOK, THAILAND 10230', { size: 8, wrap: true, valign: 'top' }); applyRangeBorder(ws, 5, 2, 6, 3)
  cell(ws, 5, 4, 'Invoice :', { bold: true, size: 8 }); applyBorder(ws, 5, 4)
  merge(ws, 5, 5, 5, 5)
  cell(ws, 5, 5, data.invoice_no, { size: 8 }); applyBorder(ws, 5, 5)
  cell(ws, 5, 6, 'Attachment :', { bold: true, size: 8 }); applyBorder(ws, 5, 6)
  merge(ws, 5, 7, 5, 8)
  cell(ws, 5, 7, data.attachment_desc || 'PO, Photos', { size: 8 }); applyRangeBorder(ws, 5, 7, 5, 8)

  // Row 6
  ws.getRow(6).height = ROW_H
  cell(ws, 6, 1, '', { size: 8 }); applyBorder(ws, 6, 1)
  // cols 2-3 already merged above
  cell(ws, 6, 4, 'PO No. :', { bold: true, size: 8 }); applyBorder(ws, 6, 4)
  cell(ws, 6, 5, data.po_no, { size: 8 }); applyBorder(ws, 6, 5)
  cell(ws, 6, 6, '', { size: 8 }); applyBorder(ws, 6, 6)
  merge(ws, 6, 7, 6, 8)
  cell(ws, 6, 7, '', { size: 8 }); applyRangeBorder(ws, 6, 7, 6, 8)

  // ── Part 1: ISSUE ─────────────────────────────────────────────────
  let r = 8
  ws.getRow(r).height = 14
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, 'Part 1 : ISSUE', { bold: true, size: 9, bg: 'FFD9D9D9' })
  applyRangeBorder(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 12
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, 'Description :', { bold: true, size: 8 })
  applyRangeBorder(ws, r, 1, r, 8)

  // Description text — give generous fixed height
  r++
  const descLines = Math.max(4, Math.ceil((data.description || '').length / 120))
  ws.getRow(r).height = descLines * 14
  merge(ws, r, 1, r, 8)
  cell(ws, r, 1, data.description || '', { size: 9, wrap: true, valign: 'top' })
  applyRangeBorder(ws, r, 1, r, 8)

  // Item table header — taller row so 2-line headers aren't squished
  r++; ws.getRow(r).height = 28
  const headers = ['NO.', 'ITEM CODE', 'PRODUCT DESCRIPTION', 'QTY\n(PCS)', 'UNIT PRICE\n(CNY)', 'TOTAL\n(CNY)', 'QTY\nDEFECTIVE', 'REMARK']
  headers.forEach((h, ci) => {
    cell(ws, r, ci + 1, h, { bold: true, size: 8, align: 'center', bg: 'FFD9D9D9', wrap: true })
    applyBorder(ws, r, ci + 1)
  })

  let totalQty = 0, totalAmount = 0, totalDefective = 0
  data.items.forEach((item, i) => {
    r++; ws.getRow(r).height = 16
    const vals = [i + 1, item.item_code, item.product_description, item.qty, item.unit_price, item.total, item.qty_defective, item.remark]
    vals.forEach((v, ci) => {
      cell(ws, r, ci + 1, v, { size: 8, align: ci >= 3 && ci <= 6 ? 'right' : 'left', wrap: ci === 2 })
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

  // ── Part 2: CORRECTIVE ACTION ─────────────────────────────────────
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
  caOptions.forEach((opt) => {
    r++; ws.getRow(r).height = 13
    const checked = data.corrective_actions.includes(opt)
    merge(ws, r, 1, r, 3)
    cell(ws, r, 1, `${checked ? '☑' : '☐'}  ${opt}`, { size: 8 }); applyRangeBorder(ws, r, 1, r, 3)
  })
  merge(ws, startCA, 4, r, 8)
  cell(ws, startCA, 4, data.corrective_action_comment || '', { size: 8, wrap: true, valign: 'top' })
  applyRangeBorder(ws, startCA, 4, r, 8)

  // ── Part 3: PREVENTIVE ACTION ─────────────────────────────────────
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

  // ── Part 4: VERIFICATION ──────────────────────────────────────────
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

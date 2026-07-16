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

function ab(ws: ExcelJS.Worksheet, r: number, c: number) { ws.getCell(r, c).border = BORDER }
function rb(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) ab(ws, r, c)
}
function mg(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  ws.mergeCells(r1, c1, r2, c2)
}
function cl(ws: ExcelJS.Worksheet, r: number, c: number, value: ExcelJS.CellValue, opts?: {
  bold?: boolean; size?: number; color?: string; bg?: string
  align?: ExcelJS.Alignment['horizontal']; wrap?: boolean; valign?: ExcelJS.Alignment['vertical']
}) {
  const cell = ws.getCell(r, c)
  cell.value = value
  cell.font = { name: 'Arial', size: opts?.size ?? 9, bold: opts?.bold ?? false, color: { argb: opts?.color ?? 'FF000000' } }
  if (opts?.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.bg } }
  cell.alignment = { horizontal: opts?.align ?? 'left', vertical: opts?.valign ?? 'middle', wrapText: opts?.wrap ?? false }
}

export async function exportQCReportExcel(data: QCReportData) {
  const wb = new ExcelJS.Workbook()

  let logoId: number | null = null
  try {
    const res = await fetch('/rbs-logo.png')
    const buf = await res.arrayBuffer()
    logoId = wb.addImage({ buffer: buf, extension: 'png' })
  } catch { /* logo optional */ }

  const ws = wb.addWorksheet('QC Report', {
    pageSetup: {
      paperSize: 9,           // A4
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,         // force single page
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
    },
  })

  // ── Column widths: total ~95 chars → fits A4 portrait ─────────────
  // A=logo/NO, B=item code/label, C=desc/value, D=qty/label,
  // E=unit price/value, F=total/label, G=qty def/value, H=remark
  ws.columns = [
    { width: 9 },   // A
    { width: 13 },  // B
    { width: 22 },  // C
    { width: 8 },   // D
    { width: 11 },  // E
    { width: 11 },  // F
    { width: 10 },  // G
    { width: 14 },  // H
  ]
  // Total: 9+13+22+8+11+11+10+14 = 98 ✓

  // ── Rows 1-2: Header ─────────────────────────────────────────────
  ws.getRow(1).height = 22
  ws.getRow(2).height = 16

  // Logo
  mg(ws, 1, 1, 2, 1)
  if (logoId !== null) {
    ws.addImage(logoId, { tl: { col: 0, row: 0 }, br: { col: 1, row: 2 }, editAs: 'oneCell' })
  } else {
    cl(ws, 1, 1, 'rbs', { bold: true, size: 16, color: 'FF004080', align: 'center' })
  }

  // Title
  mg(ws, 1, 2, 2, 6)
  cl(ws, 1, 2, 'QUALITY CLAIM REPORT', { bold: true, size: 13, align: 'center' })
  rb(ws, 1, 2, 2, 6)

  // Report No.
  mg(ws, 1, 7, 1, 8)
  cl(ws, 1, 7, 'Report No.', { bold: true, size: 8, bg: 'FFD9D9D9', align: 'center' })
  rb(ws, 1, 7, 1, 8)
  mg(ws, 2, 7, 2, 8)
  cl(ws, 2, 7, data.report_no, { bold: true, size: 10, align: 'center' })
  rb(ws, 2, 7, 2, 8)

  // ── Rows 3-6: Meta ───────────────────────────────────────────────
  const MH = 15  // meta row height
  ws.getRow(3).height = MH
  ws.getRow(4).height = MH
  ws.getRow(5).height = MH
  ws.getRow(6).height = MH

  // Row 3: Subject | Supplier Company | Destuffing Date
  cl(ws, 3, 1, 'Subject :', { bold: true, size: 8 }); ab(ws, 3, 1)
  mg(ws, 3, 2, 3, 3); cl(ws, 3, 2, 'QUALITY CLAIM', { size: 8 }); rb(ws, 3, 2, 3, 3)
  cl(ws, 3, 4, 'Supplier company :', { bold: true, size: 8 }); ab(ws, 3, 4)
  mg(ws, 3, 5, 4, 5); cl(ws, 3, 5, data.supplier_company, { size: 8, wrap: true, valign: 'top' }); rb(ws, 3, 5, 4, 5)
  cl(ws, 3, 6, 'Destuffing Date :', { bold: true, size: 8 }); ab(ws, 3, 6)
  mg(ws, 3, 7, 3, 8); cl(ws, 3, 7, data.destuffing_date, { size: 8 }); rb(ws, 3, 7, 3, 8)

  // Row 4: Customer Company | [Supplier cont.] | Issue Found Date
  cl(ws, 4, 1, 'Customer Company :', { bold: true, size: 8, wrap: true }); ab(ws, 4, 1)
  mg(ws, 4, 2, 4, 3); cl(ws, 4, 2, 'RETAIL BUSINESS SOLUTION CO., LTD', { size: 8, wrap: true }); rb(ws, 4, 2, 4, 3)
  cl(ws, 4, 4, '', { size: 8 }); ab(ws, 4, 4)
  cl(ws, 4, 6, 'Issue Found Date :', { bold: true, size: 8 }); ab(ws, 4, 6)
  mg(ws, 4, 7, 4, 8); cl(ws, 4, 7, data.issue_found_date, { size: 8 }); rb(ws, 4, 7, 4, 8)

  // Row 5: Address | Invoice | Attachment
  cl(ws, 5, 1, 'Address :', { bold: true, size: 8 }); ab(ws, 5, 1)
  mg(ws, 5, 2, 6, 3); cl(ws, 5, 2, '387 SUKHONTHASAWAT RD., LADPRAO, LADPRAO, BANGKOK, THAILAND 10230', { size: 8, wrap: true, valign: 'top' }); rb(ws, 5, 2, 6, 3)
  cl(ws, 5, 4, 'Invoice :', { bold: true, size: 8 }); ab(ws, 5, 4)
  cl(ws, 5, 5, data.invoice_no, { size: 8 }); ab(ws, 5, 5)
  cl(ws, 5, 6, 'Attachment :', { bold: true, size: 8 }); ab(ws, 5, 6)
  mg(ws, 5, 7, 5, 8); cl(ws, 5, 7, data.attachment_desc || 'PO, Photos', { size: 8 }); rb(ws, 5, 7, 5, 8)

  // Row 6: [Addr cont.] | PO No. | blank
  cl(ws, 6, 1, '', { size: 8 }); ab(ws, 6, 1)
  cl(ws, 6, 4, 'PO No. :', { bold: true, size: 8 }); ab(ws, 6, 4)
  cl(ws, 6, 5, data.po_no, { size: 8 }); ab(ws, 6, 5)
  cl(ws, 6, 6, '', { size: 8 }); ab(ws, 6, 6)
  mg(ws, 6, 7, 6, 8); cl(ws, 6, 7, '', { size: 8 }); rb(ws, 6, 7, 6, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  ws.getRow(7).height = 5

  // ── Part 1: ISSUE ─────────────────────────────────────────────────
  let r = 8
  ws.getRow(r).height = 13
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Part 1 : ISSUE', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 11
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Description :', { bold: true, size: 8 }); rb(ws, r, 1, r, 8)

  r++
  const descH = Math.max(30, Math.min(80, Math.ceil((data.description || '').length / 90) * 12))
  ws.getRow(r).height = descH
  mg(ws, r, 1, r, 8); cl(ws, r, 1, data.description || '', { size: 8, wrap: true, valign: 'top' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 26
  const hdrs = ['NO.', 'ITEM CODE', 'PRODUCT\nDESCRIPTION', 'QTY\n(PCS)', 'UNIT PRICE\n(CNY)', 'TOTAL\n(CNY)', 'QTY\nDEFECTIVE', 'REMARK']
  hdrs.forEach((h, ci) => {
    cl(ws, r, ci + 1, h, { bold: true, size: 8, align: 'center', bg: 'FFD9D9D9', wrap: true })
    ab(ws, r, ci + 1)
  })

  let totalQty = 0, totalAmt = 0, totalDef = 0
  data.items.forEach((item, i) => {
    r++; ws.getRow(r).height = 14
    const vals = [i + 1, item.item_code, item.product_description, item.qty, item.unit_price, item.total, item.qty_defective, item.remark]
    vals.forEach((v, ci) => {
      cl(ws, r, ci + 1, v, { size: 8, align: ci >= 3 && ci <= 6 ? 'right' : 'left', wrap: ci === 2 })
      ab(ws, r, ci + 1)
    })
    totalQty += item.qty || 0; totalAmt += item.total || 0; totalDef += item.qty_defective || 0
  })

  r++; ws.getRow(r).height = 13
  mg(ws, r, 1, r, 3); cl(ws, r, 1, 'Total', { bold: true, size: 8, align: 'center', bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 3)
  cl(ws, r, 4, totalQty, { bold: true, size: 8, align: 'right', bg: 'FFD9D9D9' }); ab(ws, r, 4)
  cl(ws, r, 5, '', { bg: 'FFD9D9D9' }); ab(ws, r, 5)
  cl(ws, r, 6, totalAmt, { bold: true, size: 8, align: 'right', bg: 'FFD9D9D9' }); ab(ws, r, 6)
  cl(ws, r, 7, totalDef, { bold: true, size: 8, align: 'right', bg: 'FFD9D9D9' }); ab(ws, r, 7)
  cl(ws, r, 8, '', { bg: 'FFD9D9D9' }); ab(ws, r, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 5

  // ── Part 2: CORRECTIVE ACTION ─────────────────────────────────────
  r++; ws.getRow(r).height = 13
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Part 2 : CORRECTIVE ACTION', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 11
  mg(ws, r, 1, r, 3); cl(ws, r, 1, 'CORRECTIVE ACTION', { bold: true, size: 8, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 3)
  mg(ws, r, 4, r, 8); cl(ws, r, 4, 'DESCRIPTION / COMMENT', { bold: true, size: 8, bg: 'FFD9D9D9' }); rb(ws, r, 4, r, 8)

  const caOpts = ['REPLACEMENT IN NEXT SHIPMENT', 'CREDIT NOTE / REFUND', 'REWORK / REPAIR', 'OTHER']
  const startCA = r + 1
  caOpts.forEach(opt => {
    r++; ws.getRow(r).height = 13
    const checked = data.corrective_actions.includes(opt)
    mg(ws, r, 1, r, 3); cl(ws, r, 1, `${checked ? '☑' : '☐'}  ${opt}`, { size: 8 }); rb(ws, r, 1, r, 3)
  })
  mg(ws, startCA, 4, r, 8)
  cl(ws, startCA, 4, data.corrective_action_comment || '', { size: 8, wrap: true, valign: 'top' })
  rb(ws, startCA, 4, r, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 5

  // ── Part 3: PREVENTIVE ACTION ─────────────────────────────────────
  r++; ws.getRow(r).height = 13
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Part 3 : PREVENTIVE ACTION (FILLED BY SUPPLIER)', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 11
  mg(ws, r, 1, r, 4); cl(ws, r, 1, 'ROOT CAUSE :', { bold: true, size: 8, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 4)
  mg(ws, r, 5, r, 8); cl(ws, r, 5, 'ACTION :', { bold: true, size: 8, bg: 'FFD9D9D9' }); rb(ws, r, 5, r, 8)

  r++; ws.getRow(r).height = 45
  mg(ws, r, 1, r, 4); cl(ws, r, 1, data.root_cause || '', { size: 8, wrap: true, valign: 'top' }); rb(ws, r, 1, r, 4)
  mg(ws, r, 5, r, 8); cl(ws, r, 5, data.preventive_action || '', { size: 8, wrap: true, valign: 'top' }); rb(ws, r, 5, r, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 5

  // ── Part 4: VERIFICATION ──────────────────────────────────────────
  r++; ws.getRow(r).height = 13
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Part 4 : VERIFICATION STATUS (FILLED BY RBS)', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 13
  mg(ws, r, 1, r, 3)
  cl(ws, r, 1, `${data.verification_accepted === true ? '☑' : '☐'}  ACCEPTED`, { size: 8 }); rb(ws, r, 1, r, 3)
  mg(ws, r, 4, r, 8)
  cl(ws, r, 4, `${data.verification_accepted === false ? '☑' : '☐'}  NOT ACCEPTED     COMMENT : ${data.verification_comment || ''}`, { size: 8 }); rb(ws, r, 4, r, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 5

  // ── Signature ─────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 44
  mg(ws, r, 1, r, 4)
  cl(ws, r, 1, 'FOLLOW-UP INSPECTOR\n\n\n(Mr. Weerapong Choungkrai)\nDATE…………/…………/…………', { size: 8, align: 'center', wrap: true }); rb(ws, r, 1, r, 4)
  mg(ws, r, 5, r, 8)
  cl(ws, r, 5, 'FOLLOW-UP APPROVAL\n\n\n(Mr. Noppharat Sriwichai)\nDATE…………/…………/…………', { size: 8, align: 'center', wrap: true }); rb(ws, r, 5, r, 8)

  // ── Export ────────────────────────────────────────────────────────
  const buf = await wb.xlsx.writeBuffer()
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url; a.download = `QC-Report-${data.report_no}.xlsx`; a.click()
  URL.revokeObjectURL(url)
}

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
  photo_urls?: string[]
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

  // Shared column widths for both sheets (total 98 chars ≈ A4 portrait)
  const COL_WIDTHS = [8, 15, 24, 8, 10, 10, 10, 13]
  // Shared A4 page setup — fitToWidth only, no height shrink
  const A4_SETUP = (): Partial<ExcelJS.PageSetup> => ({
    paperSize: 9, orientation: 'portrait', fitToPage: true,
    fitToWidth: 1, fitToHeight: 0,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
  })

  const ws = wb.addWorksheet('QC Report', { pageSetup: A4_SETUP() })

  ws.columns = COL_WIDTHS.map(w => ({ width: w }))

  // ── Rows 1-2: Header ─────────────────────────────────────────────
  ws.getRow(1).height = 30
  ws.getRow(2).height = 22

  // Logo
  mg(ws, 1, 1, 2, 1)
  if (logoId !== null) {
    ws.addImage(logoId, { tl: { col: 0, row: 0 }, br: { col: 1, row: 2 }, editAs: 'oneCell' })
  } else {
    cl(ws, 1, 1, 'rbs', { bold: true, size: 18, color: 'FF004080', align: 'center' })
  }

  // Title
  mg(ws, 1, 2, 2, 6)
  cl(ws, 1, 2, 'QUALITY CLAIM REPORT', { bold: true, size: 14, align: 'center' })
  rb(ws, 1, 2, 2, 6)

  // Report No.
  mg(ws, 1, 7, 1, 8)
  cl(ws, 1, 7, 'Report No.', { bold: true, size: 9, bg: 'FFD9D9D9', align: 'center' })
  rb(ws, 1, 7, 1, 8)
  mg(ws, 2, 7, 2, 8)
  cl(ws, 2, 7, data.report_no, { bold: true, size: 11, align: 'center' })
  rb(ws, 2, 7, 2, 8)

  // ── Rows 3-6: Meta ───────────────────────────────────────────────
  const MH = 18  // meta row height
  ws.getRow(3).height = MH
  ws.getRow(4).height = MH
  ws.getRow(5).height = MH
  ws.getRow(6).height = MH

  // Row 3: Subject | Supplier Company label (merged D3:D4) + value (merged E3:E4) | Destuffing Date
  cl(ws, 3, 1, 'Subject :', { bold: true, size: 9 }); ab(ws, 3, 1)
  mg(ws, 3, 2, 3, 3); cl(ws, 3, 2, 'QUALITY CLAIM', { size: 9 }); rb(ws, 3, 2, 3, 3)
  mg(ws, 3, 4, 4, 4); cl(ws, 3, 4, 'Supplier company :', { bold: true, size: 9, wrap: true, valign: 'top' }); rb(ws, 3, 4, 4, 4)
  mg(ws, 3, 5, 4, 5); cl(ws, 3, 5, data.supplier_company, { size: 9, wrap: true, valign: 'top' }); rb(ws, 3, 5, 4, 5)
  cl(ws, 3, 6, 'Destuffing Date :', { bold: true, size: 9 }); ab(ws, 3, 6)
  mg(ws, 3, 7, 3, 8); cl(ws, 3, 7, data.destuffing_date, { size: 9 }); rb(ws, 3, 7, 3, 8)

  // Row 4: Customer Company | [D4/E4 already merged above] | Issue Found Date
  cl(ws, 4, 1, 'Customer Company :', { bold: true, size: 9, wrap: true }); ab(ws, 4, 1)
  mg(ws, 4, 2, 4, 3); cl(ws, 4, 2, 'RETAIL BUSINESS SOLUTION CO., LTD', { size: 9, wrap: true }); rb(ws, 4, 2, 4, 3)
  cl(ws, 4, 6, 'Issue Found Date :', { bold: true, size: 9 }); ab(ws, 4, 6)
  mg(ws, 4, 7, 4, 8); cl(ws, 4, 7, data.issue_found_date, { size: 9 }); rb(ws, 4, 7, 4, 8)

  // Rows 5-6: Address (A5:A6) | address text (B-C merged 5-6) | Invoice/PO | Attachment (F5:F6) + value (G-H 5-6)
  mg(ws, 5, 1, 6, 1); cl(ws, 5, 1, 'Address :', { bold: true, size: 9, valign: 'top' }); rb(ws, 5, 1, 6, 1)
  mg(ws, 5, 2, 6, 3); cl(ws, 5, 2, '387 SUKHONTHASAWAT RD., LADPRAO, LADPRAO, BANGKOK, THAILAND 10230', { size: 9, wrap: true, valign: 'top' }); rb(ws, 5, 2, 6, 3)
  cl(ws, 5, 4, 'Invoice :', { bold: true, size: 9 }); ab(ws, 5, 4)
  cl(ws, 5, 5, data.invoice_no, { size: 9 }); ab(ws, 5, 5)
  mg(ws, 5, 6, 6, 6); cl(ws, 5, 6, 'Attachment :', { bold: true, size: 9, valign: 'top' }); rb(ws, 5, 6, 6, 6)
  mg(ws, 5, 7, 6, 8); cl(ws, 5, 7, data.attachment_desc || 'PO, Photos', { size: 9, valign: 'top' }); rb(ws, 5, 7, 6, 8)
  cl(ws, 6, 4, 'PO No. :', { bold: true, size: 9 }); ab(ws, 6, 4)
  cl(ws, 6, 5, data.po_no, { size: 9 }); ab(ws, 6, 5)

  // ── Spacer ───────────────────────────────────────────────────────
  ws.getRow(7).height = 8

  // ── Part 1: ISSUE ─────────────────────────────────────────────────
  let r = 8
  ws.getRow(r).height = 16
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Part 1 : ISSUE', { bold: true, size: 10, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 14
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Description :', { bold: true, size: 9 }); rb(ws, r, 1, r, 8)

  r++
  const descH = Math.max(50, Math.min(120, Math.ceil((data.description || '').length / 80) * 13))
  ws.getRow(r).height = descH
  mg(ws, r, 1, r, 8); cl(ws, r, 1, data.description || '', { size: 9, wrap: true, valign: 'top' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 26
  const hdrs = ['NO.', 'ITEM CODE', 'PRODUCT\nDESCRIPTION', 'QTY\n(PCS)', 'UNIT PRICE\n(CNY)', 'TOTAL\n(CNY)', 'QTY\nDEFECTIVE', 'REMARK']
  hdrs.forEach((h, ci) => {
    cl(ws, r, ci + 1, h, { bold: true, size: 9, align: 'center', bg: 'FFD9D9D9', wrap: true })
    ab(ws, r, ci + 1)
  })

  let totalQty = 0, totalAmt = 0, totalDef = 0
  data.items.forEach((item, i) => {
    r++; ws.getRow(r).height = 20
    const vals = [i + 1, item.item_code, item.product_description, item.qty, item.unit_price, item.total, item.qty_defective, item.remark]
    vals.forEach((v, ci) => {
      cl(ws, r, ci + 1, v, { size: 9, align: ci >= 3 && ci <= 6 ? 'right' : 'left', wrap: ci === 1 || ci === 2 })
      ab(ws, r, ci + 1)
    })
    totalQty += item.qty || 0; totalAmt += item.total || 0; totalDef += item.qty_defective || 0
  })

  r++; ws.getRow(r).height = 16
  mg(ws, r, 1, r, 3); cl(ws, r, 1, 'Total', { bold: true, size: 9, align: 'center', bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 3)
  cl(ws, r, 4, totalQty, { bold: true, size: 9, align: 'right', bg: 'FFD9D9D9' }); ab(ws, r, 4)
  cl(ws, r, 5, '', { bg: 'FFD9D9D9' }); ab(ws, r, 5)
  cl(ws, r, 6, totalAmt, { bold: true, size: 9, align: 'right', bg: 'FFD9D9D9' }); ab(ws, r, 6)
  cl(ws, r, 7, totalDef, { bold: true, size: 9, align: 'right', bg: 'FFD9D9D9' }); ab(ws, r, 7)
  cl(ws, r, 8, '', { bg: 'FFD9D9D9' }); ab(ws, r, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 8

  // ── Part 2: CORRECTIVE ACTION ─────────────────────────────────────
  r++; ws.getRow(r).height = 16
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Part 2 : CORRECTIVE ACTION', { bold: true, size: 10, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 14
  mg(ws, r, 1, r, 4); cl(ws, r, 1, 'CORRECTIVE ACTION', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 4)
  mg(ws, r, 5, r, 8); cl(ws, r, 5, 'DESCRIPTION / COMMENT', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws, r, 5, r, 8)

  const caOpts = ['REPLACEMENT IN NEXT SHIPMENT', 'CREDIT NOTE / REFUND', 'REWORK / REPAIR', 'OTHER']
  const startCA = r + 1
  caOpts.forEach(opt => {
    r++; ws.getRow(r).height = 16
    const checked = data.corrective_actions.includes(opt)
    mg(ws, r, 1, r, 4); cl(ws, r, 1, `${checked ? '☑' : '☐'}  ${opt}`, { size: 9 }); rb(ws, r, 1, r, 4)
  })
  mg(ws, startCA, 5, r, 8)
  cl(ws, startCA, 5, data.corrective_action_comment || '', { size: 9, wrap: true, valign: 'top' })
  rb(ws, startCA, 5, r, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 8

  // ── Part 3: PREVENTIVE ACTION ─────────────────────────────────────
  r++; ws.getRow(r).height = 16
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Part 3 : PREVENTIVE ACTION (FILLED BY SUPPLIER)', { bold: true, size: 10, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 14
  mg(ws, r, 1, r, 4); cl(ws, r, 1, 'ROOT CAUSE :', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 4)
  mg(ws, r, 5, r, 8); cl(ws, r, 5, 'ACTION :', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws, r, 5, r, 8)

  r++; ws.getRow(r).height = 60
  mg(ws, r, 1, r, 4); cl(ws, r, 1, data.root_cause || '', { size: 9, wrap: true, valign: 'top' }); rb(ws, r, 1, r, 4)
  mg(ws, r, 5, r, 8); cl(ws, r, 5, data.preventive_action || '', { size: 9, wrap: true, valign: 'top' }); rb(ws, r, 5, r, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 8

  // ── Part 4: VERIFICATION ──────────────────────────────────────────
  r++; ws.getRow(r).height = 16
  mg(ws, r, 1, r, 8); cl(ws, r, 1, 'Part 4 : VERIFICATION STATUS (FILLED BY RBS)', { bold: true, size: 10, bg: 'FFD9D9D9' }); rb(ws, r, 1, r, 8)

  r++; ws.getRow(r).height = 16
  mg(ws, r, 1, r, 4)
  cl(ws, r, 1, `${data.verification_accepted === true ? '☑' : '☐'}  ACCEPTED`, { size: 9 }); rb(ws, r, 1, r, 4)
  mg(ws, r, 5, r, 8)
  cl(ws, r, 5, `${data.verification_accepted === false ? '☑' : '☐'}  NOT ACCEPTED     COMMENT : ${data.verification_comment || ''}`, { size: 9 }); rb(ws, r, 5, r, 8)

  // ── Spacer ───────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 8

  // ── Signature ─────────────────────────────────────────────────────
  r++; ws.getRow(r).height = 60
  mg(ws, r, 1, r, 4)
  cl(ws, r, 1, 'FOLLOW-UP INSPECTOR\n\n\n(Mr. Weerapong Choungkrai)\nDATE…………/…………/…………', { size: 9, align: 'center', wrap: true }); rb(ws, r, 1, r, 4)
  mg(ws, r, 5, r, 8)
  cl(ws, r, 5, 'FOLLOW-UP APPROVAL\n\n\n(Mr. Noppharat Sriwichai)\nDATE…………/…………/…………', { size: 9, align: 'center', wrap: true }); rb(ws, r, 5, r, 8)

  // ── Sheet 2: Part 5 PHOTO ─────────────────────────────────────────
  const photos = data.photo_urls || []
  if (photos.length > 0) {
    const ws2 = wb.addWorksheet('Part 5 - Photo', { pageSetup: A4_SETUP() })
    ws2.columns = COL_WIDTHS.map(w => ({ width: w }))

    // ── Header (same as sheet 1) ──
    ws2.getRow(1).height = 30; ws2.getRow(2).height = 22
    mg(ws2, 1, 1, 2, 1)
    if (logoId !== null) ws2.addImage(logoId, { tl: { col: 0, row: 0 }, br: { col: 1, row: 2 }, editAs: 'oneCell' })
    else cl(ws2, 1, 1, 'rbs', { bold: true, size: 18, color: 'FF004080', align: 'center' })
    mg(ws2, 1, 2, 2, 6); cl(ws2, 1, 2, 'QUALITY CLAIM REPORT', { bold: true, size: 14, align: 'center' }); rb(ws2, 1, 2, 2, 6)
    mg(ws2, 1, 7, 1, 8); cl(ws2, 1, 7, 'Report No.', { bold: true, size: 9, bg: 'FFD9D9D9', align: 'center' }); rb(ws2, 1, 7, 1, 8)
    mg(ws2, 2, 7, 2, 8); cl(ws2, 2, 7, data.report_no, { bold: true, size: 11, align: 'center' }); rb(ws2, 2, 7, 2, 8)

    // ── Meta rows 3-6 ──
    const MH2 = 18
    ;[3, 4, 5, 6].forEach(rr => { ws2.getRow(rr).height = MH2 })
    cl(ws2, 3, 1, 'Subject :', { bold: true, size: 9 }); ab(ws2, 3, 1)
    mg(ws2, 3, 2, 3, 3); cl(ws2, 3, 2, 'QUALITY CLAIM', { size: 9 }); rb(ws2, 3, 2, 3, 3)
    mg(ws2, 3, 4, 4, 4); cl(ws2, 3, 4, 'Supplier company :', { bold: true, size: 9, wrap: true, valign: 'top' }); rb(ws2, 3, 4, 4, 4)
    mg(ws2, 3, 5, 4, 5); cl(ws2, 3, 5, data.supplier_company, { size: 9, wrap: true, valign: 'top' }); rb(ws2, 3, 5, 4, 5)
    cl(ws2, 3, 6, 'Destuffing Date :', { bold: true, size: 9 }); ab(ws2, 3, 6)
    mg(ws2, 3, 7, 3, 8); cl(ws2, 3, 7, data.destuffing_date, { size: 9 }); rb(ws2, 3, 7, 3, 8)
    cl(ws2, 4, 1, 'Customer Company :', { bold: true, size: 9, wrap: true }); ab(ws2, 4, 1)
    mg(ws2, 4, 2, 4, 3); cl(ws2, 4, 2, 'RETAIL BUSINESS SOLUTION CO., LTD', { size: 9, wrap: true }); rb(ws2, 4, 2, 4, 3)
    cl(ws2, 4, 6, 'Issue Found Date :', { bold: true, size: 9 }); ab(ws2, 4, 6)
    mg(ws2, 4, 7, 4, 8); cl(ws2, 4, 7, data.issue_found_date, { size: 9 }); rb(ws2, 4, 7, 4, 8)
    mg(ws2, 5, 1, 6, 1); cl(ws2, 5, 1, 'Address :', { bold: true, size: 9, valign: 'top' }); rb(ws2, 5, 1, 6, 1)
    mg(ws2, 5, 2, 6, 3); cl(ws2, 5, 2, '387 SUKHONTHASAWAT RD., LADPRAO, LADPRAO, BANGKOK, THAILAND 10230', { size: 9, wrap: true, valign: 'top' }); rb(ws2, 5, 2, 6, 3)
    cl(ws2, 5, 4, 'Invoice :', { bold: true, size: 9 }); ab(ws2, 5, 4)
    cl(ws2, 5, 5, data.invoice_no, { size: 9 }); ab(ws2, 5, 5)
    mg(ws2, 5, 6, 6, 6); cl(ws2, 5, 6, 'Attachment :', { bold: true, size: 9, valign: 'top' }); rb(ws2, 5, 6, 6, 6)
    mg(ws2, 5, 7, 6, 8); cl(ws2, 5, 7, data.attachment_desc || 'PO, Photos', { size: 9, valign: 'top' }); rb(ws2, 5, 7, 6, 8)
    cl(ws2, 6, 4, 'PO No. :', { bold: true, size: 9 }); ab(ws2, 6, 4)
    cl(ws2, 6, 5, data.po_no, { size: 9 }); ab(ws2, 6, 5)

    // ── Part 5 header ──
    ws2.getRow(7).height = 5
    ws2.getRow(8).height = 13
    mg(ws2, 8, 1, 8, 8); cl(ws2, 8, 1, 'Part 5 : PHOTO', { bold: true, size: 9, bg: 'FFD9D9D9' }); rb(ws2, 8, 1, 8, 8)

    // ── Photo grid: 3 columns × N rows ──────────────────────────────
    function charToCol(chars: number): number {
      let cum = 0
      for (let i = 0; i < COL_WIDTHS.length; i++) {
        if (chars <= cum + COL_WIDTHS[i]) return i + (chars - cum) / COL_WIDTHS[i]
        cum += COL_WIDTHS[i]
      }
      return COL_WIDTHS.length
    }

    const totalW = COL_WIDTHS.reduce((a, b) => a + b, 0)  // 98
    const margin = 1.5, gap = 2
    const photoW = (totalW - margin * 2 - gap * 2) / 3   // ~30.33 chars per photo
    const photoLeftChars  = [margin, margin + photoW + gap, margin + (photoW + gap) * 2]
    const photoRightChars = photoLeftChars.map(x => x + photoW)

    // Photo column width in points: 1 char ≈ 7px at 96 DPI × 0.75 pt/px = 5.25 pt/char
    const photoWidthPt = photoW * 5.25

    // Pre-load photos; use createImageBitmap for accurate display dimensions (incl. EXIF rotation)
    const photoEntries: Array<{ imgId: number; ratio: number }> = []
    for (const url of photos) {
      try {
        const res = await fetch(url)
        const buf = await res.arrayBuffer()
        const imgExt: 'jpeg' | 'png' | 'gif' = url.toLowerCase().includes('.png') ? 'png'
          : url.toLowerCase().includes('.gif') ? 'gif' : 'jpeg'
        const mimeType = imgExt === 'png' ? 'image/png' : imgExt === 'gif' ? 'image/gif' : 'image/jpeg'

        let ratio = 3 / 4  // fallback: landscape 4:3
        try {
          const bmp = await createImageBitmap(new Blob([buf], { type: mimeType }))
          ratio = bmp.height / bmp.width
          bmp.close()
        } catch { /* keep fallback */ }

        const imgId = wb.addImage({ buffer: buf, extension: imgExt })
        photoEntries.push({ imgId, ratio })
      } catch { /* skip failed photos */ }
    }

    // Place in 3-col grid. rowHPt is computed so the bounding box for the tallest photo
    // has the correct aspect ratio: PHOTO_ROWS × rowHPt = photoWidthPt × maxRatio
    const PHOTO_ROWS = 14
    let curExcelRow = 9

    for (let gi = 0; gi < Math.ceil(photoEntries.length / 3); gi++) {
      const batch = photoEntries.slice(gi * 3, (gi + 1) * 3)
      const maxRatio = Math.max(...batch.map(p => p.ratio))  // h/w of tallest photo

      // Row height that makes the tallest photo's bounding box match its aspect ratio
      const rowHPt = Math.max(6, photoWidthPt * maxRatio / PHOTO_ROWS)
      for (let pr = 0; pr < PHOTO_ROWS; pr++) ws2.getRow(curExcelRow + pr).height = rowHPt

      batch.forEach((p, pi) => {
        // br.row scaled so this photo's box matches its own aspect ratio
        const rowFraction = PHOTO_ROWS * (p.ratio / maxRatio)
        ws2.addImage(p.imgId, {
          tl: { col: charToCol(photoLeftChars[pi]),  row: curExcelRow - 1 },
          br: { col: charToCol(photoRightChars[pi]), row: curExcelRow - 1 + rowFraction },
          editAs: 'oneCell',
        })
      })

      curExcelRow += PHOTO_ROWS
      if (gi < Math.ceil(photoEntries.length / 3) - 1) {
        ws2.getRow(curExcelRow).height = 6
        curExcelRow++
      }
    }
  }

  // ── Export ────────────────────────────────────────────────────────
  const buf = await wb.xlsx.writeBuffer()
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url; a.download = `QC-Report-${data.report_no}.xlsx`; a.click()
  URL.revokeObjectURL(url)
}

import type { CanvasElement, Content, TableCell, TDocumentDefinitions } from 'pdfmake/interfaces'
import type { QCReportData } from './qc-excel'

const FONT_FAMILY = 'Sarabun'

const BANNER_GREY = '#E8E8E8'
const TABLE_HEADER_GREY = '#D9D9D9'
const BLACK_BORDER = '#000000'
const PHOTO_BORDER = '#999999'
const MM = 2.83465 // pt per mm

function arrayBufferToBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf)
  let binary = ''
  const chunkSize = 0x8000
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize))
  }
  return btoa(binary)
}

async function fetchBase64(url: string): Promise<string | null> {
  try {
    const res = await fetch(url)
    const buf = await res.arrayBuffer()
    return arrayBufferToBase64(buf)
  } catch { return null }
}

// Thin solid borders on every cell — used for the info block and item table.
const blackBorderLayout = {
  hLineWidth: () => 0.5, vLineWidth: () => 0.5,
  hLineColor: () => BLACK_BORDER, vLineColor: () => BLACK_BORDER,
  paddingLeft: () => 5, paddingRight: () => 5, paddingTop: () => 2, paddingBottom: () => 2,
}
// Photos get a lighter grey frame instead of black.
const photoGridLayout = {
  hLineWidth: () => 0.5, vLineWidth: () => 0.5,
  hLineColor: () => PHOTO_BORDER, vLineColor: () => PHOTO_BORDER,
  paddingLeft: () => 5, paddingRight: () => 5, paddingTop: () => 5, paddingBottom: () => 5,
  dontBreakRows: true,
}

// Full-width light-grey banner bar for "Part N : ..." section titles.
// A plain text node's fillColor only paints the glyph bounding box, not the
// full line — wrapping it in a single-cell borderless table is what actually
// produces a full-width bar.
function sectionBanner(title: string): Content {
  return {
    table: { widths: ['*'], body: [[{ text: title, bold: true, fontSize: 10, fillColor: BANNER_GREY, margin: [5, 4, 5, 4] }]] },
    layout: 'noBorders',
  }
}

// Standard PDF fonts have no glyph for ☐/☑ — both render as an identical
// "missing glyph" box, making checked/unchecked indistinguishable. Draw the
// checkbox as vector shapes instead, so it never depends on font glyph coverage.
function checkbox(checked: boolean, color = '#222222'): Content {
  const size = 9
  const shapes: CanvasElement[] = [
    { type: 'rect', x: 0, y: 0, w: size, h: size, lineWidth: 1, lineColor: color },
  ]
  if (checked) {
    shapes.push(
      { type: 'line', x1: 1.5, y1: 5, x2: 3.6, y2: 7.5, lineWidth: 1.4, lineColor: color },
      { type: 'line', x1: 3.6, y1: 7.5, x2: 7.7, y2: 1.2, lineWidth: 1.4, lineColor: color },
    )
  }
  return { canvas: shapes }
}

function checkboxLabel(checked: boolean, label: string, opts?: { fontSize?: number; bold?: boolean; color?: string }) {
  return {
    columns: [
      { width: 11, margin: [0, 1, 0, 0] as [number, number, number, number], stack: [checkbox(checked, opts?.color)] },
      { width: 'auto' as const, text: label, fontSize: opts?.fontSize ?? 8, bold: opts?.bold, color: opts?.color },
    ],
    columnGap: 4,
  }
}

// Photo number badge: a small dark box with a white number, pulled up with a
// negative top margin so it overlaps the bottom-left corner of the image
// directly above it in the stack (pdfmake has no true overlay/absolute
// positioning within a flowing container, so this is the standard trick).
const BADGE_DARK = '#1a1a1a'
function photoBadge(n: number): Content {
  return {
    // Fixed width (not 'auto') — pdfmake's auto-width sizing can under-measure
    // short numeric strings and wrap "10"/"11"/"12" onto two lines otherwise.
    table: { widths: [20], body: [[{ text: String(n), color: '#ffffff', bold: true, fontSize: 8, alignment: 'center', fillColor: BADGE_DARK, margin: [0, 2, 0, 2] }]] },
    // A "0-width" PDF border still renders as a hairline at device resolution
    // (that's how the spec defines it) — pdfmake's built-in 'noBorders' layout
    // relies on 0-width and can leave a faint seam under the badge. Coloring
    // the (unavoidable) hairline the same as the fill makes it invisible
    // instead of fighting the renderer for a border that's truly absent.
    layout: { hLineWidth: () => 0, vLineWidth: () => 0, hLineColor: () => BADGE_DARK, vLineColor: () => BADGE_DARK, paddingLeft: () => 0, paddingRight: () => 0, paddingTop: () => 0, paddingBottom: () => 0 },
    margin: [0, -17, 0, 0],
  }
}

function sanitizeFilename(name: string): string {
  return (name || 'QC-Report').replace(/[/\\]/g, '-')
}

interface PdfMakeModule {
  createPdf: (doc: TDocumentDefinitions) => { download: (name?: string) => Promise<void> }
  addVirtualFileSystem: (vfs: Record<string, string>) => void
  addFonts: (fonts: Record<string, Record<string, string>>) => void
}

export async function exportQCReportPDF(data: QCReportData) {
  const pdfMakeModule = (await import('pdfmake/build/pdfmake')) as unknown as { default: PdfMakeModule }
  const pdfMake = pdfMakeModule.default

  const [fontRegular, fontBold, logoB64] = await Promise.all([
    fetch('/fonts/Sarabun-Regular.ttf').then(r => r.arrayBuffer()).then(arrayBufferToBase64),
    fetch('/fonts/Sarabun-Bold.ttf').then(r => r.arrayBuffer()).then(arrayBufferToBase64),
    fetchBase64('/rbs-logo.png'),
  ])

  pdfMake.addVirtualFileSystem({
    'Sarabun-Regular.ttf': fontRegular,
    'Sarabun-Bold.ttf': fontBold,
  })
  pdfMake.addFonts({
    Sarabun: {
      normal: 'Sarabun-Regular.ttf', bold: 'Sarabun-Bold.ttf',
      italics: 'Sarabun-Regular.ttf', bolditalics: 'Sarabun-Bold.ttf',
    },
  })

  // ── Header — repeats on every page: logo | title | Report No. box ──────
  // Logo is registered once via docDefinition.images and referenced by name
  // below, so pdfmake embeds it a single time and reuses it across pages.
  const header: Content = {
    margin: [20 * MM, 12, 20 * MM, 6],
    columns: [
      logoB64
        ? { image: 'logo', width: 46, height: 46 }
        : { text: 'rbs', bold: true, fontSize: 18, color: '#004080', width: 46 },
      { text: 'QUALITY CLAIM REPORT', bold: true, fontSize: 17, alignment: 'center', margin: [0, 14, 0, 0] },
      {
        table: {
          widths: ['*'],
          body: [
            [{ text: 'Report No.', bold: true, fontSize: 8, alignment: 'center', fillColor: TABLE_HEADER_GREY }],
            [{ text: data.report_no || '—', bold: true, fontSize: 11, alignment: 'center' }],
          ],
        },
        layout: blackBorderLayout, width: 110,
      },
    ],
    columnGap: 10,
  }

  // ── Info block — 3-column bordered grid, replicating the Excel template's
  // merge layout: Subject/Customer Company/Address down the left, Supplier
  // company/Invoice/PO No. in the middle, Destuffing/Issue dates + Attachment
  // on the right.
  const addr = '387 SUKHONTHASAWAT RD., LADPRAO, LADPRAO, BANGKOK, THAILAND 10230'
  const lbl = (t: string) => ({ text: t, bold: true, fontSize: 7, fillColor: '#F2F2F2' })
  const val = (t: string) => ({ text: t || '—', fontSize: 8 })
  const span: TableCell = {}
  const metaTable: Content = {
    table: {
      widths: [56, 136, 56, 108, 56, '*'],
      body: [
        [lbl('Subject :'), val('QUALITY CLAIM'), { ...lbl('Supplier company :'), rowSpan: 2 }, { ...val(data.supplier_company), rowSpan: 2 }, lbl('Destuffing Date :'), val(data.destuffing_date)],
        [lbl('Customer Company :'), val('RETAIL BUSINESS SOLUTION CO., LTD'), span, span, lbl('Issue Found Date :'), val(data.issue_found_date)],
        [{ ...lbl('Address :'), rowSpan: 2 }, { ...val(addr), rowSpan: 2 }, lbl('Invoice :'), val(data.invoice_no), { ...lbl('Attachment :'), rowSpan: 2 }, { ...val(data.attachment_desc || 'PO, Photos'), rowSpan: 2 }],
        [span, span, lbl('PO No. :'), val(data.po_no), span, span],
      ],
    },
    layout: blackBorderLayout,
    margin: [0, 0, 0, 10],
  }

  // ── Part 1: Issue ─────────────────────────────────────────────────────
  const itemHeaders = ['NO.', 'ITEM CODE', 'PRODUCT DESCRIPTION', 'QTY\n(PCS)', 'UNIT PRICE', 'TOTAL', 'QTY\nDEFECTIVE', 'REMARK']
  let totalQty = 0, totalAmt = 0, totalDef = 0
  const itemRows: TableCell[][] = data.items.map((item, i): TableCell[] => {
    totalQty += item.qty || 0; totalAmt += item.total || 0; totalDef += item.qty_defective || 0
    return [
      { text: String(i + 1), fontSize: 7, alignment: 'center' },
      { text: item.item_code || '', fontSize: 7 },
      { text: (item.product_description || ''), fontSize: 7 },
      { text: String(item.qty || 0), fontSize: 7, alignment: 'right' },
      { text: item.unit_price != null ? item.unit_price.toLocaleString() : '', fontSize: 7, alignment: 'right' },
      { text: item.total != null ? item.total.toLocaleString() : '', fontSize: 7, alignment: 'right' },
      { text: String(item.qty_defective || 0), fontSize: 7, alignment: 'right' },
      { text: (item.remark || ''), fontSize: 7 },
    ]
  })
  const noItemsRow: TableCell[] = [{ text: 'No items', fontSize: 7, colSpan: 8, alignment: 'center' }, {}, {}, {}, {}, {}, {}, {}]
  const totalsRow: TableCell[] = [
    { text: 'Total', bold: true, fontSize: 7, alignment: 'center', fillColor: TABLE_HEADER_GREY, colSpan: 3 }, {}, {},
    { text: totalQty.toLocaleString(), bold: true, fontSize: 7, alignment: 'right', fillColor: TABLE_HEADER_GREY },
    { text: '', fillColor: TABLE_HEADER_GREY },
    { text: totalAmt.toLocaleString(undefined, { minimumFractionDigits: 2 }), bold: true, fontSize: 7, alignment: 'right', fillColor: TABLE_HEADER_GREY },
    { text: totalDef.toLocaleString(), bold: true, fontSize: 7, alignment: 'right', fillColor: TABLE_HEADER_GREY },
    { text: '', fillColor: TABLE_HEADER_GREY },
  ]
  const headerRowCells: TableCell[] = itemHeaders.map(h => ({ text: h, bold: true, fontSize: 6, alignment: 'center', fillColor: TABLE_HEADER_GREY }))
  const itemTableBody: TableCell[][] = [headerRowCells, ...(itemRows.length > 0 ? itemRows : [noItemsRow]), totalsRow]

  const part1: Content[] = [
    sectionBanner('Part 1 : ISSUE'),
    { text: 'Description:', bold: true, fontSize: 8, margin: [4, 4, 4, 1] },
    { text: (data.description || '—'), fontSize: 8, margin: [4, 0, 4, 4] },
    {
      table: {
        headerRows: 1,
        widths: [20, 65, '*', 32, 42, 42, 32, 60],
        body: itemTableBody,
      },
      layout: blackBorderLayout,
    },
  ]

  // ── Part 2: Corrective action ───────────────────────────────────────────
  const caOpts = ['REPLACEMENT IN NEXT SHIPMENT', 'CREDIT NOTE / REFUND', 'REWORK / REPAIR', 'OTHER']
  const part2: Content[] = [
    { text: '', margin: [0, 3, 0, 0] },
    sectionBanner('Part 2 : CORRECTIVE ACTION'),
    {
      columns: [
        {
          width: '50%',
          stack: caOpts.map(opt => ({ ...checkboxLabel(data.corrective_actions.includes(opt), opt), margin: [4, 2, 4, 0] as [number, number, number, number] })),
        },
        {
          width: '50%',
          stack: [
            { text: 'Description / Comment:', bold: true, fontSize: 8, margin: [4, 0, 4, 1] },
            { text: (data.corrective_action_comment || '—'), fontSize: 8, margin: [4, 0, 4, 0] },
          ],
        },
      ],
      columnGap: 10, margin: [0, 4, 0, 0],
    },
  ]

  // ── Part 3: Preventive action ────────────────────────────────────────────
  const part3: Content[] = [
    { text: '', margin: [0, 3, 0, 0] },
    sectionBanner('Part 3 : PREVENTIVE ACTION (FILLED BY SUPPLIER)'),
    {
      columns: [
        { width: '50%', stack: [{ text: 'Root Cause:', bold: true, fontSize: 8, margin: [4, 3, 4, 1] }, { text: (data.root_cause || '—'), fontSize: 8, margin: [4, 0, 4, 0] }] },
        { width: '50%', stack: [{ text: 'Action:', bold: true, fontSize: 8, margin: [4, 3, 4, 1] }, { text: (data.preventive_action || '—'), fontSize: 8, margin: [4, 0, 4, 0] }] },
      ],
      columnGap: 10,
    },
  ]

  // ── Part 4: Verification + signature ────────────────────────────────────
  // Signature sits directly below Part 4 in the normal content flow (not
  // pinned to the page bottom), with generous blank space above each printed
  // name for an actual pen signature.
  const part4: Content[] = [
    { text: '', margin: [0, 3, 0, 0] },
    sectionBanner('Part 4 : VERIFICATION STATUS (FILLED BY RBS)'),
    {
      columns: [
        { ...checkboxLabel(data.verification_accepted === true, 'ACCEPTED', { fontSize: 8.5, bold: true, color: '#1a7a3a' }), width: 'auto', margin: [4, 4, 20, 0] },
        { ...checkboxLabel(data.verification_accepted === false, 'NOT ACCEPTED', { fontSize: 8.5, bold: true, color: '#b91c1c' }), width: 'auto', margin: [0, 4, 20, 0] },
        { text: `Comment: ${data.verification_comment ? (data.verification_comment) : ''}`, fontSize: 8, width: '*', margin: [0, 4, 4, 0] },
      ],
    },
    {
      unbreakable: true,
      margin: [0, 14, 0, 0],
      columns: [
        { width: '50%', alignment: 'center', margin: [4, 0, 4, 0], stack: [{ text: 'FOLLOW-UP INSPECTOR', bold: true, fontSize: 8 }, { text: '(Mr. Weerapong Choungkrai)', fontSize: 8, margin: [0, 42, 0, 0] }, { text: 'DATE…………/…………/…………', fontSize: 8, margin: [0, 3, 0, 0] }] },
        { width: '50%', alignment: 'center', margin: [4, 0, 4, 0], stack: [{ text: 'FOLLOW-UP APPROVAL', bold: true, fontSize: 8 }, { text: '(Mr. Noppharat Sriwichai)', fontSize: 8, margin: [0, 42, 0, 0] }, { text: 'DATE…………/…………/…………', fontSize: 8, margin: [0, 3, 0, 0] }] },
      ],
    },
  ]

  // ── Part 5: Photos — forced onto page 2+, 3 per row, numbered, auto-paginated ──
  const content: Content[] = [metaTable, ...part1, ...part2, ...part3, ...part4]

  const photos = data.photo_urls || []
  if (photos.length > 0) {
    const photoB64s = await Promise.all(photos.map(fetchBase64))
    const cellW = 150
    const emptyCell: TableCell = { text: '', border: [false, false, false, false] }
    const photoCells: TableCell[] = photoB64s.map((b64, i): TableCell => {
      if (!b64) return emptyCell
      const isPng = photos[i].toLowerCase().includes('.png')
      return {
        stack: [
          { image: `data:${isPng ? 'image/png' : 'image/jpeg'};base64,${b64}`, fit: [cellW, 180], alignment: 'left' },
          photoBadge(i + 1),
        ],
      }
    })
    const rows: TableCell[][] = []
    for (let i = 0; i < photoCells.length; i += 3) {
      rows.push([photoCells[i], photoCells[i + 1] ?? emptyCell, photoCells[i + 2] ?? emptyCell])
    }
    content.push(
      { text: '', pageBreak: 'before' },
      sectionBanner('Part 5 : PHOTO'),
      { table: { widths: ['*', '*', '*'], body: rows }, layout: photoGridLayout, margin: [0, 6, 0, 0] },
    )
  }

  const sideMargin = 18 * MM
  const docDefinition: TDocumentDefinitions = {
    pageSize: 'A4',
    pageMargins: [sideMargin, 72, sideMargin, 34],
    defaultStyle: { font: FONT_FAMILY, fontSize: 8 },
    images: logoB64 ? { logo: `data:image/png;base64,${logoB64}` } : undefined,
    header,
    content,
    footer: (currentPage, pageCount) => ({ text: `${data.report_no} · ${currentPage} / ${pageCount}`, alignment: 'center', fontSize: 7, color: '#999999', margin: [0, 6, 0, 0] }),
  }

  const pdf = pdfMake.createPdf(docDefinition)
  await pdf.download(`${sanitizeFilename(data.report_no)}.pdf`)
}

import { NextRequest, NextResponse } from 'next/server'
import PptxGenJS from 'pptxgenjs'

interface QCSlideData {
  report_no: string
  supplier_company: string | null
  issue_found_date: string | null
  issue_types: string[] | null
  description: string
  corrective_actions: string[]
  corrective_action_comment: string
  root_cause: string
  preventive_action: string
  photo_urls: string[]
}

async function imgB64(url: string): Promise<{ data: string; ext: string } | null> {
  try {
    const res = await fetch(url, { signal: AbortSignal.timeout(8000) })
    const buf = await res.arrayBuffer()
    const ext = url.toLowerCase().includes('.png') ? 'png' : 'jpg'
    return { data: `data:image/${ext};base64,${Buffer.from(buf).toString('base64')}`, ext }
  } catch { return null }
}

export async function POST(request: NextRequest) {
  const { reports, fileName } = await request.json() as { reports: QCSlideData[]; fileName: string }

  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'

  const BG = 'F5F0E8', BROWN = '5C3D2E', DARK = '2C1A0E', LINE = 'C4A882', LABEL = '7A5C40'

  // Cover
  const cover = pptx.addSlide()
  cover.background = { color: '1E3340' }
  cover.addText('QC Summary Report', { x: 1, y: 2.4, w: 11.33, h: 0.9, fontSize: 36, bold: true, color: 'D4962A', align: 'center' })
  cover.addText(`รายงาน Quality Claim · ${reports.length} รายการ`, { x: 1, y: 3.4, w: 11.33, h: 0.5, fontSize: 15, color: '7A9AAA', align: 'center' })
  cover.addText(new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' }),
    { x: 1, y: 4.0, w: 11.33, h: 0.4, fontSize: 12, color: '4A6A7A', align: 'center' })

  for (let idx = 0; idx < reports.length; idx++) {
    const r = reports[idx]
    const slide = pptx.addSlide()
    slide.background = { color: BG }

    slide.addText(`ปัญหาที่ ${idx + 1}`, { x: 0.35, y: 0.18, w: 5, h: 0.48, fontSize: 24, bold: true, color: BROWN })
    slide.addShape(pptx.ShapeType.line, { x: 0.35, y: 0.76, w: 12.63, h: 0, line: { color: LINE, width: 1.5 } })

    const LX = 0.35, LW = 1.6, VX = 2.05, VW = 5.5
    const caText = [...r.corrective_actions, r.corrective_action_comment].filter(Boolean).join('\n')
    const prevText = [r.root_cause, r.preventive_action].filter(Boolean).join('\n')

    const rows: { label: string; value: string }[] = [
      { label: 'เลขที่เอกสาร :', value: r.report_no },
      { label: 'ประเภทปัญหา :', value: (r.issue_types ?? []).join(', ') || '—' },
      { label: 'รายละเอียด :', value: r.description || '—' },
      { label: 'การแก้ไข :', value: caText || '—' },
      ...(prevText ? [{ label: 'Preventive :', value: prevText }] : []),
    ]

    let curY = 0.92
    for (const row of rows) {
      const lines = Math.max(1, Math.ceil(row.value.length / 55) + (row.value.match(/\n/g)?.length ?? 0))
      const h = Math.max(0.32, lines * 0.21)
      slide.addText(row.label, { x: LX, y: curY, w: LW, h, fontSize: 12, bold: true, color: LABEL, valign: 'top' })
      slide.addText(row.value, { x: VX, y: curY, w: VW, h, fontSize: 12, color: DARK, wrap: true, valign: 'top' })
      curY += h + 0.1
    }

    // Photos
    const photos = (r.photo_urls || []).slice(0, 6)
    if (photos.length > 0) {
      const PX = 7.7, PW = 5.3, PY = 0.82, PH = 6.4
      const COLS = 2, GAP = 0.1
      const ROWS = Math.ceil(photos.length / COLS)
      const cW = (PW - GAP * (COLS - 1)) / COLS
      const cH = (PH - GAP * (ROWS - 1)) / ROWS

      await Promise.all(photos.map(async (url, pi) => {
        const img = await imgB64(url)
        if (!img) return
        const col = pi % COLS, row2 = Math.floor(pi / COLS)
        slide.addImage({
          data: img.data,
          x: PX + col * (cW + GAP), y: PY + row2 * (cH + GAP),
          w: cW, h: cH,
          sizing: { type: 'contain', w: cW, h: cH },
        })
      }))
    }
  }

  const buffer = await pptx.write('nodebuffer') as unknown as Buffer

  return new NextResponse(buffer, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-Disposition': `attachment; filename="${fileName || 'QC-Report'}-${new Date().toISOString().slice(0, 10)}.pptx"`,
    },
  })
}

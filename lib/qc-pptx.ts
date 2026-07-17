import PptxGenJS from 'pptxgenjs'

export interface QCSlideData {
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

function toBase64(buffer: ArrayBuffer): string {
  let binary = ''
  const bytes = new Uint8Array(buffer)
  for (let i = 0; i < bytes.length; i += 8192) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 8192))
  }
  return btoa(binary)
}

export async function exportQCSummaryPptx(reports: QCSlideData[], fileName = 'QC-Summary') {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'  // 13.33" × 7.5"

  const BG     = 'F5F0E8'
  const BROWN  = '5C3D2E'
  const DARK   = '2C1A0E'
  const LINE   = 'C4A882'
  const LABEL_COLOR = '7A5C40'

  // ── Cover slide ──────────────────────────────────────────────────
  const cover = pptx.addSlide()
  cover.background = { color: '1E3340' }
  cover.addText('QC Summary Report', {
    x: 1, y: 2.4, w: 11.33, h: 0.9,
    fontSize: 36, bold: true, color: 'D4962A', align: 'center',
  })
  cover.addText(`รายงาน Quality Claim · ${reports.length} รายการ`, {
    x: 1, y: 3.4, w: 11.33, h: 0.5,
    fontSize: 15, color: '7A9AAA', align: 'center',
  })
  const now = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })
  cover.addText(now, {
    x: 1, y: 4.0, w: 11.33, h: 0.4,
    fontSize: 12, color: '4A6A7A', align: 'center',
  })

  // ── One slide per report ──────────────────────────────────────────
  for (let idx = 0; idx < reports.length; idx++) {
    const r = reports[idx]
    const slide = pptx.addSlide()
    slide.background = { color: BG }

    // Title
    slide.addText(`ปัญหาที่ ${idx + 1}`, {
      x: 0.35, y: 0.18, w: 5, h: 0.48,
      fontSize: 24, bold: true, color: BROWN,
    })
    // Divider line
    slide.addShape(pptx.ShapeType.line, {
      x: 0.35, y: 0.76, w: 12.63, h: 0,
      line: { color: LINE, width: 1.5 },
    })

    // ── Left text content ──
    const LX = 0.35
    const LABEL_W = 1.6
    const VAL_X = LX + LABEL_W + 0.05
    const VAL_W = 5.6

    const caText = [
      ...r.corrective_actions,
      r.corrective_action_comment,
    ].filter(Boolean).join('\n')

    const prevText = [r.root_cause, r.preventive_action].filter(Boolean).join('\n')

    const rows: { label: string; value: string; large?: boolean }[] = [
      { label: 'เลขที่เอกสาร :', value: r.report_no },
      { label: 'ประเภทปัญหา :', value: (r.issue_types ?? []).join(', ') || '—' },
      { label: 'รายละเอียด :', value: r.description || '—', large: true },
      { label: 'การแก้ไข :', value: caText || '—', large: true },
      ...(prevText ? [{ label: 'Preventive :', value: prevText, large: true }] : []),
    ]

    let curY = 0.92
    for (const row of rows) {
      const lineCount = row.value.split('\n').length + Math.ceil(row.value.length / 60)
      const h = row.large ? Math.max(0.35, lineCount * 0.2) : 0.35
      slide.addText(row.label, {
        x: LX, y: curY, w: LABEL_W, h,
        fontSize: 12, bold: true, color: LABEL_COLOR, valign: 'top',
      })
      slide.addText(row.value, {
        x: VAL_X, y: curY, w: VAL_W, h,
        fontSize: 12, color: DARK, wrap: true, valign: 'top',
      })
      curY += h + 0.1
    }

    // ── Right photos ──
    const photos = (r.photo_urls || []).slice(0, 6)
    if (photos.length > 0) {
      const PX = 7.6
      const PW = 5.4
      const PY = 0.82
      const PH = 6.4
      const COLS = 2
      const GAP = 0.1
      const ROWS = Math.ceil(photos.length / COLS)
      const cellW = (PW - GAP * (COLS - 1)) / COLS
      const cellH = (PH - GAP * (ROWS - 1)) / ROWS

      for (let pi = 0; pi < photos.length; pi++) {
        const col = pi % COLS
        const row2 = Math.floor(pi / COLS)
        const imgX = PX + col * (cellW + GAP)
        const imgY = PY + row2 * (cellH + GAP)
        try {
          const res = await fetch(photos[pi])
          const buf = await res.arrayBuffer()
          const b64 = toBase64(buf)
          const ext = photos[pi].toLowerCase().includes('.png') ? 'png' : 'jpg'
          slide.addImage({
            data: `data:image/${ext};base64,${b64}`,
            x: imgX, y: imgY, w: cellW, h: cellH,
            sizing: { type: 'contain', w: cellW, h: cellH },
          })
        } catch { /* skip failed image */ }
      }
    }
  }

  await pptx.writeFile({ fileName: `${fileName}-${new Date().toISOString().slice(0, 10)}.pptx` })
}

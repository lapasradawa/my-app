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

export async function exportQCSummaryPptx(reports: QCSlideData[], fileName = 'QC-Report') {
  const res = await fetch('/api/qc-pptx', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ reports, fileName }),
  })
  if (!res.ok) throw new Error('Failed to generate PPTX')
  const blob = await res.blob()
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `${fileName}-${new Date().toISOString().slice(0, 10)}.pptx`
  a.click()
  URL.revokeObjectURL(url)
}

import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import type { ResultRow } from './excel-parser'

export function exportToExcel(rows: ResultRow[], containerNames: string[], filename = 'PO-Matching-Result.xlsx') {
  const headers = ['No.', 'Code', 'Description', 'PO', 'QTY', ...containerNames, 'LEFT']

  const data = rows.map(row => [
    row.no,
    row.code,
    row.description,
    row.po,
    row.qty,
    ...containerNames.map(name => row.containers[name] || 0),
    row.left,
  ])

  const ws = XLSX.utils.aoa_to_sheet([headers, ...data])

  // Column widths
  ws['!cols'] = [
    { wch: 5 },   // N
    { wch: 25 },  // Code
    { wch: 55 },  // Description
    { wch: 35 },  // PO
    { wch: 10 },  // QTY
    ...containerNames.map(() => ({ wch: 15 })),
    { wch: 10 },  // LEFT
  ]

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'PO Matching')

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
  saveAs(new Blob([wbout], { type: 'application/octet-stream' }), filename)
}

'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'
import * as XLSX from 'xlsx'

interface InvRow {
  invoice_no: string
  estimated_arrival: string | null
  total_amount: number | null
  currency: string | null
  exchange_rate: number | null
  cost_saving: number | null
  cost_saving_pct: number | null
}

function monthLabel(d: string): string {
  return new Date(d + 'T00:00:00').toLocaleDateString('en-US', { month: 'long', year: 'numeric' })
}

function yearOf(d: string): number {
  return new Date(d + 'T00:00:00').getFullYear()
}

function compute(inv: InvRow) {
  const fobCny = inv.currency === 'CNY' ? (inv.total_amount ?? null) : null
  const fobUsd = inv.currency === 'USD' ? (inv.total_amount ?? null) : null
  const actualThb =
    inv.total_amount != null && inv.exchange_rate != null
      ? inv.total_amount * inv.exchange_rate
      : null
  const pct5 = actualThb != null ? actualThb * 0.05 : null
  return { fobCny, fobUsd, actualThb, pct5 }
}

function sumN(nums: (number | null)[]): number | null {
  const valid = nums.filter((x): x is number => x != null)
  return valid.length > 0 ? valid.reduce((a, b) => a + b, 0) : null
}

function fmt(n: number | null | undefined, dec = 2): string {
  if (n == null || isNaN(n)) return ''
  return n.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec })
}

function Cell({ v, gray }: { v: string; gray?: boolean }) {
  return (
    <td className={`px-3 py-2 border border-gray-200 text-right text-gray-700 ${gray ? 'bg-gray-100' : ''}`}>
      {v || <span className="text-gray-300">—</span>}
    </td>
  )
}

export default function ReportPage() {
  const [rows, setRows] = useState<InvRow[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => { load() }, [])

  async function load() {
    const { data } = await supabase
      .from('invoices')
      .select('invoice_no, estimated_arrival, total_amount, currency, exchange_rate, cost_saving, cost_saving_pct')
      .order('estimated_arrival', { ascending: true, nullsFirst: false })
    setRows((data ?? []) as InvRow[])
    setLoading(false)
  }

  // Build grouped structure: year → months → rows
  const grouped: { year: number; months: { label: string; rows: InvRow[] }[] }[] = []
  for (const inv of rows) {
    if (!inv.estimated_arrival) continue
    const year = yearOf(inv.estimated_arrival)
    const month = monthLabel(inv.estimated_arrival)
    let yg = grouped.find(g => g.year === year)
    if (!yg) { yg = { year, months: [] }; grouped.push(yg) }
    let mg = yg.months.find(m => m.label === month)
    if (!mg) { mg = { label: month, rows: [] }; yg.months.push(mg) }
    mg.rows.push(inv)
  }

  const allWithDate = rows.filter(r => r.estimated_arrival)
  const allC = allWithDate.map(compute)
  const grandTotal = {
    fobCny: sumN(allC.map(c => c.fobCny)),
    fobUsd: sumN(allC.map(c => c.fobUsd)),
    actualThb: sumN(allC.map(c => c.actualThb)),
    pct5: sumN(allC.map(c => c.pct5)),
    costSaving: sumN(allWithDate.map(r => r.cost_saving)),
  }

  function exportExcel() {
    const header = ['เข้าคลัง MONTH', 'Invoice No.', 'FOB CNY', 'FOB USD', 'Actual FOB THB (Finance)', '5% for RBS CH', 'Cost saving (THB)', 'Cost saving (%)']
    const aoa: (string | number | null)[][] = [header]

    for (const yg of grouped) {
      const yRows = yg.months.flatMap(m => m.rows)
      const yC = yRows.map(compute)
      for (const mg of yg.months) {
        const mC = mg.rows.map(compute)
        mg.rows.forEach((inv, i) => {
          const c = mC[i]
          aoa.push([
            i === 0 ? mg.label : '',
            inv.invoice_no,
            c.fobCny, c.fobUsd, c.actualThb, c.pct5,
            inv.cost_saving,
            inv.cost_saving_pct != null ? inv.cost_saving_pct / 100 : null,
          ])
        })
        aoa.push([mg.label + ' Total', '', sumN(mC.map(c => c.fobCny)), sumN(mC.map(c => c.fobUsd)), sumN(mC.map(c => c.actualThb)), sumN(mC.map(c => c.pct5)), sumN(mg.rows.map(r => r.cost_saving)), null])
      }
      aoa.push([String(yg.year) + ' Total', '', sumN(yC.map(c => c.fobCny)), sumN(yC.map(c => c.fobUsd)), sumN(yC.map(c => c.actualThb)), sumN(yC.map(c => c.pct5)), sumN(yRows.map(r => r.cost_saving)), null])
    }

    aoa.push([`GRAND TOTAL`, '', grandTotal.fobCny, grandTotal.fobUsd, grandTotal.actualThb, grandTotal.pct5, grandTotal.costSaving, null])

    const ws = XLSX.utils.aoa_to_sheet(aoa)
    // Number format for numeric columns
    const numFmt = '#,##0.00'
    const numCols = [2, 3, 4, 5, 6]
    const pctCol = 7
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
    for (let r = 1; r <= range.e.r; r++) {
      numCols.forEach(c => {
        const addr = XLSX.utils.encode_cell({ r, c })
        if (ws[addr] && typeof ws[addr].v === 'number') ws[addr].z = numFmt
      })
      const pctAddr = XLSX.utils.encode_cell({ r, c: pctCol })
      if (ws[pctAddr] && typeof ws[pctAddr].v === 'number') ws[pctAddr].z = '0.00%'
    }
    ws['!cols'] = [{ wch: 20 }, { wch: 22 }, { wch: 16 }, { wch: 14 }, { wch: 24 }, { wch: 16 }, { wch: 18 }, { wch: 14 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Report')
    XLSX.writeFile(wb, `Import_Report_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  if (loading) return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center text-gray-400">กำลังโหลด...</div>
  )

  const th = 'px-3 py-2.5 text-right border border-amber-300 whitespace-nowrap font-bold text-xs'

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/report" className="text-sm text-blue-600 font-semibold">Report</Link>
      </nav>

      <div className="max-w-7xl mx-auto px-6 py-8">
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">Report</h1>
            <p className="text-sm text-gray-500 mt-1">สรุปยอด FOB และ Cost Saving รายเดือน</p>
          </div>
          <button
            onClick={exportExcel}
            className="px-4 py-2 bg-green-600 text-white text-sm font-medium rounded-lg hover:bg-green-700 transition-colors"
          >
            Export Excel
          </button>
        </div>

        <div className="bg-white rounded-xl border border-gray-200 overflow-auto shadow-sm">
          <table className="text-sm border-collapse w-full">
            <thead>
              <tr className="bg-amber-400 text-gray-900">
                <th className="px-3 py-2.5 text-left border border-amber-300 whitespace-nowrap font-bold text-xs">เข้าคลัง MONTH</th>
                <th className="px-3 py-2.5 text-left border border-amber-300 whitespace-nowrap font-bold text-xs">Invoice No.</th>
                <th className={th}>FOB CNY</th>
                <th className={th}>FOB USD</th>
                <th className={th}>Actual FOB THB (Finance)</th>
                <th className={th}>5% for RBS CH</th>
                <th className={`${th} bg-amber-500`}>Cost saving (THB)</th>
                <th className={`${th} bg-amber-500`}>Cost saving (%)</th>
              </tr>
            </thead>
            <tbody>
              {grouped.flatMap(yg => {
                const yRows = yg.months.flatMap(m => m.rows)
                const yC = yRows.map(compute)
                const monthRows = yg.months.flatMap(mg => {
                  const mC = mg.rows.map(compute)
                  const mFobCny = sumN(mC.map(c => c.fobCny))
                  const mFobUsd = sumN(mC.map(c => c.fobUsd))
                  const mActual = sumN(mC.map(c => c.actualThb))
                  const mPct5 = sumN(mC.map(c => c.pct5))
                  const mCost = sumN(mg.rows.map(r => r.cost_saving))

                  return [
                    ...mg.rows.map((inv, i) => {
                      const c = mC[i]
                      return (
                        <tr key={`${mg.label}-${inv.invoice_no}`} className="border-b border-gray-100 hover:bg-gray-50">
                          <td className="px-3 py-2 border border-gray-200 text-gray-700 whitespace-nowrap">
                            {i === 0 ? mg.label : ''}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 text-blue-600 font-medium whitespace-nowrap">{inv.invoice_no}</td>
                          <Cell v={fmt(c.fobCny)} gray={c.fobCny == null} />
                          <Cell v={fmt(c.fobUsd)} gray={c.fobUsd == null} />
                          <Cell v={fmt(c.actualThb)} gray={c.actualThb == null} />
                          <Cell v={fmt(c.pct5)} gray={c.pct5 == null} />
                          <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">
                            {inv.cost_saving != null ? fmt(inv.cost_saving) : <span className="text-gray-300">—</span>}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">
                            {inv.cost_saving_pct != null ? `${inv.cost_saving_pct}%` : <span className="text-gray-300">—</span>}
                          </td>
                        </tr>
                      )
                    }),
                    <tr key={`${mg.label}-total`} className="bg-amber-50 font-semibold text-gray-800">
                      <td className="px-3 py-2 border border-amber-200 text-amber-900 whitespace-nowrap">{mg.label} Total</td>
                      <td className="px-3 py-2 border border-amber-200"></td>
                      <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mFobCny)}</td>
                      <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mFobUsd)}</td>
                      <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mActual)}</td>
                      <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mPct5)}</td>
                      <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mCost)}</td>
                      <td className="px-3 py-2 border border-amber-200"></td>
                    </tr>,
                  ]
                })
                return [
                  ...monthRows,
                  <tr key={`year-${yg.year}`} className="bg-blue-100 font-bold text-blue-900 border-t-2 border-blue-300">
                    <td className="px-3 py-2 border border-blue-200 whitespace-nowrap">{yg.year} Total</td>
                    <td className="px-3 py-2 border border-blue-200"></td>
                    <td className="px-3 py-2 border border-blue-200 text-right">{fmt(sumN(yC.map(c => c.fobCny)))}</td>
                    <td className="px-3 py-2 border border-blue-200 text-right">{fmt(sumN(yC.map(c => c.fobUsd)))}</td>
                    <td className="px-3 py-2 border border-blue-200 text-right">{fmt(sumN(yC.map(c => c.actualThb)))}</td>
                    <td className="px-3 py-2 border border-blue-200 text-right">{fmt(sumN(yC.map(c => c.pct5)))}</td>
                    <td className="px-3 py-2 border border-blue-200 text-right">{fmt(sumN(yRows.map(r => r.cost_saving)))}</td>
                    <td className="px-3 py-2 border border-blue-200"></td>
                  </tr>,
                ]
              })}
              <tr className="bg-purple-700 text-white font-bold border-t-2 border-purple-800">
                <td className="px-3 py-2.5 border border-purple-600 whitespace-nowrap">
                  GRAND TOTAL {new Date().toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })}
                </td>
                <td className="px-3 py-2.5 border border-purple-600"></td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.fobCny)}</td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.fobUsd)}</td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.actualThb)}</td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.pct5)}</td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.costSaving)}</td>
                <td className="px-3 py-2.5 border border-purple-600"></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

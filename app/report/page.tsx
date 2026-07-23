'use client'

import { useEffect, useState, useMemo } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'
import * as XLSX from 'xlsx'
import { isUnlocked } from '@/lib/auth'
import LockButton from '@/components/LockButton'
import PasswordModal from '@/components/PasswordModal'

interface ExchangeRateEntry { amount: number; rate: number }

interface InvRow {
  id: string
  invoice_no: string
  estimated_arrival: string | null
  total_amount: number | null
  currency: string | null
  exchange_rate: number | null
  exchange_rates: ExchangeRateEntry[] | null
  cost_saving: number | null
  cost_saving_pct: number | null
  bl_date: string | null
  payment_date: string | null
  commission_payment_date: string | null
}

function monthLabel(d: string): string {
  return new Date(d + 'T00:00:00').toLocaleDateString('en-US', { month: 'long', year: 'numeric' })
}

function monthKey(d: string): string {
  const dt = new Date(d + 'T00:00:00')
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`
}

type GroupMode = 'arrival' | 'due' | 'payment'

function getGroupDate(inv: InvRow, mode: GroupMode): string | null {
  if (mode === 'arrival') return inv.estimated_arrival
  if (mode === 'due') {
    if (!inv.bl_date) return null
    const d = new Date(inv.bl_date + 'T00:00:00')
    d.setDate(d.getDate() + 30)
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
  }
  return inv.payment_date
}

function dueDateStr(blDate: string | null): string {
  if (!blDate) return ''
  const d = new Date(blDate + 'T00:00:00')
  d.setDate(d.getDate() + 30)
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
}

function fmtDate(d: string | null): string {
  if (!d) return ''
  return new Date(d + 'T00:00:00').toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
}

function compute(inv: InvRow) {
  const fobCny = inv.currency === 'CNY' ? (inv.total_amount ?? null) : null
  const fobUsd = inv.currency === 'USD' ? (inv.total_amount ?? null) : null
  let actualThb: number | null = null
  if (inv.exchange_rates && inv.exchange_rates.length > 0) {
    actualThb = inv.exchange_rates.reduce((s, e) => s + e.amount * e.rate, 0)
  } else if (inv.total_amount != null && inv.exchange_rate != null) {
    actualThb = inv.total_amount * inv.exchange_rate
  }
  return { fobCny, fobUsd, actualThb }
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
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(new Set())
  const [invoiceSearch, setInvoiceSearch] = useState('')
  const [unlocked, setUnlocked] = useState(false)
  const [showModal, setShowModal] = useState(false)
  const [commEdits, setCommEdits] = useState<Record<string, string>>({})
  const [commSaving, setCommSaving] = useState<Record<string, boolean>>({})
  const [rateDetail, setRateDetail] = useState<InvRow | null>(null)
  const [groupMode, setGroupMode] = useState<GroupMode>('arrival')

  useEffect(() => { load(); setUnlocked(isUnlocked()) }, [])

  async function load() {
    const { data } = await supabase
      .from('invoices')
      .select('id, invoice_no, estimated_arrival, total_amount, currency, exchange_rate, exchange_rates, cost_saving, cost_saving_pct, bl_date, payment_date, commission_payment_date')
      .order('estimated_arrival', { ascending: true, nullsFirst: false })
    const fetched = (data ?? []) as InvRow[]
    setRows(fetched)
    const ce: Record<string, string> = {}
    for (const inv of fetched) ce[inv.id] = inv.commission_payment_date || ''
    setCommEdits(ce)
    setLoading(false)
  }

  async function saveCommission(id: string) {
    const val = commEdits[id] ?? ''
    setCommSaving(s => ({ ...s, [id]: true }))
    await supabase.from('invoices').update({ commission_payment_date: val || null }).eq('id', id)
    setCommSaving(s => ({ ...s, [id]: false }))
  }

  const allMonths = useMemo(() => {
    const map = new Map<string, string>()
    for (const inv of rows) {
      const d = getGroupDate(inv, groupMode)
      if (!d) continue
      const k = monthKey(d)
      if (!map.has(k)) map.set(k, monthLabel(d))
    }
    return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0]))
  }, [rows, groupMode])

  useEffect(() => {
    setSelectedMonths(new Set(allMonths.map(m => m[0])))
  }, [allMonths])

  const grouped = useMemo(() => {
    const result: { key: string; label: string; rows: InvRow[] }[] = []
    for (const inv of rows) {
      const d = getGroupDate(inv, groupMode)
      if (!d) continue
      const k = monthKey(d)
      if (!selectedMonths.has(k)) continue
      const q = invoiceSearch.trim().toLowerCase()
      if (q && !inv.invoice_no.toLowerCase().includes(q)) continue
      let mg = result.find(m => m.key === k)
      if (!mg) {
        mg = { key: k, label: monthLabel(d), rows: [] }
        result.push(mg)
      }
      mg.rows.push(inv)
    }
    return result
  }, [rows, selectedMonths, invoiceSearch, groupMode])

  const allVisible = grouped.flatMap(m => m.rows)
  const allVisC = allVisible.map(compute)
  const grandTotal = {
    fobCny: sumN(allVisC.map(c => c.fobCny)),
    fobUsd: sumN(allVisC.map(c => c.fobUsd)),
    actualThb: sumN(allVisC.map(c => c.actualThb)),
    costSaving: sumN(allVisible.map(r => r.cost_saving)),
  }

  const groupLabel = groupMode === 'arrival' ? 'เข้าคลัง MONTH'
    : groupMode === 'due' ? 'Due Date MONTH'
    : 'Payment MONTH'

  function toggleMonth(k: string) {
    setSelectedMonths(prev => {
      const next = new Set(prev)
      if (next.has(k)) next.delete(k); else next.add(k)
      return next
    })
  }

  function selectAll() { setSelectedMonths(new Set(allMonths.map(m => m[0]))) }
  function clearAll() { setSelectedMonths(new Set()) }

  function exportExcel() {
    const header = [groupLabel, 'Invoice No.', 'FOB CNY', 'FOB USD', 'Actual FOB THB (Finance)', 'Due Date', 'Payment Date', 'Cost saving (THB)', 'Cost saving (%)', 'Commission Payment']
    const aoa: (string | number | null)[][] = [header]

    for (const mg of grouped) {
      const mC = mg.rows.map(compute)
      mg.rows.forEach((inv, i) => {
        const c = mC[i]
        aoa.push([
          i === 0 ? mg.label : '',
          inv.invoice_no,
          c.fobCny, c.fobUsd, c.actualThb,
          dueDateStr(inv.bl_date),
          fmtDate(inv.payment_date),
          inv.cost_saving,
          inv.cost_saving_pct != null ? inv.cost_saving_pct / 100 : null,
          fmtDate(commEdits[inv.id] || inv.commission_payment_date),
        ])
      })
      aoa.push([mg.label + ' Total', '', sumN(mC.map(c => c.fobCny)), sumN(mC.map(c => c.fobUsd)), sumN(mC.map(c => c.actualThb)), '', '', sumN(mg.rows.map(r => r.cost_saving)), null, ''])
    }

    aoa.push(['GRAND TOTAL', '', grandTotal.fobCny, grandTotal.fobUsd, grandTotal.actualThb, '', '', grandTotal.costSaving, null, ''])

    const ws = XLSX.utils.aoa_to_sheet(aoa)
    const numFmt = '#,##0.00'
    const numCols = [2, 3, 4, 7]
    const pctCol = 8
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
    for (let r = 1; r <= range.e.r; r++) {
      numCols.forEach(c => {
        const addr = XLSX.utils.encode_cell({ r, c })
        if (ws[addr] && typeof ws[addr].v === 'number') ws[addr].z = numFmt
      })
      const pctAddr = XLSX.utils.encode_cell({ r, c: pctCol })
      if (ws[pctAddr] && typeof ws[pctAddr].v === 'number') ws[pctAddr].z = '0.00%'
    }
    ws['!cols'] = [{ wch: 20 }, { wch: 22 }, { wch: 16 }, { wch: 14 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 18 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Report')
    XLSX.writeFile(wb, `Import_Report_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  if (loading) return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center text-gray-400">กำลังโหลด...</div>
  )

  const th = 'px-3 py-2.5 text-right border border-amber-300 whitespace-nowrap font-bold text-xs'
  const thC = 'px-3 py-2.5 text-center border border-amber-300 whitespace-nowrap font-bold text-xs'

  return (
    <div className="min-h-screen bg-gray-50">
      {showModal && (
        <PasswordModal
          onSuccess={() => { setUnlocked(true); setShowModal(false) }}
          onCancel={() => setShowModal(false)}
        />
      )}

      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4 sticky top-0 z-20 shadow-sm">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/calendar" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Calendar</Link>
        <Link href="/report" className="text-sm text-blue-600 font-semibold">Report</Link>
        <Link href="/summary" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Summary</Link>
        <Link href="/compare" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Builder</Link>
        <Link href="/qc" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">QC Report</Link>
        <Link href="/guide" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Guide</Link>
        <div className="ml-auto">
          <LockButton onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />
        </div>
      </nav>

      <div className="max-w-full mx-auto px-6 py-8">
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

        {/* Filters */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 mb-4 space-y-3">
          <div className="flex items-center gap-3 flex-wrap">
            <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide whitespace-nowrap">จัดกลุ่มตาม</span>
            <div className="flex rounded-lg overflow-hidden border border-gray-200">
              {(['arrival', 'due', 'payment'] as GroupMode[]).map((mode, i) => (
                <button
                  key={mode}
                  onClick={() => setGroupMode(mode)}
                  className={`px-3 py-1.5 text-xs font-medium transition-colors ${
                    groupMode === mode ? 'bg-amber-400 text-gray-900' : 'bg-white text-gray-500 hover:bg-amber-50'
                  } ${i > 0 ? 'border-l border-gray-200' : ''}`}
                >
                  {mode === 'arrival' ? 'เข้าคลัง' : mode === 'due' ? 'Due Date' : 'Payment'}
                </button>
              ))}
            </div>
          </div>

          <div>
            <div className="flex items-center gap-3 mb-2">
              <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide">เดือน</span>
              <button onClick={selectAll} className="text-xs text-blue-600 hover:underline">เลือกทั้งหมด</button>
              <button onClick={clearAll} className="text-xs text-gray-400 hover:underline">ยกเลิกทั้งหมด</button>
            </div>
            <div className="flex flex-wrap gap-2">
              {allMonths.map(([k, label]) => {
                const active = selectedMonths.has(k)
                return (
                  <button
                    key={k}
                    onClick={() => toggleMonth(k)}
                    className={`px-3 py-1 rounded-full text-xs font-medium border transition-colors ${
                      active
                        ? 'bg-amber-400 border-amber-400 text-gray-900'
                        : 'bg-white border-gray-300 text-gray-500 hover:border-amber-300'
                    }`}
                  >
                    {label}
                  </button>
                )
              })}
            </div>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide whitespace-nowrap">Invoice No.</span>
            <input
              type="text"
              value={invoiceSearch}
              onChange={e => setInvoiceSearch(e.target.value)}
              placeholder="ค้นหา Invoice..."
              className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm text-gray-800 w-64 focus:outline-none focus:ring-2 focus:ring-amber-300"
            />
            {invoiceSearch && (
              <button onClick={() => setInvoiceSearch('')} className="text-xs text-gray-400 hover:text-gray-700">ล้าง</button>
            )}
          </div>
        </div>

        <div className="bg-white rounded-xl border border-gray-200 overflow-auto shadow-sm">
          <table className="text-sm border-collapse w-full">
            <thead>
              <tr className="bg-amber-400 text-gray-900">
                <th className="px-3 py-2.5 text-left border border-amber-300 whitespace-nowrap font-bold text-xs">{groupLabel}</th>
                <th className="px-3 py-2.5 text-left border border-amber-300 whitespace-nowrap font-bold text-xs">Invoice No.</th>
                <th className={th}>FOB CNY</th>
                <th className={th}>FOB USD</th>
                <th className={th}>Actual FOB THB (Finance)</th>
                <th className={thC}>Due Date</th>
                <th className={thC}>Payment Date</th>
                <th className={`${th} bg-amber-500`}>Cost saving (THB)</th>
                <th className={`${th} bg-amber-500`}>Cost saving (%)</th>
                <th className="px-3 py-2.5 text-center border border-amber-300 whitespace-nowrap font-bold text-xs bg-blue-100 text-blue-900">Commission Payment</th>
              </tr>
            </thead>
            <tbody>
              {grouped.flatMap(mg => {
                const mC = mg.rows.map(compute)
                const mFobCny = sumN(mC.map(c => c.fobCny))
                const mFobUsd = sumN(mC.map(c => c.fobUsd))
                const mActual = sumN(mC.map(c => c.actualThb))
                const mCost = sumN(mg.rows.map(r => r.cost_saving))

                return [
                  ...mg.rows.map((inv, i) => {
                    const c = mC[i]
                    const commVal = commEdits[inv.id] ?? ''
                    return (
                      <tr key={`${mg.key}-${inv.invoice_no}`} className="border-b border-gray-100 hover:bg-gray-50">
                        <td className="px-3 py-2 border border-gray-200 text-gray-700 whitespace-nowrap">
                          {i === 0 ? mg.label : ''}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-blue-600 font-medium whitespace-nowrap">{inv.invoice_no}</td>
                        <Cell v={fmt(c.fobCny)} gray={c.fobCny == null} />
                        <Cell v={fmt(c.fobUsd)} gray={c.fobUsd == null} />
                        <td
                          className={`px-3 py-2 border border-gray-200 text-right text-gray-700 ${c.actualThb == null ? 'bg-gray-100' : 'cursor-pointer hover:bg-blue-50 hover:underline'}`}
                          onClick={() => { if (c.actualThb != null) setRateDetail(inv) }}
                        >
                          {fmt(c.actualThb) || <span className="text-gray-300">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-center text-gray-600 whitespace-nowrap text-xs">
                          {dueDateStr(inv.bl_date) || <span className="text-gray-300">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-center text-gray-600 whitespace-nowrap text-xs">
                          {fmtDate(inv.payment_date) || <span className="text-gray-300">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">
                          {inv.cost_saving != null ? fmt(inv.cost_saving) : <span className="text-gray-300">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">
                          {inv.cost_saving_pct != null ? `${inv.cost_saving_pct}%` : <span className="text-gray-300">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-center">
                          {unlocked ? (
                            <div className="flex items-center gap-1 justify-center">
                              <input
                                type="date"
                                value={commVal}
                                onChange={ev => setCommEdits(e => ({ ...e, [inv.id]: ev.target.value }))}
                                onBlur={() => saveCommission(inv.id)}
                                className="text-xs border border-gray-300 rounded px-2 py-1 outline-none focus:border-blue-400 text-gray-700"
                              />
                              {commSaving[inv.id] && <span className="text-xs text-gray-400">...</span>}
                            </div>
                          ) : (
                            <span className="text-xs text-gray-600 whitespace-nowrap">
                              {fmtDate(commVal) || <span className="text-gray-300">—</span>}
                            </span>
                          )}
                        </td>
                      </tr>
                    )
                  }),
                  <tr key={`${mg.key}-total`} className="bg-amber-50 font-semibold text-gray-800">
                    <td className="px-3 py-2 border border-amber-200 text-amber-900 whitespace-nowrap">{mg.label} Total</td>
                    <td className="px-3 py-2 border border-amber-200"></td>
                    <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mFobCny)}</td>
                    <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mFobUsd)}</td>
                    <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mActual)}</td>
                    <td className="px-3 py-2 border border-amber-200"></td>
                    <td className="px-3 py-2 border border-amber-200"></td>
                    <td className="px-3 py-2 border border-amber-200 text-right">{fmt(mCost)}</td>
                    <td className="px-3 py-2 border border-amber-200"></td>
                    <td className="px-3 py-2 border border-amber-200"></td>
                  </tr>,
                ]
              })}
              {grouped.length === 0 && (
                <tr>
                  <td colSpan={10} className="px-4 py-8 text-center text-gray-400 text-sm">ไม่มีข้อมูลที่ตรงกับตัวกรอง</td>
                </tr>
              )}
              <tr className="bg-purple-700 text-white font-bold border-t-2 border-purple-800">
                <td className="px-3 py-2.5 border border-purple-600 whitespace-nowrap">
                  GRAND TOTAL {new Date().toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })}
                </td>
                <td className="px-3 py-2.5 border border-purple-600"></td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.fobCny)}</td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.fobUsd)}</td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.actualThb)}</td>
                <td className="px-3 py-2.5 border border-purple-600"></td>
                <td className="px-3 py-2.5 border border-purple-600"></td>
                <td className="px-3 py-2.5 border border-purple-600 text-right">{fmt(grandTotal.costSaving)}</td>
                <td className="px-3 py-2.5 border border-purple-600"></td>
                <td className="px-3 py-2.5 border border-purple-600"></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      {rateDetail && (() => {
        const entries = rateDetail.exchange_rates && rateDetail.exchange_rates.length > 0
          ? rateDetail.exchange_rates
          : (rateDetail.total_amount != null && rateDetail.exchange_rate != null
              ? [{ amount: rateDetail.total_amount, rate: rateDetail.exchange_rate }]
              : [])
        const total = entries.reduce((s, e) => s + e.amount * e.rate, 0)
        return (
          <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4" onClick={() => setRateDetail(null)}>
            <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm" onClick={e => e.stopPropagation()}>
              <div className="flex items-center justify-between p-5 border-b border-gray-100">
                <div>
                  <h3 className="font-semibold text-gray-800">Exchange Rate</h3>
                  <p className="text-xs text-gray-500 mt-0.5">{rateDetail.invoice_no}</p>
                </div>
                <button onClick={() => setRateDetail(null)} className="text-gray-300 hover:text-gray-500 text-xl">✕</button>
              </div>
              <div className="p-5">
                <p className="text-xs text-gray-400 mb-2">{rateDetail.currency}</p>
                <div className="space-y-1 mb-3">
                  {entries.map((e, i) => (
                    <div key={i} className="flex justify-between text-sm text-gray-600">
                      <span>{fmt(e.amount)} × {e.rate}</span>
                      <span className="text-gray-400">= {fmt(e.amount * e.rate)}</span>
                    </div>
                  ))}
                </div>
                <div className="flex justify-between border-t border-gray-100 pt-3 font-semibold text-gray-800">
                  <span>รวม</span>
                  <span>{fmt(total)} THB</span>
                </div>
              </div>
            </div>
          </div>
        )
      })()}
    </div>
  )
}

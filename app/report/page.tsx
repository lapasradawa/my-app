'use client'

import { useEffect, useState, useMemo } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'
import * as XLSX from 'xlsx'
import { isUnlocked } from '@/lib/auth'
import LockButton from '@/components/LockButton'
import PasswordModal from '@/components/PasswordModal'
import NavBar from '@/components/NavBar'

// ── PO Perspective types ──────────────────────────────────────────────────
interface PORow {
  id: string
  supplier: string
  project: string
  currency: string
  po_date: string | null
  po_rbs_ch_no: string | null
  po_rbs_th_no: string | null
  total_amount: number
  cost_saving: number | null
  cost_saving_pct: number | null
}

function poMonthKey(d: string): string {
  const dt = new Date(d + 'T00:00:00')
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`
}

function poMonthLabel(d: string): string {
  return new Date(d + 'T00:00:00').toLocaleDateString('en-US', { month: 'long', year: 'numeric' })
}

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
      {v || <span className="text-gray-400">—</span>}
    </td>
  )
}

export default function ReportPage() {
  const [perspective, setPerspective] = useState<'invoice' | 'po'>('invoice')

  // Invoice perspective state
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

  // PO perspective state
  const [poRows, setPORows] = useState<PORow[]>([])
  const [poSelectedMonths, setPOSelectedMonths] = useState<Set<string>>(new Set())
  const [cnyRate, setCnyRate] = useState(4.85)
  const [usdRate, setUsdRate] = useState(33.00)

  useEffect(() => { loadInvoices(); loadPO(); setUnlocked(isUnlocked()) }, [])

  async function loadInvoices() {
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

  async function loadPO() {
    const [{ data: pos }, { data: settings }] = await Promise.all([
      supabase.from('po_uploads')
        .select('id, supplier, project, currency, po_date, po_rbs_ch_no, po_rbs_th_no, total_amount, cost_saving, cost_saving_pct')
        .not('po_date', 'is', null)
        .order('po_date', { ascending: true }),
      supabase.from('cost_settings').select('key, value'),
    ])
    setPORows((pos ?? []) as PORow[])
    if (settings) {
      const m = Object.fromEntries((settings as { key: string; value: string }[]).map(r => [r.key, r.value]))
      if (m.cny_rate) setCnyRate(parseFloat(m.cny_rate))
      if (m.usd_rate) setUsdRate(parseFloat(m.usd_rate))
    }
  }

  async function load() { await loadInvoices() }

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

  // PO perspective derived data
  const poAllMonths = useMemo(() => {
    const map = new Map<string, string>()
    for (const p of poRows) {
      if (!p.po_date) continue
      const k = poMonthKey(p.po_date)
      if (!map.has(k)) map.set(k, poMonthLabel(p.po_date))
    }
    return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0]))
  }, [poRows])

  useEffect(() => {
    setPOSelectedMonths(new Set(poAllMonths.map(m => m[0])))
  }, [poAllMonths])

  const poGrouped = useMemo(() => {
    const result: { key: string; label: string; rows: PORow[] }[] = []
    for (const p of poRows) {
      if (!p.po_date) continue
      const k = poMonthKey(p.po_date)
      if (!poSelectedMonths.has(k)) continue
      let mg = result.find(m => m.key === k)
      if (!mg) { mg = { key: k, label: poMonthLabel(p.po_date), rows: [] }; result.push(mg) }
      mg.rows.push(p)
    }
    return result
  }, [poRows, poSelectedMonths])

  function getEstFobThb(p: PORow): number {
    return p.total_amount * (p.currency === 'USD' ? usdRate : cnyRate)
  }

  const poGrandTotal = useMemo(() => ({
    fobCny: sumN(poGrouped.flatMap(g => g.rows).map(p => p.currency === 'CNY' ? p.total_amount : null)),
    fobUsd: sumN(poGrouped.flatMap(g => g.rows).map(p => p.currency === 'USD' ? p.total_amount : null)),
    estThb: sumN(poGrouped.flatMap(g => g.rows).map(p => getEstFobThb(p))),
    costSaving: sumN(poGrouped.flatMap(g => g.rows).map(p => p.cost_saving)),
  }), [poGrouped, cnyRate, usdRate])

  function exportPOExcel() {
    const header = ['PO Issue MONTH', 'PO RBS CH No.', 'PO RBS TH No.', 'FOB CNY', 'FOB USD', 'Estimated FOB THB', 'Cost saving (THB)', 'Cost saving (%)']
    const aoa: (string | number | null)[][] = [header]
    for (const mg of poGrouped) {
      mg.rows.forEach((p, i) => {
        aoa.push([
          i === 0 ? mg.label : '',
          p.po_rbs_ch_no ?? '',
          p.po_rbs_th_no ?? '',
          p.currency === 'CNY' ? p.total_amount : null,
          p.currency === 'USD' ? p.total_amount : null,
          getEstFobThb(p),
          p.cost_saving,
          p.cost_saving_pct != null ? p.cost_saving_pct / 100 : null,
        ])
      })
      const g = mg.rows
      aoa.push([
        mg.label + ' Total', '', '',
        sumN(g.map(p => p.currency === 'CNY' ? p.total_amount : null)),
        sumN(g.map(p => p.currency === 'USD' ? p.total_amount : null)),
        sumN(g.map(p => getEstFobThb(p))),
        sumN(g.map(p => p.cost_saving)),
        null,
      ])
    }
    aoa.push(['GRAND TOTAL', '', '', poGrandTotal.fobCny, poGrandTotal.fobUsd, poGrandTotal.estThb, poGrandTotal.costSaving, null])
    const ws = XLSX.utils.aoa_to_sheet(aoa)
    const numCols = [3, 4, 5, 6]
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
    for (let r = 1; r <= range.e.r; r++) {
      numCols.forEach(c => {
        const addr = XLSX.utils.encode_cell({ r, c })
        if (ws[addr] && typeof ws[addr].v === 'number') ws[addr].z = '#,##0.00'
      })
      const pctAddr = XLSX.utils.encode_cell({ r, c: 7 })
      if (ws[pctAddr] && typeof ws[pctAddr].v === 'number') ws[pctAddr].z = '0.00%'
    }
    ws['!cols'] = [{ wch: 20 }, { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 14 }, { wch: 20 }, { wch: 18 }, { wch: 14 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'PO Report')
    XLSX.writeFile(wb, `PO_Report_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

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

      <NavBar onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />

      <div className="max-w-full mx-auto px-6 py-8">
        <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">Report</h1>
            <p className="text-sm text-gray-500 mt-1">สรุปยอด FOB และ Cost Saving รายเดือน</p>
          </div>
          <div className="flex items-center gap-3">
            {/* Perspective dropdown */}
            <div className="flex rounded-lg overflow-hidden border border-gray-200 text-sm">
              <button
                onClick={() => setPerspective('invoice')}
                className={`px-4 py-2 font-medium transition-colors ${perspective === 'invoice' ? 'bg-amber-400 text-gray-900' : 'bg-white text-gray-500 hover:bg-amber-50'}`}
              >
                Invoice Perspective
              </button>
              <button
                onClick={() => setPerspective('po')}
                className={`px-4 py-2 font-medium transition-colors border-l border-gray-200 ${perspective === 'po' ? 'bg-blue-600 text-white' : 'bg-white text-gray-500 hover:bg-blue-50'}`}
              >
                PO Perspective
              </button>
            </div>
            <button
              onClick={perspective === 'invoice' ? exportExcel : exportPOExcel}
              className="px-4 py-2 bg-green-600 text-white text-sm font-medium rounded-lg hover:bg-green-700 transition-colors"
            >
              Export Excel
            </button>
          </div>
        </div>

        {/* ── Invoice Perspective ──────────────────────────────────────── */}
        {perspective === 'invoice' && (<>

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
                          {fmt(c.actualThb) || <span className="text-gray-400">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-center text-gray-600 whitespace-nowrap text-xs">
                          {dueDateStr(inv.bl_date) || <span className="text-gray-400">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-center text-gray-600 whitespace-nowrap text-xs">
                          {fmtDate(inv.payment_date) || <span className="text-gray-400">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">
                          {inv.cost_saving != null ? fmt(inv.cost_saving) : <span className="text-gray-400">—</span>}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">
                          {inv.cost_saving_pct != null ? `${inv.cost_saving_pct}%` : <span className="text-gray-400">—</span>}
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
                              {fmtDate(commVal) || <span className="text-gray-400">—</span>}
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

        </>)} {/* end invoice perspective */}

        {/* ── PO Perspective ───────────────────────────────────────────── */}
        {perspective === 'po' && (<>
          {/* Month filter */}
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 mb-4">
            <div className="flex items-center gap-3 mb-2">
              <span className="text-xs font-semibold text-gray-600 uppercase tracking-wide">PO Issue Month</span>
              <button onClick={() => setPOSelectedMonths(new Set(poAllMonths.map(m => m[0])))} className="text-xs text-blue-600 hover:underline">เลือกทั้งหมด</button>
              <button onClick={() => setPOSelectedMonths(new Set())} className="text-xs text-gray-400 hover:underline">ยกเลิกทั้งหมด</button>
            </div>
            <div className="flex flex-wrap gap-2">
              {poAllMonths.map(([k, label]) => {
                const active = poSelectedMonths.has(k)
                return (
                  <button key={k} onClick={() => setPOSelectedMonths(prev => { const n = new Set(prev); n.has(k) ? n.delete(k) : n.add(k); return n })}
                    className={`px-3 py-1 rounded-full text-xs font-medium border transition-colors ${active ? 'bg-blue-600 border-blue-600 text-white' : 'bg-white border-gray-300 text-gray-500 hover:border-blue-300'}`}>
                    {label}
                  </button>
                )
              })}
              {poAllMonths.length === 0 && <p className="text-xs text-gray-400">ไม่มีข้อมูล PO ที่มีวันที่เปิด</p>}
            </div>
          </div>

          {/* PO table */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-auto shadow-sm">
            <table className="text-sm border-collapse w-full">
              <thead>
                <tr className="bg-blue-600 text-white">
                  {(['PO Issue MONTH', 'PO RBS CH No.', 'PO RBS TH No.', 'FOB CNY', 'FOB USD', 'Estimated FOB THB', 'Cost saving (THB)', 'Cost saving (%)']).map((h, i) => (
                    <th key={h} className={`px-3 py-2.5 border border-blue-500 whitespace-nowrap font-bold text-xs ${i >= 3 ? 'text-right' : 'text-left'}`}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {poGrouped.flatMap(mg => {
                  const mFobCny = sumN(mg.rows.map(p => p.currency === 'CNY' ? p.total_amount : null))
                  const mFobUsd = sumN(mg.rows.map(p => p.currency === 'USD' ? p.total_amount : null))
                  const mEstThb = sumN(mg.rows.map(p => getEstFobThb(p)))
                  const mCost = sumN(mg.rows.map(p => p.cost_saving))
                  return [
                    ...mg.rows.map((p, i) => (
                      <tr key={p.id} className="border-b border-gray-100 hover:bg-gray-50">
                        <td className="px-3 py-2 border border-gray-200 text-gray-700 whitespace-nowrap">{i === 0 ? mg.label : ''}</td>
                        <td className="px-3 py-2 border border-gray-200 font-mono text-gray-800 text-xs">{p.po_rbs_ch_no || <span className="text-gray-400">—</span>}</td>
                        <td className="px-3 py-2 border border-gray-200 font-mono text-gray-800 text-xs">{p.po_rbs_th_no || <span className="text-gray-400">—</span>}</td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">{p.currency === 'CNY' ? fmt(p.total_amount) : <span className="text-gray-300">—</span>}</td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">{p.currency === 'USD' ? fmt(p.total_amount) : <span className="text-gray-300">—</span>}</td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">{fmt(getEstFobThb(p))}</td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">{p.cost_saving != null ? fmt(p.cost_saving) : <span className="text-gray-400">—</span>}</td>
                        <td className="px-3 py-2 border border-gray-200 text-right text-gray-700">{p.cost_saving_pct != null ? `${p.cost_saving_pct}%` : <span className="text-gray-400">—</span>}</td>
                      </tr>
                    )),
                    <tr key={`${mg.key}-total`} className="bg-blue-50 font-semibold text-gray-800">
                      <td className="px-3 py-2 border border-blue-200 text-blue-900 whitespace-nowrap">{mg.label} Total</td>
                      <td className="px-3 py-2 border border-blue-200" colSpan={2}></td>
                      <td className="px-3 py-2 border border-blue-200 text-right">{fmt(mFobCny)}</td>
                      <td className="px-3 py-2 border border-blue-200 text-right">{fmt(mFobUsd)}</td>
                      <td className="px-3 py-2 border border-blue-200 text-right">{fmt(mEstThb)}</td>
                      <td className="px-3 py-2 border border-blue-200 text-right">{fmt(mCost)}</td>
                      <td className="px-3 py-2 border border-blue-200"></td>
                    </tr>,
                  ]
                })}
                {poGrouped.length === 0 && (
                  <tr><td colSpan={8} className="px-4 py-8 text-center text-gray-400 text-sm">ไม่มีข้อมูล PO</td></tr>
                )}
                <tr className="bg-blue-800 text-white font-bold border-t-2 border-blue-900">
                  <td className="px-3 py-2.5 border border-blue-700 whitespace-nowrap">GRAND TOTAL {new Date().toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' })}</td>
                  <td className="px-3 py-2.5 border border-blue-700" colSpan={2}></td>
                  <td className="px-3 py-2.5 border border-blue-700 text-right">{fmt(poGrandTotal.fobCny)}</td>
                  <td className="px-3 py-2.5 border border-blue-700 text-right">{fmt(poGrandTotal.fobUsd)}</td>
                  <td className="px-3 py-2.5 border border-blue-700 text-right">{fmt(poGrandTotal.estThb)}</td>
                  <td className="px-3 py-2.5 border border-blue-700 text-right">{fmt(poGrandTotal.costSaving)}</td>
                  <td className="px-3 py-2.5 border border-blue-700"></td>
                </tr>
              </tbody>
            </table>
          </div>
          <p className="text-xs text-gray-400 mt-2">Estimated FOB THB คำนวณจาก ESTIMATE RATES ในหน้า Cost Compare (CNY: {cnyRate}, USD: {usdRate})</p>
        </>)}

      </div> {/* end max-w-full */}

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

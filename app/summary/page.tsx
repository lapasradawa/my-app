'use client'

import { useEffect, useState, useMemo, useCallback } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import LockButton from '@/components/LockButton'

// ─── Types ────────────────────────────────────────────────────────────────────
interface ExchangeRateEntry { amount: number; rate: number }
interface InvoiceData {
  id: string
  invoice_no: string
  supplier: string | null
  vendor_code: string | null
  bl_date: string | null
  estimated_arrival: string | null
  currency: string | null
  total_amount: number | null
  exchange_rate: number | null
  exchange_rates: ExchangeRateEntry[] | null
  rows: { code: string; description: string; qty: number; po: string }[] | null
}
interface PoItem {
  item_code: string
  supplier: string
  fob_price: number | null
  currency: string | null
}
interface LineItem {
  item_code: string
  description: string
  qty: number
  po: string
  invoice_no: string
  invoice_id: string
  supplier: string
  vendor_code: string | null
  bl_date: string | null
  month_key: string | null
  currency: string | null
  fob_price: number | null
  fob_total: number | null
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function mKey(d: string | null): string | null {
  if (!d) return null
  const dt = new Date(d + 'T00:00:00')
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`
}
function mLabel(k: string): string {
  const [y, m] = k.split('-')
  const names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
  return `${names[parseInt(m) - 1]} ${y}`
}
function fmtFobByCurrency(map: Map<string, number>): string {
  return Array.from(map.entries())
    .filter(([, v]) => v > 0)
    .sort((a, b) => b[1] - a[1])
    .map(([ccy, amt]) => `${ccy} ${fmt(amt, 2)}`)
    .join(' · ')
}

function fmt(n: number, dec = 2) {
  return n.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec })
}
function generateMonthKeys(count = 18): string[] {
  const keys: string[] = []
  const d = new Date()
  for (let i = 0; i < count; i++) {
    keys.unshift(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`)
    d.setMonth(d.getMonth() - 1)
  }
  return keys
}

const PALETTE = ['#3d8b82','#d4962a','#c85a3a','#6b5ea8','#2a7c9a','#c87a3a','#5a9a6b','#c85a82']

// ─── DonutChart ───────────────────────────────────────────────────────────────
function DonutChart({ slices, total }: { slices: { label: string; value: number; color: string }[]; total: number }) {
  const cx = 80, cy = 80, R = 68, r = 44
  let angle = -Math.PI / 2
  const arcs = slices.filter(s => s.value > 0).map(s => {
    const sweep = total > 0 ? (s.value / total) * 2 * Math.PI : 0
    const sa = angle, ea = angle + sweep
    angle = ea
    const x1 = cx + R * Math.cos(sa), y1 = cy + R * Math.sin(sa)
    const x2 = cx + R * Math.cos(ea), y2 = cy + R * Math.sin(ea)
    const ix1 = cx + r * Math.cos(ea), iy1 = cy + r * Math.sin(ea)
    const ix2 = cx + r * Math.cos(sa), iy2 = cy + r * Math.sin(sa)
    const large = sweep > Math.PI ? 1 : 0
    return { ...s, path: `M ${x1} ${y1} A ${R} ${R} 0 ${large} 1 ${x2} ${y2} L ${ix1} ${iy1} A ${r} ${r} 0 ${large} 0 ${ix2} ${iy2} Z` }
  })
  return (
    <svg viewBox="0 0 160 160" width={160} height={160}>
      {arcs.map((arc, i) => <path key={i} d={arc.path} fill={arc.color} opacity={0.88} />)}
      <circle cx={cx} cy={cy} r={r - 3} fill="#faf5ee" />
      <text x={cx} y={cy - 8} textAnchor="middle" fontSize={10} fill="#8a7a6a" fontWeight="600">Total</text>
      <text x={cx} y={cy + 9} textAnchor="middle" fontSize={15} fill="#3a2a1a" fontWeight="800">
        {total >= 1000 ? `${(total / 1000).toFixed(1)}k` : total.toLocaleString()}
      </text>
      <text x={cx} y={cy + 23} textAnchor="middle" fontSize={9} fill="#aaa">units</text>
    </svg>
  )
}

// ─── LineChart ────────────────────────────────────────────────────────────────
function LineChart({ months, series }: { months: string[]; series: { label: string; color: string; values: number[] }[] }) {
  const W = 520, H = 160, padL = 36, padR = 16, padT = 14, padB = 28
  const cW = W - padL - padR, cH = H - padT - padB
  const allVals = series.flatMap(s => s.values)
  const maxVal = Math.max(...allVals, 1)
  const xPos = (i: number) => padL + (months.length < 2 ? cW / 2 : (i / (months.length - 1)) * cW)
  const yPos = (v: number) => padT + cH - (v / maxVal) * cH
  return (
    <svg viewBox={`0 0 ${W} ${H}`} className="w-full" style={{ maxHeight: 175 }}>
      {[0, 0.25, 0.5, 0.75, 1].map(frac => {
        const y = padT + cH * (1 - frac)
        const val = Math.round(maxVal * frac)
        return (
          <g key={frac}>
            <line x1={padL} y1={y} x2={W - padR} y2={y} stroke="#e8dcc8" strokeWidth={1} strokeDasharray={frac === 0 ? '0' : '3,3'} />
            <text x={padL - 4} y={y + 4} textAnchor="end" fontSize={8} fill="#bbb">
              {val >= 1000 ? `${(val / 1000).toFixed(0)}k` : val}
            </text>
          </g>
        )
      })}
      {series.map(s => (
        <g key={s.label}>
          <polyline points={months.map((_, i) => `${xPos(i)},${yPos(s.values[i])}`).join(' ')}
            fill="none" stroke={s.color} strokeWidth={2.5} strokeLinejoin="round" strokeLinecap="round" />
          {months.map((_, i) => s.values[i] > 0 && (
            <circle key={i} cx={xPos(i)} cy={yPos(s.values[i])} r={4}
              fill={s.color} stroke="#faf5ee" strokeWidth={2} />
          ))}
        </g>
      ))}
      {months.map((m, i) => (
        <text key={m} x={xPos(i)} y={H - 5} textAnchor="middle" fontSize={8} fill="#bbb">
          {mLabel(m).slice(0, 3)}
        </text>
      ))}
    </svg>
  )
}

// ─── MiniBarChart ─────────────────────────────────────────────────────────────
function MiniBarChart({ data }: { data: { label: string; value: number }[] }) {
  const max = Math.max(...data.map(d => d.value), 1)
  const W = 320, H = 90, padB = 18, padT = 14
  const barW = Math.floor((W - 10) / data.length) - 3
  return (
    <svg viewBox={`0 0 ${W} ${H}`} className="w-full" style={{ maxHeight: 95 }}>
      {data.map((d, i) => {
        const bH = Math.max(2, (d.value / max) * (H - padB - padT))
        const x = 5 + i * (barW + 3)
        const y = H - padB - bH
        return (
          <g key={d.label}>
            <rect x={x} y={y} width={barW} height={bH} fill="#3d8b82" rx={2} opacity={0.8} />
            {d.value > 0 && (
              <text x={x + barW / 2} y={y - 2} textAnchor="middle" fontSize={7} fill="#6a5a4a" fontWeight="600">
                {d.value >= 1000 ? `${(d.value / 1000).toFixed(1)}k` : d.value}
              </text>
            )}
            <text x={x + barW / 2} y={H - 3} textAnchor="middle" fontSize={7} fill="#bbb">{d.label}</text>
          </g>
        )
      })}
    </svg>
  )
}

// ─── Page ─────────────────────────────────────────────────────────────────────
export default function SummaryPage() {
  const [invoices, setInvoices] = useState<InvoiceData[]>([])
  const [poItems, setPoItems] = useState<PoItem[]>([])
  const [loading, setLoading] = useState(true)
  const [rowsLoading, setRowsLoading] = useState(true)
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(() => {
    const now = new Date()
    return new Set([`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`])
  })
  const [selectedItem, setSelectedItem] = useState<string | null>(null)
  const [itemSearch, setItemSearch] = useState('')
  const [showDetail, setShowDetail] = useState(false)
  const [periodOpen, setPeriodOpen] = useState(false)

  const allMonthKeys = useMemo(() => generateMonthKeys(18), [])
  const chart12Months = useMemo(() => generateMonthKeys(12), [])

  useEffect(() => {
    async function load() {
      // Stage 1: metadata only — KPI / charts appear immediately
      const [metaRes, poRes] = await Promise.all([
        supabase.from('invoices').select('id, invoice_no, supplier, vendor_code, bl_date, estimated_arrival, currency, total_amount, exchange_rate, exchange_rates').order('bl_date', { ascending: false }),
        supabase.from('po_items').select('item_code, supplier, fob_price, currency')
      ])
      if (metaRes.data) setInvoices(metaRes.data as InvoiceData[])
      if (poRes.data) setPoItems(poRes.data as PoItem[])
      setLoading(false)

      // Stage 2: fetch rows in background — item cards / table populate after
      const rowsRes = await supabase.from('invoices').select('id, rows')
      if (rowsRes.data) {
        const rowMap = new Map(rowsRes.data.map((r: { id: string; rows: InvoiceData['rows'] }) => [r.id, r.rows]))
        setInvoices(prev => prev.map(inv => ({ ...inv, rows: rowMap.get(inv.id) || [] })))
      }
      setRowsLoading(false)
    }
    load()
  }, [])

  const priceMap = useMemo(() => {
    const m = new Map<string, { price: number | null; currency: string | null }>()
    for (const p of poItems) m.set(`${p.item_code}|${p.supplier}`, { price: p.fob_price, currency: p.currency })
    return m
  }, [poItems])

  const vendorCodeMap = useMemo(() => {
    const m = new Map<string, string>()
    for (const inv of invoices) if (inv.supplier && inv.vendor_code) m.set(inv.supplier, inv.vendor_code)
    return m
  }, [invoices])

  const allLines = useMemo<LineItem[]>(() => {
    const lines: LineItem[] = []
    for (const inv of invoices) {
      if (!inv.rows || inv.rows.length === 0) continue
      const supplier = inv.supplier || '—'
      const refDate = inv.bl_date || inv.estimated_arrival
      const mk = mKey(refDate)
      const resolvedVendorCode = inv.vendor_code || vendorCodeMap.get(supplier) || null
      for (const row of inv.rows) {
        if (!row.code) continue
        const lookup = priceMap.get(`${row.code}|${supplier}`)
        const fob = lookup?.price ?? null
        lines.push({
          item_code: row.code, description: row.description || '', qty: row.qty || 0, po: row.po || '',
          invoice_no: inv.invoice_no, invoice_id: inv.id, supplier, vendor_code: resolvedVendorCode,
          bl_date: inv.bl_date, month_key: mk, currency: lookup?.currency ?? inv.currency,
          fob_price: fob, fob_total: fob != null && row.qty ? fob * row.qty : null,
        })
      }
    }
    return lines
  }, [invoices, priceMap, vendorCodeMap])

  const filteredLines = useMemo(() => {
    if (selectedMonths.size === 0) return allLines
    return allLines.filter(l => l.month_key && selectedMonths.has(l.month_key))
  }, [allLines, selectedMonths])

  const supplierSummary = useMemo(() => {
    const map = new Map<string, { qty: number; fob: number; fobByCurrency: Map<string, number>; items: Set<string>; invoices: Set<string> }>()
    for (const l of filteredLines) {
      const cur = map.get(l.supplier) || { qty: 0, fob: 0, fobByCurrency: new Map(), items: new Set(), invoices: new Set() }
      cur.qty += l.qty
      if (l.fob_total) {
        cur.fob += l.fob_total
        const ccy = l.currency || '?'
        cur.fobByCurrency.set(ccy, (cur.fobByCurrency.get(ccy) || 0) + l.fob_total)
      }
      cur.items.add(l.item_code); cur.invoices.add(l.invoice_no)
      map.set(l.supplier, cur)
    }
    return Array.from(map.entries())
      .map(([supplier, v]) => ({ supplier, ...v, itemCount: v.items.size, invoiceCount: v.invoices.size }))
      .sort((a, b) => b.qty - a.qty)
  }, [filteredLines])

  const itemSummary = useMemo(() => {
    const map = new Map<string, { description: string; totalQty: number; totalFob: number; fobByCurrency: Map<string, number>; suppliers: Map<string, { qty: number; fob: number; fobByCurrency: Map<string, number> }>; invoiceList: Map<string, { id: string; qty: number; supplier: string }>; byMonth: Map<string, number> }>()
    for (const l of filteredLines) {
      const cur = map.get(l.item_code) || { description: l.description, totalQty: 0, totalFob: 0, fobByCurrency: new Map(), suppliers: new Map(), invoiceList: new Map(), byMonth: new Map() }
      cur.totalQty += l.qty
      if (l.fob_total) {
        cur.totalFob += l.fob_total
        const ccy = l.currency || '?'
        cur.fobByCurrency.set(ccy, (cur.fobByCurrency.get(ccy) || 0) + l.fob_total)
      }
      const sup = cur.suppliers.get(l.supplier) || { qty: 0, fob: 0, fobByCurrency: new Map() }
      sup.qty += l.qty
      if (l.fob_total) {
        sup.fob += l.fob_total
        const ccy = l.currency || '?'
        sup.fobByCurrency.set(ccy, (sup.fobByCurrency.get(ccy) || 0) + l.fob_total)
      }
      cur.suppliers.set(l.supplier, sup)
      const inv = cur.invoiceList.get(l.invoice_no) || { id: l.invoice_id, qty: 0, supplier: l.supplier }
      inv.qty += l.qty
      cur.invoiceList.set(l.invoice_no, inv)
      if (l.month_key) cur.byMonth.set(l.month_key, (cur.byMonth.get(l.month_key) || 0) + l.qty)
      map.set(l.item_code, cur)
    }
    return Array.from(map.entries()).map(([code, v]) => ({ code, ...v })).sort((a, b) => b.totalQty - a.totalQty)
  }, [filteredLines])

  const filteredItemSummary = useMemo(() => {
    if (!itemSearch.trim()) return itemSummary
    const q = itemSearch.toLowerCase()
    return itemSummary.filter(i => i.code.toLowerCase().includes(q) || i.description.toLowerCase().includes(q))
  }, [itemSummary, itemSearch])

  const selectedItemData = useMemo(() =>
    selectedItem ? itemSummary.find(i => i.code === selectedItem) : null,
    [selectedItem, itemSummary])

  // actualThb per invoice: exchange_rates array sum, or total_amount × exchange_rate
  const invoiceThbMap = useMemo(() => {
    const m = new Map<string, number>()
    for (const inv of invoices) {
      let thb: number | null = null
      if (inv.exchange_rates && inv.exchange_rates.length > 0) {
        thb = inv.exchange_rates.reduce((s, e) => s + e.amount * e.rate, 0)
      } else if (inv.total_amount != null && inv.exchange_rate != null) {
        thb = inv.total_amount * inv.exchange_rate
      }
      if (thb != null) m.set(inv.id, thb)
    }
    return m
  }, [invoices])

  const totalQty = filteredLines.reduce((s, l) => s + l.qty, 0)
  // Use estimated_arrival for grouping (matches Report page "เข้าคลัง" tab)
  const totalActualThb = useMemo(() => {
    let sum = 0
    for (const inv of invoices) {
      if (selectedMonths.size > 0) {
        const mk = mKey(inv.estimated_arrival)
        if (!mk || !selectedMonths.has(mk)) continue
      }
      sum += invoiceThbMap.get(inv.id) || 0
    }
    return sum
  }, [invoices, selectedMonths, invoiceThbMap])
  const totalInvoices = useMemo(() => new Set(filteredLines.map(l => l.invoice_no)).size, [filteredLines])
  const totalSuppliers = useMemo(() => new Set(filteredLines.map(l => l.supplier)).size, [filteredLines])

  const lineSeries = useMemo(() => {
    const top4 = supplierSummary.slice(0, 4)
    return top4.map((s, i) => ({
      label: s.supplier,
      color: PALETTE[i % PALETTE.length],
      values: chart12Months.map(month =>
        allLines.filter(l => l.supplier === s.supplier && l.month_key === month).reduce((sum, l) => sum + l.qty, 0)
      )
    }))
  }, [supplierSummary, chart12Months, allLines])

  const toggleMonth = useCallback((k: string) => {
    setSelectedMonths(prev => { const n = new Set(prev); n.has(k) ? n.delete(k) : n.add(k); return n })
  }, [])

  function exportExcel() {
    const lines = selectedItem ? filteredLines.filter(l => l.item_code === selectedItem) : filteredLines
    const rows = lines.map(l => ({
      'Item no': l.item_code, 'Item name': l.description, 'Supplier': l.supplier,
      'Vendor Code': l.vendor_code || '', 'Quantity': l.qty, 'Invoice': l.invoice_no, 'PO': l.po,
      'FOB Unit Price': l.fob_price ?? '', 'FOB Total Price': l.fob_total ?? '', 'Original Currency': l.currency || '',
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    ws['!cols'] = [{ wch: 22 }, { wch: 50 }, { wch: 18 }, { wch: 16 }, { wch: 10 }, { wch: 18 }, { wch: 16 }, { wch: 15 }, { wch: 15 }, { wch: 16 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, selectedItem ? `Item_${selectedItem.slice(0, 20)}` : 'Summary')
    XLSX.writeFile(wb, `item-summary-${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div style={{ background: '#ede5d4', minHeight: '100vh', fontFamily: 'system-ui,-apple-system,sans-serif' }}>

      {/* ── Nav ── */}
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-6 text-sm sticky top-0 z-20 shadow-sm flex-wrap">
        <span className="font-bold text-gray-900">Import PO</span>
        <Link href="/" className="text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/calendar" className="text-gray-500 hover:text-gray-800 transition-colors">Calendar</Link>
        <Link href="/report" className="text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/compare" className="text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-gray-500 hover:text-gray-800 transition-colors">PO Builder</Link>
        <div className="relative group">
          <span className="text-blue-600 cursor-default">Summary ▾</span>
          <div className="absolute left-0 top-full pt-1 hidden group-hover:block z-50">
            <div className="bg-white border border-gray-200 rounded-lg shadow-lg py-1 min-w-[150px]">
              <Link href="/summary" className="block px-4 py-2 text-sm text-blue-600 hover:bg-blue-50">Item Summary</Link>
              <Link href="/qc/summary" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">QC Summary</Link>
            </div>
          </div>
        </div>
        <Link href="/qc" className="text-gray-500 hover:text-gray-800 transition-colors">QC Report</Link>
        <Link href="/guide" className="text-gray-500 hover:text-gray-800 transition-colors">Guide</Link>
        <div className="ml-auto"><LockButton /></div>
      </nav>

      {/* ── Hero header ── */}
      <div style={{ background: '#1e3340', padding: '24px 32px' }}>
        <div style={{ maxWidth: 1400, margin: '0 auto', display: 'flex', alignItems: 'center', gap: 24, flexWrap: 'wrap' }}>
          {/* Title */}
          <div style={{ flex: 1, minWidth: 200 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: '#3d8b82', textTransform: 'uppercase', letterSpacing: '0.12em', marginBottom: 4 }}>Import PO</div>
            <div style={{ fontSize: 30, fontWeight: 900, color: '#d4962a', lineHeight: 1.1 }}>ITEM SUMMARY</div>
            <div style={{ fontSize: 13, color: '#7a9aaa', marginTop: 4 }}>สรุปยอดสั่งซื้อแยกตาม Supplier และ Item</div>
          </div>
          {/* KPI cards */}
          {[
            { icon: '💰', value: totalActualThb > 0 ? `฿${fmt(totalActualThb, 0)}` : '—', label: 'Actual FOB THB', bg: '#d4962a' },
            { icon: '🏭', value: String(totalSuppliers), label: 'Suppliers', bg: '#c85a3a' },
            { icon: '📋', value: String(totalInvoices), label: 'Invoices', bg: '#6b5ea8' },
          ].map(card => (
            <div key={card.label} style={{ background: card.bg, borderRadius: 14, padding: '14px 20px', minWidth: 130, display: 'flex', alignItems: 'center', gap: 12 }}>
              <span style={{ fontSize: 26 }}>{card.icon}</span>
              <div>
                <div style={{ fontSize: 22, fontWeight: 900, color: '#fff', lineHeight: 1.1 }}>{card.value}</div>
                <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.75)', fontWeight: 600, marginTop: 2 }}>{card.label}</div>
              </div>
            </div>
          ))}
          <button onClick={exportExcel}
            style={{ background: 'rgba(61,139,130,0.25)', border: '1.5px solid #3d8b82', color: '#5ec5ba', borderRadius: 10, padding: '10px 18px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
            ↓ Export Excel
          </button>
        </div>
      </div>

      {/* ── Period dropdown bar ── */}
      <div style={{ background: '#18303c', padding: '10px 32px', position: 'relative', zIndex: 30 }}>
        <div style={{ maxWidth: 1400, margin: '0 auto', display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 10, fontWeight: 800, color: '#5a8a9a', textTransform: 'uppercase', letterSpacing: '0.12em' }}>Period:</span>

          {/* Dropdown trigger */}
          <div style={{ position: 'relative' }}>
            <button onClick={() => setPeriodOpen(o => !o)} style={{
              display: 'flex', alignItems: 'center', gap: 8,
              padding: '6px 14px', borderRadius: 8, cursor: 'pointer',
              background: '#1e3a4a', border: '1px solid #2e5060',
              color: '#d4c8a8', fontSize: 11, fontWeight: 700,
              minWidth: 200,
            }}>
              <span style={{ flex: 1, textAlign: 'left' }}>
                {selectedMonths.size === 0
                  ? 'Select period…'
                  : selectedMonths.size === allMonthKeys.length
                  ? 'All months'
                  : selectedMonths.size === 1
                  ? mLabel([...selectedMonths][0])
                  : `${selectedMonths.size} months selected`}
              </span>
              <span style={{ fontSize: 9, color: '#5a8a9a' }}>{periodOpen ? '▲' : '▼'}</span>
            </button>

            {periodOpen && (
              <div style={{
                position: 'absolute', top: 'calc(100% + 6px)', left: 0,
                background: '#1a2e3c', border: '1px solid #2e5060', borderRadius: 12,
                boxShadow: '0 8px 32px rgba(0,0,0,0.4)', padding: 12, minWidth: 280, zIndex: 100,
              }}>
                {/* Quick actions */}
                <div style={{ display: 'flex', gap: 8, marginBottom: 10, paddingBottom: 8, borderBottom: '1px solid #2a4455' }}>
                  <button onClick={() => setSelectedMonths(new Set(allMonthKeys))}
                    style={{ flex: 1, padding: '4px 0', borderRadius: 6, border: 'none', background: '#3d8b82', color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    All
                  </button>
                  <button onClick={() => setSelectedMonths(new Set())}
                    style={{ flex: 1, padding: '4px 0', borderRadius: 6, border: 'none', background: '#2a4455', color: '#8a9aaa', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    Clear
                  </button>
                  <button onClick={() => setPeriodOpen(false)}
                    style={{ padding: '4px 10px', borderRadius: 6, border: 'none', background: '#d4962a', color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    Done
                  </button>
                </div>
                {/* Month grid */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 5 }}>
                  {allMonthKeys.map(k => (
                    <button key={k} onClick={() => toggleMonth(k)} style={{
                      padding: '5px 4px', borderRadius: 6, cursor: 'pointer', fontSize: 10, fontWeight: 700,
                      background: selectedMonths.has(k) ? '#d4962a' : 'transparent',
                      color: selectedMonths.has(k) ? '#1a2d3a' : '#8a9aaa',
                      border: selectedMonths.has(k) ? '1px solid #d4962a' : '1px solid #2a4455',
                      transition: 'all 0.12s',
                    }}>{mLabel(k)}</button>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Summary chips */}
          {selectedMonths.size > 0 && [...selectedMonths].sort().map(k => (
            <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 4, background: '#d4962a', borderRadius: 6, padding: '3px 8px' }}>
              <span style={{ fontSize: 10, fontWeight: 700, color: '#1a2d3a' }}>{mLabel(k)}</span>
              <button onClick={() => toggleMonth(k)} style={{ background: 'none', border: 'none', color: '#1a2d3a', fontSize: 11, cursor: 'pointer', padding: 0, lineHeight: 1, opacity: 0.7 }}>×</button>
            </div>
          ))}

          {selectedMonths.size > 0 && (
            <span style={{ color: '#4a7a8a', fontSize: 10, marginLeft: 'auto' }}>
              {filteredLines.length.toLocaleString()} lines
            </span>
          )}
        </div>
      </div>

      {loading ? (
        <div style={{ textAlign: 'center', padding: '80px 0', color: '#aaa', fontSize: 14 }}>กำลังโหลด...</div>
      ) : (
        <div style={{ maxWidth: 1400, margin: '0 auto', padding: '20px 24px', display: 'grid', gridTemplateColumns: '260px 1fr', gap: 16 }}>

          {/* ── Left column ── */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>

            {/* Donut: Units by Supplier */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, padding: '16px 16px 12px' }}>
              <div style={{ fontSize: 11, fontWeight: 800, color: '#5a4a3a', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10, textAlign: 'center' }}>
                Units by Supplier
              </div>
              <div style={{ display: 'flex', justifyContent: 'center' }}>
                <DonutChart
                  slices={supplierSummary.slice(0, 8).map((s, i) => ({ label: s.supplier, value: s.qty, color: PALETTE[i % PALETTE.length] }))}
                  total={totalQty}
                />
              </div>
              <div style={{ marginTop: 8, display: 'flex', flexDirection: 'column', gap: 5 }}>
                {supplierSummary.slice(0, 6).map((s, i) => (
                  <div key={s.supplier} style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
                    <div style={{ width: 9, height: 9, borderRadius: '50%', background: PALETTE[i % PALETTE.length], flexShrink: 0 }} />
                    <span style={{ fontSize: 11, color: '#6a5a4a', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{s.supplier}</span>
                    <span style={{ fontSize: 10, fontWeight: 800, color: '#4a3a2a' }}>
                      {totalQty > 0 ? ((s.qty / totalQty) * 100).toFixed(0) : 0}%
                    </span>
                  </div>
                ))}
              </div>
            </div>

            {/* Supplier ranked list */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, overflow: 'hidden', flex: 1 }}>
              <div style={{ padding: '10px 14px', borderBottom: '1px solid #e2d8c8', fontSize: 11, fontWeight: 800, color: '#5a4a3a', textTransform: 'uppercase', letterSpacing: '0.08em' }}>
                Supplier Breakdown
              </div>
              <div style={{ maxHeight: 340, overflowY: 'auto' }}>
                {supplierSummary.length === 0 ? (
                  <div style={{ padding: 20, textAlign: 'center', color: '#bbb', fontSize: 12 }}>ไม่มีข้อมูล</div>
                ) : supplierSummary.map((s, i) => (
                  <div key={s.supplier} style={{ padding: '9px 14px', borderBottom: '1px solid #f0ebe0', display: 'flex', alignItems: 'flex-start', gap: 9 }}>
                    <div style={{ width: 10, height: 10, borderRadius: '50%', background: PALETTE[i % PALETTE.length], marginTop: 3, flexShrink: 0 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: '#3a2a1a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{s.supplier}</div>
                      <div style={{ fontSize: 10, color: '#9a8a7a', marginTop: 2 }}>
                        <strong style={{ color: '#3a2a1a' }}>{s.qty.toLocaleString()}</strong> pcs · {s.itemCount} items · {s.invoiceCount} inv
                      </div>
                      {s.fobByCurrency.size > 0 && <div style={{ fontSize: 10, color: '#3d8b82', fontWeight: 700, marginTop: 1 }}>{fmtFobByCurrency(s.fobByCurrency)}</div>}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* ── Right column ── */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>

            {/* Line chart */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, padding: '16px 20px' }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10, flexWrap: 'wrap', gap: 8 }}>
                <div style={{ fontSize: 12, fontWeight: 800, color: '#5a4a3a', textTransform: 'uppercase', letterSpacing: '0.08em' }}>
                  Monthly Units — Top Suppliers (12 months)
                </div>
                <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
                  {lineSeries.map(s => (
                    <div key={s.label} style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                      <div style={{ width: 18, height: 3, borderRadius: 2, background: s.color }} />
                      <span style={{ fontSize: 10, color: '#8a7a6a' }}>{s.label.length > 14 ? s.label.slice(0, 14) + '…' : s.label}</span>
                    </div>
                  ))}
                </div>
              </div>
              {lineSeries.length === 0 ? (
                <div style={{ height: 100, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#ccc', fontSize: 12 }}>ไม่มีข้อมูล</div>
              ) : (
                <LineChart months={chart12Months} series={lineSeries} />
              )}
            </div>

            {/* Item detail panel */}
            {selectedItemData && (
              <div style={{ background: '#faf5ee', border: '2px solid #3d8b82', borderRadius: 16, padding: '18px 20px' }}>
                <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 14 }}>
                  <div>
                    <div style={{ fontSize: 9, fontWeight: 800, color: '#3d8b82', textTransform: 'uppercase', letterSpacing: '0.12em', marginBottom: 4 }}>Item Detail</div>
                    <div style={{ fontSize: 20, fontWeight: 900, color: '#2a2a1a', fontFamily: 'monospace' }}>{selectedItemData.code}</div>
                    <div style={{ fontSize: 12, color: '#8a7a6a', marginTop: 2 }}>{selectedItemData.description}</div>
                  </div>
                  <button onClick={() => setSelectedItem(null)}
                    style={{ background: 'none', border: 'none', fontSize: 18, color: '#ccc', cursor: 'pointer', padding: 4 }}>✕</button>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
                  <div>
                    <div style={{ fontSize: 9, fontWeight: 800, color: '#9a8a7a', textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: 10 }}>By Supplier</div>
                    {Array.from(selectedItemData.suppliers.entries()).sort((a, b) => b[1].qty - a[1].qty).map(([sup, v], si) => {
                      const pct = selectedItemData.totalQty > 0 ? (v.qty / selectedItemData.totalQty) * 100 : 0
                      return (
                        <div key={sup} style={{ marginBottom: 9 }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, marginBottom: 4 }}>
                            <span style={{ fontWeight: 600, color: '#4a3a2a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '65%' }}>{sup}</span>
                            <span style={{ color: '#8a7a6a', flexShrink: 0 }}>{v.qty.toLocaleString()} · {pct.toFixed(0)}%</span>
                          </div>
                          <div style={{ height: 7, background: '#ede8df', borderRadius: 4, overflow: 'hidden' }}>
                            <div style={{ height: '100%', width: `${pct}%`, background: PALETTE[si % PALETTE.length], borderRadius: 4 }} />
                          </div>
                        </div>
                      )
                    })}
                    <div style={{ marginTop: 10, paddingTop: 8, borderTop: '1px solid #e8dcc8', fontSize: 11, color: '#8a7a6a' }}>
                      รวม <strong style={{ color: '#3a2a1a' }}>{selectedItemData.totalQty.toLocaleString()}</strong> pcs
                      {selectedItemData.fobByCurrency.size > 0 && <> · FOB <strong style={{ color: '#3d7a5a' }}>{fmtFobByCurrency(selectedItemData.fobByCurrency)}</strong></>}
                    </div>
                  </div>
                  <div>
                    <div style={{ fontSize: 9, fontWeight: 800, color: '#9a8a7a', textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: 8 }}>ยอดสั่งรายเดือน (pcs)</div>
                    <MiniBarChart data={generateMonthKeys(12).map(k => ({ label: mLabel(k).slice(0, 3), value: selectedItemData.byMonth.get(k) || 0 }))} />
                  </div>
                </div>

                {/* Invoice list */}
                <div style={{ marginTop: 14, paddingTop: 12, borderTop: '1px solid #e8dcc8' }}>
                  <div style={{ fontSize: 9, fontWeight: 800, color: '#9a8a7a', textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: 8 }}>Invoices</div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    {Array.from(selectedItemData.invoiceList.entries())
                      .sort((a, b) => b[1].qty - a[1].qty)
                      .map(([invoiceNo, inv]) => (
                        <a key={invoiceNo} href={`/dashboard/${inv.id}`}
                          style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 10px', borderRadius: 8, background: '#f0ece4', textDecoration: 'none', transition: 'background 0.12s' }}
                          onMouseEnter={e => (e.currentTarget.style.background = '#e4ddd0')}
                          onMouseLeave={e => (e.currentTarget.style.background = '#f0ece4')}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                            <span style={{ fontSize: 11, fontWeight: 700, color: '#3d8b82', fontFamily: 'monospace' }}>{invoiceNo}</span>
                            <span style={{ fontSize: 10, color: '#9a8a7a' }}>{inv.supplier}</span>
                          </div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                            <span style={{ fontSize: 11, fontWeight: 700, color: '#3a2a1a' }}>{inv.qty.toLocaleString()} pcs</span>
                            <span style={{ fontSize: 10, color: '#3d8b82' }}>→</span>
                          </div>
                        </a>
                      ))}
                  </div>
                </div>
              </div>
            )}

            {/* Items grid + detail table */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, overflow: 'hidden' }}>
              {/* Header row */}
              <div style={{ padding: '12px 16px', borderBottom: '1px solid #e2d8c8', display: 'flex', alignItems: 'center', gap: 10, background: '#f5efe4' }}>
                <button onClick={() => setShowDetail(false)}
                  style={{ padding: '4px 14px', borderRadius: 8, fontSize: 11, fontWeight: 700, cursor: 'pointer', border: 'none',
                    background: !showDetail ? '#d4962a' : 'transparent', color: !showDetail ? '#fff' : '#8a7a6a' }}>
                  By Item
                </button>
                <button onClick={() => setShowDetail(true)}
                  style={{ padding: '4px 14px', borderRadius: 8, fontSize: 11, fontWeight: 700, cursor: 'pointer', border: 'none',
                    background: showDetail ? '#d4962a' : 'transparent', color: showDetail ? '#fff' : '#8a7a6a' }}>
                  Detail Lines
                </button>
                <input type="text" placeholder="Search item code / name…" value={itemSearch} onChange={e => setItemSearch(e.target.value)}
                  style={{ flex: 1, border: '1px solid #e2d8c8', borderRadius: 8, padding: '5px 10px', fontSize: 11, background: '#fff', outline: 'none', color: '#3a2a1a' }} />
                <button onClick={exportExcel}
                  style={{ padding: '5px 14px', borderRadius: 8, fontSize: 11, fontWeight: 700, background: '#3d8b82', color: '#fff', border: 'none', cursor: 'pointer', flexShrink: 0 }}>
                  ↓ Export
                </button>
              </div>

              {rowsLoading ? (
                <div style={{ textAlign: 'center', padding: '40px 0', color: '#bbb', fontSize: 12 }}>กำลังโหลด items…</div>
              ) : !showDetail ? (
                /* Item cards */
                <div style={{ maxHeight: 360, overflowY: 'auto', padding: 12, display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(145px, 1fr))', gap: 8 }}>
                  {filteredItemSummary.slice(0, 120).map(item => (
                    <button key={item.code}
                      onClick={() => setSelectedItem(selectedItem === item.code ? null : item.code)}
                      style={{
                        textAlign: 'left', padding: '11px 12px', borderRadius: 10, cursor: 'pointer', transition: 'all 0.15s',
                        background: selectedItem === item.code ? '#3d8b82' : '#fff',
                        border: `1.5px solid ${selectedItem === item.code ? '#3d8b82' : '#e2d8c8'}`,
                        boxShadow: selectedItem === item.code ? '0 2px 8px rgba(61,139,130,0.25)' : 'none',
                      }}>
                      <div style={{ fontSize: 11, fontWeight: 800, color: selectedItem === item.code ? '#fff' : '#3a2a1a', fontFamily: 'monospace', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {item.code}
                      </div>
                      <div style={{ fontSize: 9, color: selectedItem === item.code ? 'rgba(255,255,255,0.65)' : '#9a8a7a', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {item.description || '—'}
                      </div>
                      <div style={{ marginTop: 8, fontSize: 18, fontWeight: 900, color: selectedItem === item.code ? '#fff' : '#d4962a', lineHeight: 1 }}>
                        {item.totalQty.toLocaleString()}
                      </div>
                      <div style={{ fontSize: 9, color: selectedItem === item.code ? 'rgba(255,255,255,0.55)' : '#bbb', marginTop: 1 }}>pcs</div>
                    </button>
                  ))}
                  {filteredItemSummary.length === 0 && (
                    <div style={{ gridColumn: '1/-1', textAlign: 'center', padding: '30px 0', color: '#bbb', fontSize: 12 }}>ไม่พบ item</div>
                  )}
                </div>
              ) : (
                /* Detail table */
                <div style={{ overflowX: 'auto', maxHeight: 360, overflowY: 'auto' }}>
                  <table style={{ fontSize: 11, width: '100%', borderCollapse: 'collapse', minWidth: 800 }}>
                    <thead>
                      <tr style={{ background: '#f0ebe0', position: 'sticky', top: 0, zIndex: 1 }}>
                        {['Item no', 'Item name', 'Supplier', 'Vendor Code', 'QTY', 'Invoice', 'PO', 'FOB Unit', 'FOB Total', 'CCY'].map(h => (
                          <th key={h} style={{ padding: '8px 10px', textAlign: h === 'QTY' || h === 'FOB Unit' || h === 'FOB Total' ? 'right' : 'left', fontSize: 10, fontWeight: 800, color: '#8a7a6a', whiteSpace: 'nowrap', borderBottom: '1px solid #e2d8c8' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {(selectedItem ? filteredLines.filter(l => l.item_code === selectedItem) : filteredLines).slice(0, 500).map((l, i) => (
                        <tr key={i} onClick={() => setSelectedItem(l.item_code)}
                          style={{ borderBottom: '1px solid #f5efe8', cursor: 'pointer', background: i % 2 === 0 ? '#fff' : '#faf7f2' }}>
                          <td style={{ padding: '7px 10px', fontFamily: 'monospace', color: '#3a2a1a', whiteSpace: 'nowrap', fontWeight: 700 }}>{l.item_code}</td>
                          <td style={{ padding: '7px 10px', color: '#6a5a4a', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{l.description}</td>
                          <td style={{ padding: '7px 10px', color: '#5a4a3a', whiteSpace: 'nowrap' }}>{l.supplier}</td>
                          <td style={{ padding: '7px 10px', fontFamily: 'monospace', color: '#9a8a7a' }}>{l.vendor_code || '—'}</td>
                          <td style={{ padding: '7px 10px', textAlign: 'right', fontWeight: 700, color: '#2a2a1a' }}>{l.qty.toLocaleString()}</td>
                          <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                            <Link href={`/dashboard/${l.invoice_id}`} onClick={e => e.stopPropagation()}
                              style={{ color: '#3d8b82', textDecoration: 'none', fontWeight: 600 }}>{l.invoice_no}</Link>
                          </td>
                          <td style={{ padding: '7px 10px', color: '#9a8a7a', whiteSpace: 'nowrap' }}>{l.po || '—'}</td>
                          <td style={{ padding: '7px 10px', textAlign: 'right', color: '#5a4a3a' }}>{l.fob_price != null ? fmt(l.fob_price) : '—'}</td>
                          <td style={{ padding: '7px 10px', textAlign: 'right', fontWeight: 700, color: '#3d7a5a' }}>{l.fob_total != null ? fmt(l.fob_total, 0) : '—'}</td>
                          <td style={{ padding: '7px 10px', color: '#aaa' }}>{l.currency || '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filteredLines.length > 500 && !selectedItem && (
                    <div style={{ textAlign: 'center', padding: '10px 0', color: '#aaa', fontSize: 11 }}>
                      แสดง 500 รายการแรก จาก {filteredLines.length.toLocaleString()} — เลือก Item เพื่อดูทั้งหมด
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

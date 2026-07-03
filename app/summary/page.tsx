'use client'

import { useEffect, useState, useMemo, useCallback } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import LockButton from '@/components/LockButton'

// ─── Types ────────────────────────────────────────────────────────────────────
interface InvoiceData {
  id: string
  invoice_no: string
  supplier: string | null
  vendor_code: string | null
  bl_date: string | null
  estimated_arrival: string | null
  currency: string | null
  rows: { code: string; description: string; qty: number; po: string }[]
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
  month_key: string | null  // YYYY-MM from bl_date
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

// ─── Simple SVG bar chart ─────────────────────────────────────────────────────
function BarChart({ data }: { data: { label: string; value: number; color?: string }[] }) {
  const max = Math.max(...data.map(d => d.value), 1)
  const H = 100
  const barW = Math.max(20, Math.min(50, Math.floor(400 / data.length) - 8))
  const gap = 8
  const totalW = data.length * (barW + gap) - gap + 40

  return (
    <svg viewBox={`0 0 ${totalW} ${H + 36}`} className="w-full" style={{ maxHeight: 160 }}>
      {data.map((d, i) => {
        const barH = Math.max(2, (d.value / max) * H)
        const x = 20 + i * (barW + gap)
        const y = H - barH
        return (
          <g key={d.label}>
            <rect x={x} y={y} width={barW} height={barH}
              fill={d.color || '#3b82f6'} rx={3} opacity={0.85} />
            {d.value > 0 && (
              <text x={x + barW / 2} y={y - 3} textAnchor="middle"
                fontSize={9} fill="#374151" fontWeight="600">
                {d.value >= 1000 ? `${(d.value/1000).toFixed(1)}k` : d.value}
              </text>
            )}
            <text x={x + barW / 2} y={H + 14} textAnchor="middle"
              fontSize={8} fill="#9ca3af">
              {d.label}
            </text>
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
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(() => {
    const now = new Date()
    return new Set([`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`])
  })
  const [itemSearch, setItemSearch] = useState('')
  const [selectedItem, setSelectedItem] = useState<string | null>(null)
  const [activeTab, setActiveTab] = useState<'supplier' | 'item'>('supplier')

  const allMonthKeys = useMemo(() => generateMonthKeys(18), [])

  useEffect(() => {
    async function load() {
      const [invRes, poRes] = await Promise.all([
        supabase.from('invoices')
          .select('id, invoice_no, supplier, vendor_code, bl_date, estimated_arrival, currency, rows')
          .order('bl_date', { ascending: false }),
        supabase.from('po_items')
          .select('item_code, supplier, fob_price, currency')
      ])
      if (invRes.data) setInvoices(invRes.data as InvoiceData[])
      if (poRes.data) setPoItems(poRes.data as PoItem[])
      setLoading(false)
    }
    load()
  }, [])

  // Build price lookup: "item_code|supplier" → fob_price
  const priceMap = useMemo(() => {
    const m = new Map<string, { price: number | null; currency: string | null }>()
    for (const p of poItems) {
      m.set(`${p.item_code}|${p.supplier}`, { price: p.fob_price, currency: p.currency })
    }
    return m
  }, [poItems])

  // Build vendor_code lookup: supplier → vendor_code (from any invoice that has it)
  const vendorCodeMap = useMemo(() => {
    const m = new Map<string, string>()
    for (const inv of invoices) {
      if (inv.supplier && inv.vendor_code) m.set(inv.supplier, inv.vendor_code)
    }
    return m
  }, [invoices])

  // Flatten all invoices into line items
  const allLines = useMemo<LineItem[]>(() => {
    const lines: LineItem[] = []
    for (const inv of invoices) {
      if (!inv.rows) continue
      const supplier = inv.supplier || '—'
      const refDate = inv.bl_date || inv.estimated_arrival
      const mk = mKey(refDate)
      // Use invoice's own vendor_code, or fall back to lookup from same supplier
      const resolvedVendorCode = inv.vendor_code || vendorCodeMap.get(supplier) || null
      for (const row of inv.rows) {
        if (!row.code) continue
        const lookup = priceMap.get(`${row.code}|${supplier}`)
        const fob = lookup?.price ?? null
        lines.push({
          item_code: row.code,
          description: row.description || '',
          qty: row.qty || 0,
          po: row.po || '',
          invoice_no: inv.invoice_no,
          invoice_id: inv.id,
          supplier,
          vendor_code: resolvedVendorCode,
          bl_date: inv.bl_date,
          month_key: mk,
          currency: lookup?.currency ?? inv.currency,
          fob_price: fob,
          fob_total: fob != null && row.qty ? fob * row.qty : null,
        })
      }
    }
    return lines
  }, [invoices, priceMap])

  // Filter by selected months (null month_key = no date, always include if "All" selected)
  const filteredLines = useMemo(() => {
    if (selectedMonths.size === 0) return allLines
    return allLines.filter(l => l.month_key && selectedMonths.has(l.month_key))
  }, [allLines, selectedMonths])

  // Supplier summary
  const supplierSummary = useMemo(() => {
    const map = new Map<string, { qty: number; fob: number; items: Set<string>; invoices: Set<string> }>()
    for (const l of filteredLines) {
      const cur = map.get(l.supplier) || { qty: 0, fob: 0, items: new Set(), invoices: new Set() }
      cur.qty += l.qty
      if (l.fob_total) cur.fob += l.fob_total
      cur.items.add(l.item_code)
      cur.invoices.add(l.invoice_no)
      map.set(l.supplier, cur)
    }
    return Array.from(map.entries())
      .map(([supplier, v]) => ({ supplier, ...v, itemCount: v.items.size, invoiceCount: v.invoices.size }))
      .sort((a, b) => b.qty - a.qty)
  }, [filteredLines])

  // Item summary
  const itemSummary = useMemo(() => {
    const map = new Map<string, {
      description: string; totalQty: number; totalFob: number;
      suppliers: Map<string, { qty: number; fob: number }>
      byMonth: Map<string, number>
    }>()
    for (const l of filteredLines) {
      const cur = map.get(l.item_code) || {
        description: l.description, totalQty: 0, totalFob: 0,
        suppliers: new Map(), byMonth: new Map(),
      }
      cur.totalQty += l.qty
      if (l.fob_total) cur.totalFob += l.fob_total
      const sup = cur.suppliers.get(l.supplier) || { qty: 0, fob: 0 }
      sup.qty += l.qty; if (l.fob_total) sup.fob += l.fob_total
      cur.suppliers.set(l.supplier, sup)
      if (l.month_key) {
        cur.byMonth.set(l.month_key, (cur.byMonth.get(l.month_key) || 0) + l.qty)
      }
      map.set(l.item_code, cur)
    }
    return Array.from(map.entries())
      .map(([code, v]) => ({ code, ...v }))
      .sort((a, b) => b.totalQty - a.totalQty)
  }, [filteredLines])

  const filteredItemSummary = useMemo(() => {
    if (!itemSearch.trim()) return itemSummary
    const q = itemSearch.toLowerCase()
    return itemSummary.filter(i =>
      i.code.toLowerCase().includes(q) || i.description.toLowerCase().includes(q)
    )
  }, [itemSummary, itemSearch])

  const selectedItemData = useMemo(() =>
    selectedItem ? itemSummary.find(i => i.code === selectedItem) : null,
    [selectedItem, itemSummary])

  const toggleMonth = useCallback((k: string) => {
    setSelectedMonths(prev => {
      const next = new Set(prev)
      if (next.has(k)) next.delete(k); else next.add(k)
      return next
    })
  }, [])

  function exportExcel() {
    const lines = selectedItem
      ? filteredLines.filter(l => l.item_code === selectedItem)
      : itemSearch.trim()
        ? filteredLines.filter(l =>
            l.item_code.toLowerCase().includes(itemSearch.toLowerCase()) ||
            l.description.toLowerCase().includes(itemSearch.toLowerCase()))
        : filteredLines

    const rows = lines.map(l => ({
      'Item no': l.item_code,
      'Item name': l.description,
      'Supplier': l.supplier,
      'Vendor Code': l.vendor_code || '',
      'Quantity': l.qty,
      'Invoice': l.invoice_no,
      'PO': l.po,
      'FOB Unit Price': l.fob_price ?? '',
      'FOB Total Price': l.fob_total ?? '',
      'Original Currency': l.currency || '',
    }))

    const ws = XLSX.utils.json_to_sheet(rows)
    ws['!cols'] = [
      { wch: 22 }, { wch: 50 }, { wch: 18 }, { wch: 16 },
      { wch: 10 }, { wch: 18 }, { wch: 16 }, { wch: 15 }, { wch: 15 }, { wch: 16 },
    ]
    const wb = XLSX.utils.book_new()
    const sheetName = selectedItem
      ? `Item_${selectedItem.slice(0, 20)}`
      : `Summary_${Array.from(selectedMonths).sort().join('_').slice(0, 20)}`
    XLSX.utils.book_append_sheet(wb, ws, sheetName)
    XLSX.writeFile(wb, `item-summary-${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  const SUPPLIER_COLORS = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#f97316','#ec4899']

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Nav */}
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4 sticky top-0 z-20 shadow-sm">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/calendar" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Calendar</Link>
        <Link href="/report" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/summary" className="text-sm text-blue-600 font-semibold">Summary</Link>
        <Link href="/compare" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Builder</Link>
        <Link href="/guide" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Guide</Link>
        <div className="ml-auto"><LockButton /></div>
      </nav>

      <div className="max-w-screen-2xl mx-auto px-6 py-6">
        {/* Header */}
        <div className="flex items-start justify-between gap-4 mb-5 flex-wrap">
          <div>
            <h1 className="text-2xl font-black text-gray-900">Item Summary</h1>
            <p className="text-sm text-gray-400 mt-0.5">สรุปยอดสั่งซื้อแยกตาม Supplier และ Item</p>
          </div>
          <button
            onClick={exportExcel}
            className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white text-sm font-semibold rounded-xl hover:bg-green-700 shadow-sm transition"
          >
            ↓ Export Excel
          </button>
        </div>

        {/* Month picker */}
        <div className="bg-white border border-gray-200 rounded-2xl px-5 py-4 mb-5 shadow-sm">
          <div className="flex items-center justify-between mb-3">
            <span className="text-xs font-black text-gray-400 uppercase tracking-widest">เลือกเดือน (BL Date)</span>
            <div className="flex gap-2">
              <button
                onClick={() => setSelectedMonths(new Set(allMonthKeys))}
                className="text-xs text-blue-600 hover:underline"
              >ทั้งหมด</button>
              <span className="text-gray-300">|</span>
              <button
                onClick={() => setSelectedMonths(new Set())}
                className="text-xs text-gray-400 hover:underline"
              >ล้าง</button>
            </div>
          </div>
          <div className="flex flex-wrap gap-2">
            {allMonthKeys.map(k => (
              <button
                key={k}
                onClick={() => toggleMonth(k)}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all ${
                  selectedMonths.has(k)
                    ? 'bg-blue-600 text-white border-blue-600 shadow-sm'
                    : 'bg-white text-gray-500 border-gray-200 hover:border-blue-300 hover:text-blue-600'
                }`}
              >
                {mLabel(k)}
              </button>
            ))}
          </div>
          {selectedMonths.size > 0 && (
            <p className="text-xs text-gray-400 mt-2">
              เลือก {selectedMonths.size} เดือน · {filteredLines.length.toLocaleString()} รายการ · {[...new Set(filteredLines.map(l => l.invoice_no))].length} Invoice
            </p>
          )}
        </div>

        {loading ? (
          <div className="text-center py-20 text-gray-400">กำลังโหลด...</div>
        ) : (
          <div className="flex gap-5">
            {/* ─── Left panel ─── */}
            <div className="w-72 shrink-0 flex flex-col gap-4">
              {/* Tab toggle */}
              <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
                <div className="flex border-b border-gray-100">
                  {(['supplier', 'item'] as const).map(t => (
                    <button
                      key={t}
                      onClick={() => setActiveTab(t)}
                      className={`flex-1 py-2.5 text-xs font-bold transition ${
                        activeTab === t ? 'bg-blue-50 text-blue-700 border-b-2 border-blue-600' : 'text-gray-400 hover:text-gray-700'
                      }`}
                    >
                      {t === 'supplier' ? 'By Supplier' : 'By Item'}
                    </button>
                  ))}
                </div>

                {activeTab === 'supplier' ? (
                  <div className="divide-y divide-gray-50 max-h-[560px] overflow-y-auto">
                    {supplierSummary.length === 0 ? (
                      <p className="text-xs text-gray-400 p-4 text-center">ไม่มีข้อมูล</p>
                    ) : supplierSummary.map((s, si) => (
                      <div key={s.supplier} className="px-4 py-3">
                        <div className="flex items-start justify-between gap-2">
                          <div className="flex items-center gap-2 min-w-0">
                            <div className="w-2.5 h-2.5 rounded-full shrink-0"
                              style={{ background: SUPPLIER_COLORS[si % SUPPLIER_COLORS.length] }} />
                            <span className="text-sm font-bold text-gray-800 truncate">{s.supplier}</span>
                          </div>
                          <span className="text-xs text-gray-400 shrink-0">{s.invoiceCount} inv</span>
                        </div>
                        <div className="mt-1.5 flex items-center gap-3 text-xs text-gray-500 pl-4">
                          <span><strong className="text-gray-800">{s.qty.toLocaleString()}</strong> pcs</span>
                          <span>{s.itemCount} items</span>
                          {s.fob > 0 && <span className="text-emerald-600 font-semibold">{fmt(s.fob, 0)}</span>}
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="p-3 flex flex-col gap-2">
                    <input
                      type="text"
                      placeholder="ค้นหา Item Code / ชื่อ..."
                      value={itemSearch}
                      onChange={e => setItemSearch(e.target.value)}
                      className="w-full border border-gray-200 rounded-lg px-3 py-1.5 text-xs outline-none focus:border-blue-400"
                    />
                    <div className="max-h-[510px] overflow-y-auto divide-y divide-gray-50">
                      {filteredItemSummary.slice(0, 100).map(item => (
                        <button
                          key={item.code}
                          onClick={() => setSelectedItem(selectedItem === item.code ? null : item.code)}
                          className={`w-full text-left px-2 py-2.5 rounded-lg transition-all ${
                            selectedItem === item.code
                              ? 'bg-blue-50 border border-blue-200'
                              : 'hover:bg-gray-50'
                          }`}
                        >
                          <div className="text-xs font-bold text-gray-800 font-mono truncate">{item.code}</div>
                          <div className="text-[10px] text-gray-400 truncate mt-0.5">{item.description}</div>
                          <div className="text-xs text-blue-700 font-semibold mt-0.5">{item.totalQty.toLocaleString()} pcs</div>
                        </button>
                      ))}
                      {filteredItemSummary.length === 0 && (
                        <p className="text-xs text-gray-400 p-3 text-center">ไม่พบ item</p>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>

            {/* ─── Main content ─── */}
            <div className="flex-1 min-w-0 flex flex-col gap-4">

              {/* Item detail panel (when item selected) */}
              {selectedItemData && (
                <div className="bg-white border border-blue-200 rounded-2xl shadow-sm p-5">
                  <div className="flex items-start justify-between mb-4">
                    <div>
                      <div className="text-xs font-black text-blue-500 uppercase tracking-widest mb-1">Item Detail</div>
                      <h2 className="text-lg font-black text-gray-900 font-mono">{selectedItemData.code}</h2>
                      <p className="text-sm text-gray-500 mt-0.5">{selectedItemData.description}</p>
                    </div>
                    <button onClick={() => setSelectedItem(null)}
                      className="text-gray-300 hover:text-gray-600 text-xl leading-none">✕</button>
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    {/* Supplier breakdown */}
                    <div>
                      <p className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-2">ยอดแยก Supplier</p>
                      <div className="space-y-2">
                        {Array.from(selectedItemData.suppliers.entries())
                          .sort((a, b) => b[1].qty - a[1].qty)
                          .map(([sup, v], si) => {
                            const pct = selectedItemData.totalQty > 0 ? (v.qty / selectedItemData.totalQty) * 100 : 0
                            return (
                              <div key={sup}>
                                <div className="flex items-center justify-between text-xs mb-0.5">
                                  <span className="font-semibold text-gray-700">{sup}</span>
                                  <span className="text-gray-500">{v.qty.toLocaleString()} pcs · {pct.toFixed(0)}%</span>
                                </div>
                                <div className="h-1.5 bg-gray-100 rounded-full overflow-hidden">
                                  <div className="h-full rounded-full"
                                    style={{ width: `${pct}%`, background: SUPPLIER_COLORS[si % SUPPLIER_COLORS.length] }} />
                                </div>
                              </div>
                            )
                          })}
                      </div>
                      <div className="mt-3 pt-2 border-t border-gray-100 text-xs text-gray-500">
                        รวม <strong className="text-gray-800">{selectedItemData.totalQty.toLocaleString()}</strong> pcs
                        {selectedItemData.totalFob > 0 && (
                          <> · FOB <strong className="text-emerald-600">{fmt(selectedItemData.totalFob, 0)}</strong></>
                        )}
                      </div>
                    </div>

                    {/* Monthly chart */}
                    <div>
                      <p className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-2">ยอดสั่งรายเดือน (pcs)</p>
                      {(() => {
                        const chartMonths = generateMonthKeys(12)
                        const chartData = chartMonths.map(k => ({
                          label: mLabel(k).slice(0, 3),
                          value: selectedItemData.byMonth.get(k) || 0,
                        }))
                        return <BarChart data={chartData} />
                      })()}
                    </div>
                  </div>
                </div>
              )}

              {/* Lines table */}
              <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
                <div className="px-5 py-3 border-b border-gray-100 flex items-center justify-between">
                  <span className="text-sm font-bold text-gray-700">
                    {selectedItem ? `รายการของ ${selectedItem}` : 'รายการทั้งหมด'}
                    <span className="ml-2 text-xs font-normal text-gray-400">
                      {(selectedItem
                        ? filteredLines.filter(l => l.item_code === selectedItem)
                        : filteredLines
                      ).length.toLocaleString()} rows
                    </span>
                  </span>
                  {selectedItem && (
                    <button onClick={() => setSelectedItem(null)}
                      className="text-xs text-blue-500 hover:underline">ดูทั้งหมด</button>
                  )}
                </div>
                <div className="overflow-x-auto">
                  <table className="text-xs w-full border-collapse">
                    <thead>
                      <tr className="bg-gray-50 text-gray-500 text-left">
                        {['Item no','Item name','Supplier','Vendor Code','QTY','Invoice','PO','FOB Unit','FOB Total','Currency'].map(h => (
                          <th key={h} className="px-3 py-2.5 font-bold border-b border-gray-100 whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {(selectedItem
                        ? filteredLines.filter(l => l.item_code === selectedItem)
                        : filteredLines
                      ).slice(0, 500).map((l, i) => (
                        <tr key={i}
                          className={`border-b border-gray-50 hover:bg-blue-50/40 cursor-pointer transition-colors ${
                            selectedItem === l.item_code ? 'bg-blue-50/20' : ''
                          }`}
                          onClick={() => setSelectedItem(l.item_code)}
                        >
                          <td className="px-3 py-2 font-mono text-gray-800 whitespace-nowrap">{l.item_code}</td>
                          <td className="px-3 py-2 text-gray-600 max-w-[200px] truncate" title={l.description}>{l.description}</td>
                          <td className="px-3 py-2 text-gray-700 whitespace-nowrap">{l.supplier}</td>
                          <td className="px-3 py-2 font-mono text-gray-500">{l.vendor_code || '—'}</td>
                          <td className="px-3 py-2 text-right font-semibold text-gray-800">{l.qty.toLocaleString()}</td>
                          <td className="px-3 py-2 whitespace-nowrap">
                            <Link href={`/dashboard/${l.invoice_id}`}
                              onClick={e => e.stopPropagation()}
                              className="text-blue-600 hover:underline">
                              {l.invoice_no}
                            </Link>
                          </td>
                          <td className="px-3 py-2 text-gray-500 whitespace-nowrap">{l.po || '—'}</td>
                          <td className="px-3 py-2 text-right text-gray-700">
                            {l.fob_price != null ? fmt(l.fob_price) : <span className="text-gray-300">—</span>}
                          </td>
                          <td className="px-3 py-2 text-right font-semibold text-emerald-700">
                            {l.fob_total != null ? fmt(l.fob_total, 0) : <span className="text-gray-300 font-normal">—</span>}
                          </td>
                          <td className="px-3 py-2 text-gray-400">{l.currency || '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filteredLines.length > 500 && !selectedItem && (
                    <p className="text-xs text-gray-400 text-center py-3">
                      แสดง 500 รายการแรก จาก {filteredLines.length.toLocaleString()} — เลือก Item หรือ filter เดือนเพื่อดูทั้งหมด
                    </p>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

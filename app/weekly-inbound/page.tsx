'use client'

import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import NavBar from '@/components/NavBar'
import { supplierColor } from '@/lib/supplier-colors'
import { HUB_COLORS, hubColor } from '@/lib/hub-colors'
import type { ResultRow } from '@/lib/excel-parser'

const DEFAULT_HUB = 'มัยลาภ'
const ALL_HUBS = Object.keys(HUB_COLORS) // fixed, known set — matches /calendar's hub filter

// ── Date helpers (Monday-first week, same convention as /calendar) ─────────
function dow(d: Date) { return (d.getDay() + 6) % 7 } // Mon=0 … Sun=6
function ds(d: Date) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}` }
function pd(s: string) { return new Date(s + 'T00:00:00') }
function addDays(d: Date, n: number) { const r = new Date(d); r.setDate(r.getDate() + n); return r }
function mondayOf(d: Date) { return addDays(d, -dow(d)) }

function getWeekNum(date: Date): number {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()))
  const day = d.getUTCDay() || 7
  d.setUTCDate(d.getUTCDate() + 4 - day)
  const y = new Date(Date.UTC(d.getUTCFullYear(), 0, 1))
  return Math.ceil(((d.getTime() - y.getTime()) / 86400000 + 1) / 7)
}

const MONTHS_TH_SHORT = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.']
const DAY_TH_SHORT = ['จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.', 'อา.']

function formatWeekLabel(mon: Date, sun: Date): string {
  const wn = getWeekNum(mon)
  const yearBE = sun.getFullYear() + 543
  const sameMonth = mon.getMonth() === sun.getMonth()
  const range = sameMonth
    ? `${mon.getDate()}–${sun.getDate()} ${MONTHS_TH_SHORT[mon.getMonth()]}`
    : `${mon.getDate()} ${MONTHS_TH_SHORT[mon.getMonth()]} – ${sun.getDate()} ${MONTHS_TH_SHORT[sun.getMonth()]}`
  return `สัปดาห์ที่ ${wn} | ${range} ${yearBE}`
}

function formatEtaRange(startStr: string, endStr: string): string {
  const s = pd(startStr); const e = pd(endStr)
  if (startStr === endStr) return `${s.getDate()} ${MONTHS_TH_SHORT[s.getMonth()]}`
  if (s.getMonth() === e.getMonth()) return `${s.getDate()}–${e.getDate()} ${MONTHS_TH_SHORT[s.getMonth()]}`
  return `${s.getDate()} ${MONTHS_TH_SHORT[s.getMonth()]} – ${e.getDate()} ${MONTHS_TH_SHORT[e.getMonth()]}`
}

// ── Data shapes ──────────────────────────────────────────────────────────
interface InvoiceRow {
  id: string
  invoice_no: string
  supplier: string | null
  estimated_arrival: string | null
  estimated_arrival_end: string | null
  container_names: string[] | null
  rows: ResultRow[] | null
}

interface InvoiceBreakdown {
  invoiceNo: string
  supplier: string | null
  qty: number
}

interface AggItem {
  code: string
  description: string
  qty: number
  etaStart: string
  etaEnd: string
  suppliers: string[]
  bySupplier: Record<string, number>
  byInvoice: Record<string, InvoiceBreakdown> // keyed by invoice id
  hubs: string[]
}

type SortKey = 'code' | 'description' | 'qty' | 'eta'

const PAGE_SIZE = 15

interface ContainerHubInfo {
  hub: string
  date: string
}

// item_code → project lookup: Cost Compare (po_items) is the primary source
// (latest upload per item_code wins); when an item_code isn't in po_items,
// fall back to the PO Insights upload history (po_uploads.rows), also
// latest-wins.
function buildProjectMap(
  poItems: { item_code: string; project: string; uploaded_at: string }[],
  poUploads: { project: string; rows: { item_code: string }[] | null; created_at: string }[],
): Map<string, string> {
  const map = new Map<string, string>()
  const sortedItems = [...poItems].sort((a, b) => b.uploaded_at.localeCompare(a.uploaded_at))
  for (const it of sortedItems) {
    if (!map.has(it.item_code)) map.set(it.item_code, it.project)
  }
  const sortedUploads = [...poUploads].sort((a, b) => b.created_at.localeCompare(a.created_at))
  for (const up of sortedUploads) {
    for (const r of up.rows ?? []) {
      if (!map.has(r.item_code)) map.set(r.item_code, up.project)
    }
  }
  return map
}

// Aggregates invoice/container/item data into per-item rows for the given
// week. supplierFilter/hubFilter/projectFilter narrow the SOURCE
// containers/rows that get summed — so picking a hub (or project) shows
// only the qty actually matching it, not the item's total across everything.
function buildAggItems(
  invoices: InvoiceRow[],
  hubArrival: Map<string, ContainerHubInfo>,
  projectMap: Map<string, string>,
  wMonStr: string,
  wSunStr: string,
  supplierFilter: string,
  hubFilter: string,
  projectFilter: string,
): AggItem[] {
  const map = new Map<string, AggItem>()

  for (const inv of invoices) {
    if (supplierFilter && inv.supplier !== supplierFilter) continue
    const containerNames = inv.container_names ?? []
    if (containerNames.length === 0 || !inv.rows) continue

    for (const containerName of containerNames) {
      const hubInfo = hubArrival.get(`${inv.id}::${containerName}`)
      const hub = hubInfo?.hub ?? DEFAULT_HUB
      if (hubFilter && hub !== hubFilter) continue

      let etaStart: string; let etaEnd: string
      if (hubInfo) {
        etaStart = hubInfo.date; etaEnd = hubInfo.date
      } else {
        if (!inv.estimated_arrival) continue
        etaStart = inv.estimated_arrival
        etaEnd = inv.estimated_arrival_end || inv.estimated_arrival
      }
      // Skip containers whose ETA range doesn't touch the selected week
      if (etaEnd < wMonStr || etaStart > wSunStr) continue
      const dispStart = etaStart < wMonStr ? wMonStr : etaStart
      const dispEnd = etaEnd > wSunStr ? wSunStr : etaEnd

      for (const row of inv.rows) {
        if (projectFilter && (projectMap.get(row.code) ?? '') !== projectFilter) continue
        const qty = row.containers?.[containerName] || 0
        if (qty <= 0) continue
        let agg = map.get(row.code)
        if (!agg) {
          agg = { code: row.code, description: row.description || '', qty: 0, etaStart: dispStart, etaEnd: dispEnd, suppliers: [], bySupplier: {}, byInvoice: {}, hubs: [] }
          map.set(row.code, agg)
        }
        agg.qty += qty
        if (dispStart < agg.etaStart) agg.etaStart = dispStart
        if (dispEnd > agg.etaEnd) agg.etaEnd = dispEnd
        if (!agg.description && row.description) agg.description = row.description
        if (!agg.hubs.includes(hub)) agg.hubs.push(hub)
        if (inv.supplier) {
          if (!agg.suppliers.includes(inv.supplier)) agg.suppliers.push(inv.supplier)
          agg.bySupplier[inv.supplier] = (agg.bySupplier[inv.supplier] ?? 0) + qty
        }
        if (!agg.byInvoice[inv.id]) agg.byInvoice[inv.id] = { invoiceNo: inv.invoice_no, supplier: inv.supplier, qty: 0 }
        agg.byInvoice[inv.id].qty += qty
      }
    }
  }
  return Array.from(map.values())
}

export default function WeeklyInboundPlanPage() {
  const [invoices, setInvoices] = useState<InvoiceRow[]>([])
  const [hubArrival, setHubArrival] = useState<Map<string, ContainerHubInfo>>(new Map()) // `${invoice_id}::${container_name}` -> {hub, date}
  const [projectMap, setProjectMap] = useState<Map<string, string>>(new Map()) // item_code -> project
  const [loading, setLoading] = useState(true)

  const [weekMon, setWeekMon] = useState(() => mondayOf(new Date()))
  const [supplierFilter, setSupplierFilter] = useState('')
  const [hubFilter, setHubFilter] = useState('')
  const [projectFilter, setProjectFilter] = useState('')
  const [sortKey, setSortKey] = useState<SortKey>('eta')
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('asc')
  const [page, setPage] = useState(1)
  const [detailItem, setDetailItem] = useState<AggItem | null>(null)

  const weekSun = useMemo(() => addDays(weekMon, 6), [weekMon])

  async function load() {
    setLoading(true)
    const [{ data: invData }, { data: hubData }, { data: poItemsData }, { data: poUploadsData }] = await Promise.all([
      supabase.from('invoices').select('id, invoice_no, supplier, estimated_arrival, estimated_arrival_end, container_names, rows'),
      supabase.from('container_hub_requests').select('invoice_id, container_name, hub, hub_arrival_date').eq('status', 'confirmed').not('hub_arrival_date', 'is', null),
      supabase.from('po_items').select('item_code, project, uploaded_at'),
      supabase.from('po_uploads').select('project, rows, created_at'),
    ])
    if (invData) setInvoices(invData as InvoiceRow[])
    if (hubData) {
      const m = new Map<string, ContainerHubInfo>()
      for (const h of hubData as { invoice_id: string; container_name: string; hub: string; hub_arrival_date: string }[]) {
        m.set(`${h.invoice_id}::${h.container_name}`, { hub: h.hub, date: h.hub_arrival_date })
      }
      setHubArrival(m)
    }
    setProjectMap(buildProjectMap(
      (poItemsData ?? []) as { item_code: string; project: string; uploaded_at: string }[],
      (poUploadsData ?? []) as { project: string; rows: { item_code: string }[] | null; created_at: string }[],
    ))
    setLoading(false)
  }

  useEffect(() => { load() }, [])

  // Reset to page 1 whenever the week or filters change — adjusted during
  // render (React's recommended pattern) rather than in an effect, so it
  // doesn't trigger an extra render pass.
  const filterKey = `${ds(weekMon)}|${supplierFilter}|${hubFilter}|${projectFilter}`
  const [prevFilterKey, setPrevFilterKey] = useState(filterKey)
  if (filterKey !== prevFilterKey) {
    setPrevFilterKey(filterKey)
    setPage(1)
  }

  // ── Aggregate items arriving within the selected week ────────────────────
  const wMonStr = useMemo(() => ds(weekMon), [weekMon])
  const wSunStr = useMemo(() => ds(weekSun), [weekSun])

  // Unfiltered, so the Supplier dropdown always lists every supplier
  // shipping this week regardless of the currently-selected filters.
  const baseItems = useMemo(
    () => buildAggItems(invoices, hubArrival, projectMap, wMonStr, wSunStr, '', '', ''),
    [invoices, hubArrival, projectMap, wMonStr, wSunStr]
  )
  const allSuppliers = useMemo(() => {
    const s = new Set<string>()
    baseItems.forEach(i => i.suppliers.forEach(sup => s.add(sup)))
    return [...s].sort()
  }, [baseItems])
  const allProjects = useMemo(() => [...new Set(projectMap.values())].sort(), [projectMap])

  // Filtered: when a hub/supplier/project is selected, quantities are
  // summed only from the matching containers/invoices/items — e.g. picking
  // "ขอนแก่น" shows just the qty actually going to that hub, not the
  // item's grand total.
  const filteredItems = useMemo(
    () => buildAggItems(invoices, hubArrival, projectMap, wMonStr, wSunStr, supplierFilter, hubFilter, projectFilter),
    [invoices, hubArrival, projectMap, wMonStr, wSunStr, supplierFilter, hubFilter, projectFilter]
  )

  const sortedItems = useMemo(() => {
    const arr = [...filteredItems]
    const dir = sortDir === 'asc' ? 1 : -1
    arr.sort((a, b) => {
      switch (sortKey) {
        case 'code': return a.code.localeCompare(b.code) * dir
        case 'description': return (a.description || '').localeCompare(b.description || '') * dir
        case 'qty': return (a.qty - b.qty) * dir
        case 'eta': return (a.etaStart === b.etaStart ? a.code.localeCompare(b.code) : a.etaStart.localeCompare(b.etaStart)) * dir
        default: return 0
      }
    })
    return arr
  }, [filteredItems, sortKey, sortDir])

  const totalPages = Math.max(1, Math.ceil(sortedItems.length / PAGE_SIZE))
  const pagedItems = sortedItems.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE)

  function toggleSort(key: SortKey) {
    if (sortKey === key) setSortDir(d => d === 'asc' ? 'desc' : 'asc')
    else { setSortKey(key); setSortDir('asc') }
  }

  // ── Summary by supplier (top 4 + Others) ──────────────────────────────
  const supplierSummary = useMemo(() => {
    const counts = new Map<string, number>()
    filteredItems.forEach(i => i.suppliers.forEach(s => counts.set(s, (counts.get(s) ?? 0) + 1)))
    const sorted = [...counts.entries()].sort((a, b) => b[1] - a[1])
    const top = sorted.slice(0, 4)
    const othersCount = sorted.slice(4).reduce((s, [, c]) => s + c, 0)
    const rows = top.map(([name, count], idx) => ({ name, count, color: supplierColor(name, idx).dot }))
    rows.push({ name: 'Others', count: othersCount, color: '#cbd5e1' })
    return rows
  }, [filteredItems])
  const supplierSummaryTotal = supplierSummary.reduce((s, r) => s + r.count, 0)

  // ── ETA by date (Mon..Sun) ─────────────────────────────────────────────
  const etaByDate = useMemo(() => {
    const days = Array.from({ length: 7 }, (_, i) => addDays(weekMon, i))
    return days.map(d => {
      const dStr = ds(d)
      const count = filteredItems.filter(i => i.etaStart <= dStr && i.etaEnd >= dStr).length
      return { date: d, count }
    })
  }, [weekMon, filteredItems])
  const maxEtaCount = Math.max(1, ...etaByDate.map(e => e.count))

  function exportExcel() {
    const header = ['Item Code', 'Description', 'จำนวนรวม (pcs)', 'กำหนดเข้าคลัง', 'Supplier', 'Hub', 'Project']
    const body = sortedItems.map(i => [
      i.code, i.description || '-', i.qty, formatEtaRange(i.etaStart, i.etaEnd),
      i.suppliers.join(', '), i.hubs.join(', '), projectMap.get(i.code) || '-',
    ])
    const ws = XLSX.utils.aoa_to_sheet([header, ...body])
    ws['!cols'] = [{ wch: 20 }, { wch: 36 }, { wch: 14 }, { wch: 16 }, { wch: 24 }, { wch: 20 }, { wch: 16 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Weekly Inbound Plan')
    XLSX.writeFile(wb, `Weekly_Inbound_Plan_W${getWeekNum(weekMon)}_${ds(weekMon)}.xlsx`)
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <NavBar />

      <div className="max-w-7xl mx-auto px-6 py-8">
        {/* Header */}
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">Weekly Inbound Plan</h1>
          <p className="text-sm text-gray-500 mt-1">
            ภาพรวมสินค้าที่จะเข้าคลังในแต่ละสัปดาห์ เพื่อให้ทีมคลังสามารถวางแผนการรับสินค้าได้ล่วงหน้า
          </p>
        </div>

        {/* Controls */}
        <div className="flex items-center justify-between gap-3 flex-wrap mb-6">
          <div className="flex items-center gap-2 bg-white border border-gray-200 rounded-lg shadow-sm px-2 py-1.5">
            <button onClick={() => setWeekMon(w => addDays(w, -7))} className="w-7 h-7 flex items-center justify-center rounded-md text-gray-500 hover:bg-gray-50">‹</button>
            <span className="text-sm font-semibold text-gray-800 px-2 whitespace-nowrap">📅 {formatWeekLabel(weekMon, weekSun)}</span>
            <button onClick={() => setWeekMon(w => addDays(w, 7))} className="w-7 h-7 flex items-center justify-center rounded-md text-gray-500 hover:bg-gray-50">›</button>
            <button onClick={() => setWeekMon(mondayOf(new Date()))} className="ml-1 px-3 py-1 text-xs font-semibold rounded-md bg-blue-600 text-white hover:bg-blue-700">วันนี้</button>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <select value={supplierFilter} onChange={e => setSupplierFilter(e.target.value)}
              className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 bg-white">
              <option value="">All Suppliers</option>
              {allSuppliers.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
            <select value={projectFilter} onChange={e => setProjectFilter(e.target.value)}
              className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 bg-white">
              <option value="">All Projects</option>
              {allProjects.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
            <select value={hubFilter} onChange={e => setHubFilter(e.target.value)}
              className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 bg-white">
              <option value="">All Hubs</option>
              {ALL_HUBS.map(h => <option key={h} value={h}>{h}</option>)}
            </select>
            <button onClick={exportExcel} disabled={sortedItems.length === 0}
              className="flex items-center gap-1.5 px-4 py-1.5 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700 disabled:opacity-40 transition-colors">
              ↓ Export
            </button>
          </div>
        </div>

        {loading ? (
          <p className="text-sm text-gray-400">กำลังโหลด...</p>
        ) : (
          <>
            {/* Summary cards */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-6">
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5 flex items-center gap-3">
                <span className="text-2xl">📦</span>
                <div>
                  <p className="text-xs text-gray-500">จำนวนรายการสินค้า</p>
                  <p className="text-2xl font-bold text-gray-900">{filteredItems.length} <span className="text-sm font-normal text-gray-400">items</span></p>
                </div>
              </div>
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5 flex items-center gap-3">
                <span className="text-2xl">📆</span>
                <div>
                  <p className="text-xs text-gray-500">ช่วงวันที่ ETA</p>
                  <p className="text-base font-bold text-gray-900">{formatEtaRange(ds(weekMon), ds(weekSun))} {weekSun.getFullYear() + 543}</p>
                </div>
              </div>
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5 flex items-center gap-3">
                <span className="text-2xl">🏭</span>
                <div>
                  <p className="text-xs text-gray-500">จำนวน Supplier</p>
                  <p className="text-2xl font-bold text-gray-900">{allSuppliers.length} <span className="text-sm font-normal text-gray-400">suppliers</span></p>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-[1fr_320px] gap-6">
              {/* Main table */}
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
                {sortedItems.length === 0 ? (
                  <div className="text-center py-16 text-gray-400">
                    <p className="text-4xl mb-3">📭</p>
                    <p className="text-sm">ไม่มีสินค้าเข้าคลังในสัปดาห์นี้</p>
                  </div>
                ) : (
                  <>
                    <div className="overflow-x-auto">
                      <table className="w-full text-sm border-collapse">
                        <thead>
                          <tr className="bg-gray-50 text-gray-500 text-xs border-b border-gray-200">
                            <SortHeader label="Item Code" k="code" activeKey={sortKey} activeDir={sortDir} onToggle={toggleSort} />
                            <SortHeader label="Description" k="description" activeKey={sortKey} activeDir={sortDir} onToggle={toggleSort} />
                            <SortHeader label="จำนวนรวม (pcs)" k="qty" activeKey={sortKey} activeDir={sortDir} onToggle={toggleSort} align="right" />
                            <SortHeader label="กำหนดเข้าคลัง" k="eta" activeKey={sortKey} activeDir={sortDir} onToggle={toggleSort} />
                            <th className="px-4 py-3 text-left">Supplier</th>
                            <th className="px-4 py-3 text-left">Hub</th>
                            <th className="px-2 py-3 w-8"></th>
                          </tr>
                        </thead>
                        <tbody>
                          {pagedItems.map(item => (
                            <tr key={item.code} onClick={() => setDetailItem(item)} className="border-b border-gray-100 hover:bg-gray-50 cursor-pointer transition-colors">
                              <td className="px-4 py-3 font-mono text-xs font-semibold text-gray-800">{item.code}</td>
                              <td className="px-4 py-3 text-gray-600">{item.description || '-'}</td>
                              <td className="px-4 py-3 text-right font-semibold text-gray-800">{item.qty.toLocaleString()}</td>
                              <td className="px-4 py-3 text-gray-600 whitespace-nowrap">{formatEtaRange(item.etaStart, item.etaEnd)}</td>
                              <td className="px-4 py-3">
                                <div className="flex flex-wrap gap-1">
                                  {item.suppliers.map((s, idx) => {
                                    const c = supplierColor(s, idx)
                                    return <span key={s} className="text-[11px] font-semibold px-2 py-0.5 rounded-full" style={{ background: c.bg, color: c.text }}>{s}</span>
                                  })}
                                </div>
                              </td>
                              <td className="px-4 py-3">
                                <div className="flex flex-wrap gap-1">
                                  {item.hubs.map(h => {
                                    const c = hubColor(h)
                                    return <span key={h} className="text-[11px] font-semibold px-2 py-0.5 rounded-full" style={{ background: c.bg, color: c.text }}>{h}</span>
                                  })}
                                </div>
                              </td>
                              <td className="px-2 py-3 text-gray-300 text-xs">›</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {/* Pagination */}
                    <div className="flex items-center justify-between px-4 py-3 border-t border-gray-100 text-xs text-gray-500">
                      <span>แสดง {(page - 1) * PAGE_SIZE + 1}–{Math.min(page * PAGE_SIZE, sortedItems.length)} จาก {sortedItems.length} รายการ</span>
                      <div className="flex items-center gap-1.5">
                        <button onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page === 1}
                          className="w-7 h-7 flex items-center justify-center rounded-md border border-gray-200 disabled:opacity-30 hover:bg-gray-50">‹</button>
                        <span className="w-7 h-7 flex items-center justify-center rounded-md bg-blue-600 text-white font-semibold">{page}</span>
                        <button onClick={() => setPage(p => Math.min(totalPages, p + 1))} disabled={page === totalPages}
                          className="w-7 h-7 flex items-center justify-center rounded-md border border-gray-200 disabled:opacity-30 hover:bg-gray-50">›</button>
                      </div>
                    </div>
                  </>
                )}
              </div>

              {/* Right column */}
              <div className="flex flex-col gap-4">
                {/* Summary by Supplier */}
                <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5">
                  <h3 className="text-sm font-bold text-gray-800 mb-4">สรุปตาม Supplier</h3>
                  {supplierSummaryTotal === 0 ? (
                    <p className="text-xs text-gray-400">ไม่มีข้อมูล</p>
                  ) : (
                    <div className="flex items-center gap-4">
                      <Donut segments={supplierSummary} total={supplierSummaryTotal} centerLabel={String(filteredItems.length)} centerSub="items" />
                      <div className="flex-1 space-y-1.5 min-w-0">
                        {supplierSummary.filter(s => s.count > 0).map(s => (
                          <div key={s.name} className="flex items-center justify-between gap-2 text-xs">
                            <span className="flex items-center gap-1.5 min-w-0">
                              <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: s.color }} />
                              <span className="truncate text-gray-700 font-medium">{s.name}</span>
                            </span>
                            <span className="text-gray-500 whitespace-nowrap">{s.count} <span className="text-gray-300">({Math.round(s.count / supplierSummaryTotal * 100)}%)</span></span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* ETA by Date */}
                <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5">
                  <h3 className="text-sm font-bold text-gray-800 mb-4">กำหนดเข้าคลังตามวัน</h3>
                  <div className="space-y-2.5">
                    {etaByDate.map(({ date, count }) => (
                      <div key={ds(date)}>
                        <div className="flex items-center justify-between text-xs mb-1">
                          <span className="text-gray-600 font-medium">{date.getDate()} {MONTHS_TH_SHORT[date.getMonth()]} ({DAY_TH_SHORT[dow(date)]})</span>
                          <span className="text-gray-400">{count} items</span>
                        </div>
                        <div className="h-1.5 bg-gray-100 rounded-full overflow-hidden">
                          <div className="h-full bg-blue-500 rounded-full" style={{ width: `${(count / maxEtaCount) * 100}%` }} />
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Note */}
                <div className="bg-blue-50 border border-blue-100 rounded-xl p-4 text-xs text-blue-800 flex gap-2">
                  <span>💡</span>
                  <div>
                    <p className="font-semibold mb-0.5">หมายเหตุ</p>
                    <p className="text-blue-700">ยอดรวมในแต่ละรายการเป็นจำนวนรวมทั้งหมดในสัปดาห์นี้ (อาจมาจากหลาย Invoice)</p>
                  </div>
                </div>
              </div>
            </div>
          </>
        )}
      </div>

      {/* Item detail popup — breakdown by invoice */}
      {detailItem && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center z-50 p-4" onClick={() => setDetailItem(null)}>
          <div className="bg-white rounded-xl shadow-2xl p-6 max-w-sm w-full" onClick={e => e.stopPropagation()}>
            <div className="flex items-start justify-between gap-3 mb-1">
              <h3 className="font-mono font-bold text-gray-900">{detailItem.code}</h3>
              <button onClick={() => setDetailItem(null)} className="text-gray-300 hover:text-gray-500">✕</button>
            </div>
            <p className="text-sm text-gray-500 mb-4">{detailItem.description || '-'}</p>
            <p className="text-xs font-semibold text-gray-400 uppercase tracking-wide mb-2">แยกตาม Invoice</p>
            <div className="space-y-2 mb-4 max-h-64 overflow-y-auto">
              {Object.values(detailItem.byInvoice).sort((a, b) => b.qty - a.qty).map((inv, idx) => {
                const c = inv.supplier ? supplierColor(inv.supplier, idx) : null
                return (
                  <div key={inv.invoiceNo} className="flex items-center justify-between gap-2 text-sm">
                    <span className="flex items-center gap-1.5 min-w-0">
                      <span className="font-mono text-xs text-gray-700 truncate">{inv.invoiceNo}</span>
                      {inv.supplier && c && (
                        <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded-full shrink-0" style={{ background: c.bg, color: c.text }}>{inv.supplier}</span>
                      )}
                    </span>
                    <span className="font-semibold text-gray-700 whitespace-nowrap">{inv.qty.toLocaleString()} pcs</span>
                  </div>
                )
              })}
            </div>
            <div className="flex items-center justify-between pt-3 border-t border-gray-100 text-sm">
              <span className="text-gray-500">รวมทั้งหมด</span>
              <span className="font-bold text-gray-900">{detailItem.qty.toLocaleString()} pcs</span>
            </div>
            <div className="flex items-center justify-between mt-1 text-sm">
              <span className="text-gray-500">กำหนดเข้าคลัง</span>
              <span className="font-semibold text-gray-700">{formatEtaRange(detailItem.etaStart, detailItem.etaEnd)}</span>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

// ── Sort-toggling table header, hoisted out of the page component so it
// isn't recreated (and reset) on every render ────────────────────────────
function SortHeader({ label, k, activeKey, activeDir, onToggle, align = 'left' }: {
  label: string
  k: SortKey
  activeKey: SortKey
  activeDir: 'asc' | 'desc'
  onToggle: (k: SortKey) => void
  align?: 'left' | 'right'
}) {
  return (
    <th
      className={`px-4 py-3 ${align === 'right' ? 'text-right' : 'text-left'} cursor-pointer select-none hover:text-gray-800 whitespace-nowrap`}
      onClick={() => onToggle(k)}
    >
      <span className={`inline-flex items-center gap-1 ${align === 'right' ? 'justify-end w-full' : ''}`}>
        {label}
        <span className="text-[10px] text-gray-300">{activeKey === k ? (activeDir === 'asc' ? '▲' : '▼') : '↕'}</span>
      </span>
    </th>
  )
}

// ── Donut chart (plain SVG, no chart library) ───────────────────────────
function Donut({ segments, total, centerLabel, centerSub }: {
  segments: { name: string; count: number; color: string }[]
  total: number
  centerLabel: string
  centerSub: string
}) {
  const size = 120; const thickness = 16
  const r = (size - thickness) / 2
  const c = 2 * Math.PI * r
  const drawn = segments.filter(s => s.count > 0)
  const dashes = drawn.map(s => (total > 0 ? s.count / total : 0) * c)
  const arcs = drawn.map((s, i) => ({
    ...s,
    dash: dashes[i],
    offset: dashes.slice(0, i).reduce((sum, d) => sum + d, 0),
  }))
  return (
    <div className="relative shrink-0" style={{ width: size, height: size }}>
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
        <g transform={`rotate(-90 ${size / 2} ${size / 2})`}>
          <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke="#f1f5f9" strokeWidth={thickness} />
          {arcs.map(s => (
            <circle key={s.name} cx={size / 2} cy={size / 2} r={r} fill="none" stroke={s.color}
              strokeWidth={thickness} strokeDasharray={`${s.dash} ${c - s.dash}`} strokeDashoffset={-s.offset} strokeLinecap="butt" />
          ))}
        </g>
      </svg>
      <div className="absolute inset-0 flex flex-col items-center justify-center">
        <span className="text-lg font-bold text-gray-900 leading-none">{centerLabel}</span>
        <span className="text-[10px] text-gray-400 leading-none mt-0.5">{centerSub}</span>
      </div>
    </div>
  )
}

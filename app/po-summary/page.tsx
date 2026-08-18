'use client'

import { useCallback, useEffect, useState, useMemo } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import NavBar from '@/components/NavBar'
import { extractGroup } from '@/lib/po-group'

// ── Types ────────────────────────────────────────────────────────────────────
interface POItemRow {
  item_code: string
  description: string
  qty: number
  unit_price: number
  total: number
}

interface POItem {
  project: string
  supplier: string
  item_code: string
  description: string | null
  fob_price: number
  currency: string
}

interface DetailEntry {
  supplier: string
  qty: number
  unit_price: number
  currency: string
  unit_thb: number
  total_thb: number
  poId: string
  poLabel: string
}

interface GroupDetailItem {
  item_code: string
  description: string | null
  supplierEntries: DetailEntry[]
  total_thb: number
}

interface POUpload {
  id: string
  supplier: string
  project: string
  currency: string
  total_amount: number
  exchange_rate: number | null
  po_rbs_ch_no: string | null
  po_rbs_th_no: string | null
  po_date: string | null
  filename: string | null
  created_at: string
  rows: POItemRow[] | null
}

// ── Period helpers ────────────────────────────────────────────────────────────
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
function generateMonthKeys(count = 18): string[] {
  const keys: string[] = []
  const d = new Date()
  for (let i = 0; i < count; i++) {
    keys.unshift(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`)
    d.setMonth(d.getMonth() - 1)
  }
  return keys
}

// ── Helpers ──────────────────────────────────────────────────────────────────
const PALETTE = ['#3d8b82','#d4962a','#c85a3a','#6b5ea8','#2a7c9a','#c87a3a','#5a9a6b','#c85a82','#a89a3a','#3a82c8']

function fmt(n: number, dec = 0) {
  return n.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec })
}

// ── DonutChart ───────────────────────────────────────────────────────────────
function DonutChart({ slices, total, size = 140, centerLabel, centerSub }: {
  slices: { label: string; value: number; color: string }[]
  total: number
  size?: number
  centerLabel?: string
  centerSub?: string
}) {
  const cx = size / 2, cy = size / 2, R = size * 0.44, r = size * 0.29
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
  const label = centerLabel ?? (total >= 1000 ? `${(total / 1000).toFixed(1)}k` : fmt(total))
  return (
    <svg viewBox={`0 0 ${size} ${size}`} width={size} height={size}>
      {arcs.map((arc, i) => <path key={i} d={arc.path} fill={arc.color} opacity={0.88} />)}
      <circle cx={cx} cy={cy} r={r - 3} fill="#faf5ee" />
      <text x={cx} y={cy - 6} textAnchor="middle" fontSize={size * 0.065} fill="#8a7a6a" fontWeight="600">{centerSub ?? 'items'}</text>
      <text x={cx} y={cy + 9} textAnchor="middle" fontSize={size * 0.11} fill="#3a2a1a" fontWeight="800">{label}</text>
    </svg>
  )
}

// ── Page ─────────────────────────────────────────────────────────────────────
export default function POSummaryPage() {
  const [items, setItems] = useState<POItem[]>([])
  const [uploads, setUploads] = useState<POUpload[]>([])
  const [loading, setLoading] = useState(true)
  const [cnyRate, setCnyRate] = useState(4.85)
  const [usdRate, setUsdRate] = useState(33.00)
  const [vendorCodeMap, setVendorCodeMap] = useState<Map<string, string>>(new Map())
  const [selectedProject, setSelectedProject] = useState<string>('all')
  const [expandedGroup, setExpandedGroup] = useState<string | null>(null)
  const [exporting, setExporting] = useState(false)

  // Period filter
  const allMonthKeys = useMemo(() => generateMonthKeys(18), [])
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(() => {
    const now = new Date()
    return new Set([`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`])
  })
  const [periodOpen, setPeriodOpen] = useState(false)
  const toggleMonth = useCallback((k: string) => {
    setSelectedMonths(prev => { const n = new Set(prev); n.has(k) ? n.delete(k) : n.add(k); return n })
  }, [])

  useEffect(() => {
    async function load() {
      const [{ data: poItems }, { data: poUploads }, { data: settings }, { data: invs }] = await Promise.all([
        supabase.from('po_items').select('project, supplier, item_code, description, fob_price, currency'),
        supabase.from('po_uploads').select('id, supplier, project, currency, total_amount, exchange_rate, po_rbs_ch_no, po_rbs_th_no, po_date, filename, created_at, rows'),
        supabase.from('cost_settings').select('key, value'),
        supabase.from('invoices').select('supplier, vendor_code').not('vendor_code', 'is', null),
      ])
      setItems((poItems ?? []) as POItem[])
      setUploads((poUploads ?? []) as POUpload[])
      if (settings) {
        const m = Object.fromEntries((settings as { key: string; value: string }[]).map(r => [r.key, r.value]))
        if (m.cny_rate) setCnyRate(parseFloat(m.cny_rate))
        if (m.usd_rate) setUsdRate(parseFloat(m.usd_rate))
      }
      if (invs) {
        const vcMap = new Map<string, string>()
        for (const inv of invs as { supplier: string | null; vendor_code: string | null }[]) {
          if (inv.supplier && inv.vendor_code) vcMap.set(inv.supplier, inv.vendor_code)
        }
        setVendorCodeMap(vcMap)
      }
      setLoading(false)
    }
    load()
  }, [])

  // ── Derived data ──────────────────────────────────────────────────────────

  // Uploads filtered by period
  const periodFilteredUploads = useMemo(() => {
    if (selectedMonths.size === 0) return uploads
    return uploads.filter(u => { const k = mKey(u.po_date); return k !== null && selectedMonths.has(k) })
  }, [uploads, selectedMonths])

  // Keys of uploads that have actual po records (from po_uploads, not Cost Compare)
  const poKeys = useMemo(() => new Set(periodFilteredUploads.map(u => `${u.supplier}|${u.project}`)), [periodFilteredUploads])

  // po_items filtered to only those in period-filtered uploads
  const poOnlyItems = useMemo(() => items.filter(i => poKeys.has(`${i.supplier}|${i.project}`)), [items, poKeys])

  // Latest po_upload per (supplier, project) for linking in item detail
  const poMap = useMemo(() => {
    const map = new Map<string, POUpload>()
    for (const u of periodFilteredUploads) {
      const key = `${u.supplier}|${u.project}`
      const ex = map.get(key)
      if (!ex || u.created_at > ex.created_at) map.set(key, u)
    }
    return map
  }, [periodFilteredUploads])

  // Row-level lookup: {supplier|project|item_code} → {qty, unit_price, total}
  const rowLookup = useMemo(() => {
    const map = new Map<string, { qty: number; unit_price: number; total: number }>()
    for (const u of uploads) {
      if (!u.rows) continue
      for (const row of u.rows) {
        const key = `${u.supplier}|${u.project}|${row.item_code}`
        if (!map.has(key)) map.set(key, { qty: row.qty, unit_price: row.unit_price, total: row.total })
      }
    }
    return map
  }, [uploads])

  const allProjects = useMemo(() =>
    [...new Set(uploads.map(u => u.project))].filter(Boolean).sort()
  , [uploads])

  const filteredItems = useMemo(() =>
    selectedProject === 'all' ? poOnlyItems : poOnlyItems.filter(i => i.project === selectedProject)
  , [poOnlyItems, selectedProject])

  const filteredUploads = useMemo(() =>
    selectedProject === 'all' ? periodFilteredUploads : periodFilteredUploads.filter(u => u.project === selectedProject)
  , [periodFilteredUploads, selectedProject])

  const rateFor = (u: POUpload) => u.exchange_rate ?? (u.currency === 'USD' ? usdRate : cnyRate)

  // Unique item_codes from PO rows — dedup by item_code only (1 item = 1 SKU regardless of supplier)
  const rowItems = useMemo((): POItem[] => {
    const seen = new Map<string, POItem>()
    for (const u of filteredUploads) {
      if (!u.rows) continue
      for (const row of u.rows) {
        if (!row.item_code) continue
        if (!seen.has(row.item_code)) {
          seen.set(row.item_code, {
            project: u.project,
            supplier: u.supplier,
            item_code: row.item_code,
            description: row.description || null,
            fob_price: row.unit_price,
            currency: u.currency,
          })
        }
      }
    }
    return Array.from(seen.values())
  }, [filteredUploads])

  // Per-item detail: item_code → all supplier entries (for detail table)
  const groupDetailMap = useMemo(() => {
    const map = new Map<string, DetailEntry[]>()
    // latest PO per supplier in filteredUploads
    const latestPo = new Map<string, POUpload>()
    for (const u of filteredUploads) {
      const ex = latestPo.get(u.supplier)
      if (!ex || u.created_at > ex.created_at) latestPo.set(u.supplier, u)
    }
    for (const u of filteredUploads) {
      if (!u.rows) continue
      const rate = rateFor(u)
      const po = latestPo.get(u.supplier)!
      for (const row of u.rows) {
        if (!row.item_code) continue
        if (!map.has(row.item_code)) map.set(row.item_code, [])
        const entries = map.get(row.item_code)!
        if (!entries.find(e => e.supplier === u.supplier)) {
          entries.push({
            supplier: u.supplier,
            qty: row.qty,
            unit_price: row.unit_price,
            currency: u.currency,
            unit_thb: row.unit_price * rate,
            total_thb: row.qty * row.unit_price * rate,
            poId: po.id,
            poLabel: po.po_rbs_ch_no || po.filename?.replace(/\.xlsx?$/i, '') || po.id.slice(0, 8),
          })
        }
      }
    }
    return map
  }, [filteredUploads, cnyRate, usdRate])

  const grandPoThb = useMemo(() =>
    filteredUploads.reduce((s, u) => s + u.total_amount * rateFor(u), 0)
  , [filteredUploads, cnyRate, usdRate])

  // Allocate every PO's total_amount×rate to product groups.
  // Priority: (1) po_uploads.rows item totals, (2) po_items matched by supplier+project,
  // (3) po_items matched by supplier only (catches project field mismatch), (4) 'Unknown'.
  // Guarantees sum(groupFobAlloc) === grandPoThb.
  const groupFobAlloc = useMemo(() => {
    const byGroup = new Map<string, number>()
    const bySuppGroup = new Map<string, Map<string, number>>() // group → supplier → thb

    const add = (group: string, supplier: string, thb: number) => {
      byGroup.set(group, (byGroup.get(group) || 0) + thb)
      if (!bySuppGroup.has(group)) bySuppGroup.set(group, new Map())
      const sm = bySuppGroup.get(group)!
      sm.set(supplier, (sm.get(supplier) || 0) + thb)
    }

    for (const u of filteredUploads) {
      const rate = rateFor(u)
      const totalThb = u.total_amount * rate

      // 1. Use row-level totals when available and non-zero
      if (u.rows && u.rows.length > 0) {
        const rowsSum = u.rows.reduce((s, r) => s + (r.total || 0), 0)
        if (rowsSum > 0) {
          for (const row of u.rows) {
            add(extractGroup(row.description, row.item_code), u.supplier, (row.total / rowsSum) * totalThb)
          }
          continue
        }
        // rows exist but all totals are 0 → fall through to po_items
      }

      // 2. po_items: supplier + project exact match (period-filtered)
      let matched = filteredItems.filter(i => i.supplier === u.supplier && i.project === u.project)

      // 3a. Supplier only in period-filtered items (project mismatch within same period)
      if (matched.length === 0) matched = filteredItems.filter(i => i.supplier === u.supplier)

      // 3b. Supplier only across all periods (last resort — cross-period fallback)
      if (matched.length === 0) matched = items.filter(i => i.supplier === u.supplier)

      if (matched.length > 0) {
        const wSum = matched.reduce((s, i) => s + (i.fob_price || 1), 0) || 1
        for (const item of matched) {
          add(extractGroup(item.description, item.item_code), u.supplier, totalThb * ((item.fob_price || 1) / wSum))
        }
      } else {
        // 4. No item data at all → Unknown
        add('Unknown', u.supplier, totalThb)
      }
    }

    return { byGroup, bySuppGroup }
  }, [filteredUploads, filteredItems, items, cnyRate, usdRate])

  // Group items by product category — unique item_codes, multi-supplier detail from groupDetailMap
  const groupData = useMemo(() => {
    const map = new Map<string, Set<string>>() // group → unique item_codes
    for (const item of rowItems) {
      const g = extractGroup(item.description, item.item_code)
      if (!map.has(g)) map.set(g, new Set())
      map.get(g)!.add(item.item_code)
    }
    for (const [group] of groupFobAlloc.byGroup) {
      if (!map.has(group)) map.set(group, new Set())
    }

    return Array.from(map.entries())
      .map(([group, itemCodes], gi) => {
        const itemCount = itemCodes.size
        const pct = rowItems.length > 0 ? (itemCount / rowItems.length) * 100 : 0

        const fobThb = groupFobAlloc.byGroup.get(group) ?? 0
        const suppFobMap = groupFobAlloc.bySuppGroup.get(group) ?? new Map<string, number>()

        // Supplier item-code counts from groupDetailMap
        const suppItemCount = new Map<string, number>()
        for (const ic of itemCodes) {
          for (const e of (groupDetailMap.get(ic) ?? [])) {
            suppItemCount.set(e.supplier, (suppItemCount.get(e.supplier) ?? 0) + 1)
          }
        }
        const suppSet = new Set([...suppItemCount.keys(), ...suppFobMap.keys()])
        const suppliers = Array.from(suppSet)
          .map((supplier, si) => {
            const count = suppItemCount.get(supplier) ?? 0
            const suppFob = suppFobMap.get(supplier) ?? 0
            return {
              supplier, count, suppFob,
              pct: itemCount > 0 ? (count / itemCount) * 100 : 0,
              fobPct: fobThb > 0 ? (suppFob / fobThb) * 100 : 0,
              color: PALETTE[si % PALETTE.length],
            }
          })
          .sort((a, b) => b.suppFob - a.suppFob)

        // Detail items: 1 row per unique item_code, with all supplier entries
        const sortedItems: GroupDetailItem[] = Array.from(itemCodes)
          .map(ic => {
            const baseItem = rowItems.find(i => i.item_code === ic)
            const supplierEntries = (groupDetailMap.get(ic) ?? [])
              .sort((a, b) => b.total_thb - a.total_thb)
            return {
              item_code: ic,
              description: baseItem?.description ?? null,
              supplierEntries,
              total_thb: supplierEntries.reduce((s, e) => s + e.total_thb, 0),
            }
          })
          .sort((a, b) => b.total_thb - a.total_thb)

        return { group, itemCount, pct, fobThb, suppliers, suppSet, sortedItems, color: PALETTE[gi % PALETTE.length] }
      })
      .sort((a, b) => b.itemCount - a.itemCount)
  }, [rowItems, groupFobAlloc, groupDetailMap])

  // Sum all allocated groups (includes Unknown) — equals grandPoThb when allocation is complete
  const grandGroupFobThb = useMemo(
    () => Array.from(groupFobAlloc.byGroup.values()).reduce((s, v) => s + v, 0),
    [groupFobAlloc]
  )

  // Supplier totals from po_uploads
  const supplierTotals = useMemo(() => {
    const map = new Map<string, number>()
    for (const u of filteredUploads) map.set(u.supplier, (map.get(u.supplier) || 0) + u.total_amount * rateFor(u))
    return Array.from(map.entries()).map(([s, v]) => ({ supplier: s, thb: v })).sort((a, b) => b.thb - a.thb)
  }, [filteredUploads, cnyRate, usdRate])

  const grandItemCount = rowItems.length
  const overviewSlices = groupData.map(g => ({ label: g.group, value: g.itemCount, color: g.color }))

  // FOB donut slices — all groups are now in groupData (including rows-only groups with itemCount=0)
  const fobDonutSlices = useMemo(
    () => groupData.map(g => ({ label: g.group, value: g.fobThb, color: g.color })),
    [groupData]
  )

  // ── Export Excel ────────────────────────────────────────────────────────────
  async function exportExcel() {
    setExporting(true)
    try {
      const excelRows: Record<string, string | number>[] = []
      for (const u of filteredUploads) {
        const po = u
        const vendorCode = vendorCodeMap.get(u.supplier) || ''
        if (po.rows && po.rows.length > 0) {
          for (const row of po.rows) {
            excelRows.push({
              'Item Code': row.item_code,
              'Description': row.description || '',
              'Supplier': u.supplier,
              'Vendor Code': vendorCode,
              'Quantity': row.qty,
              'FOB Unit Price': row.unit_price,
              'FOB Total Price': row.total,
              'Original Currency': u.currency,
              'PO RBS CH': u.po_rbs_ch_no || '',
              'Group': extractGroup(row.description, row.item_code),
            })
          }
        } else {
          // No row detail — use po_items for this (supplier, project)
          const matchingItems = (selectedProject === 'all' ? items : items.filter(i => i.project === selectedProject))
            .filter(i => i.supplier === u.supplier && i.project === u.project)
          for (const item of matchingItems) {
            const rowData = rowLookup.get(`${item.supplier}|${item.project}|${item.item_code}`)
            excelRows.push({
              'Item Code': item.item_code,
              'Description': item.description || '',
              'Supplier': item.supplier,
              'Vendor Code': vendorCode,
              'Quantity': rowData?.qty ?? '',
              'FOB Unit Price': item.fob_price,
              'FOB Total Price': rowData?.total ?? '',
              'Original Currency': item.currency,
              'PO RBS CH': u.po_rbs_ch_no || '',
              'Group': extractGroup(item.description, item.item_code),
            })
          }
        }
      }

      const ws = XLSX.utils.json_to_sheet(excelRows)
      ws['!cols'] = [
        { wch: 22 }, { wch: 50 }, { wch: 20 }, { wch: 16 },
        { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 14 },
        { wch: 20 }, { wch: 20 },
      ]
      const wb = XLSX.utils.book_new()
      const sheetName = selectedMonths.size === 1 ? mLabel([...selectedMonths][0]) : 'PO Summary'
      XLSX.utils.book_append_sheet(wb, ws, sheetName)
      XLSX.writeFile(wb, `po-summary-${new Date().toISOString().slice(0, 10)}.xlsx`)
    } finally {
      setExporting(false)
    }
  }

  if (loading) {
    return (
      <div className="min-h-screen flex flex-col" style={{ background: '#faf5ee' }}>
        <NavBar />
        <div className="flex-1 flex items-center justify-center">
          <p className="text-sm" style={{ color: '#d4962a' }}>กำลังโหลด...</p>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#faf5ee' }}>
      <NavBar />

      <div className="max-w-7xl mx-auto w-full px-6 py-8">

        {/* Header */}
        <div className="flex items-start justify-between flex-wrap gap-4 mb-4">
          <div>
            <h1 className="text-2xl font-bold" style={{ color: '#3a2a1a' }}>PO Summary</h1>
            <p className="text-sm mt-0.5" style={{ color: '#8a7a6a' }}>สัดส่วนรายการสินค้าตามกลุ่มและ Supplier จากทุก PO</p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <button onClick={exportExcel} disabled={exporting || filteredUploads.length === 0}
              className="px-4 py-2 text-sm font-medium rounded-lg border transition-colors disabled:opacity-40"
              style={{ background: '#fff', borderColor: '#d4c8b0', color: '#5a4a3a' }}>
              {exporting ? '...' : '↓ Export Excel'}
            </button>
            {/* Project filter */}
            <div className="flex gap-1 flex-wrap">
              <button onClick={() => setSelectedProject('all')}
                className="px-3 py-1 rounded-full text-xs font-medium border transition-colors"
                style={selectedProject === 'all'
                  ? { background: '#d4962a', color: '#fff', borderColor: '#d4962a' }
                  : { background: '#fdf8f0', color: '#8a7a6a', borderColor: '#d4c8b0' }}>
                ทั้งหมด
              </button>
              {allProjects.map(p => (
                <button key={p} onClick={() => setSelectedProject(p)}
                  className="px-3 py-1 rounded-full text-xs font-medium border transition-colors"
                  style={selectedProject === p
                    ? { background: '#3d8b82', color: '#fff', borderColor: '#3d8b82' }
                    : { background: '#fdf8f0', color: '#5a7a78', borderColor: '#b8d8d4' }}>
                  {p}
                </button>
              ))}
            </div>
          </div>
        </div>

        {/* Period filter bar */}
        <div className="bg-gray-800 rounded-xl px-5 py-3 mb-5 flex items-center gap-3 flex-wrap relative">
          <span className="text-xs font-bold text-gray-400 uppercase tracking-widest shrink-0">Period:</span>
          <div className="relative">
            <button onClick={() => setPeriodOpen(o => !o)}
              className="flex items-center gap-2 px-4 py-1.5 rounded-lg text-xs font-bold text-gray-200"
              style={{ background: '#1e3a4a', border: '1px solid #2e5060', minWidth: 170 }}>
              <span className="flex-1 text-left">
                {selectedMonths.size === 0 ? 'ทั้งหมด'
                  : selectedMonths.size === 1 ? mLabel([...selectedMonths][0])
                  : `${selectedMonths.size} เดือนที่เลือก`}
              </span>
              <span className="text-gray-500 text-xs">{periodOpen ? '▲' : '▼'}</span>
            </button>
            {periodOpen && (
              <div className="absolute top-full left-0 mt-1 z-50 rounded-xl p-3 shadow-2xl"
                style={{ background: '#1a2e3c', border: '1px solid #2e5060', minWidth: 280 }}>
                <div className="flex gap-2 mb-3 pb-2" style={{ borderBottom: '1px solid #2a4455' }}>
                  <button onClick={() => setSelectedMonths(new Set())}
                    className="flex-1 py-1 rounded-lg text-xs font-bold" style={{ background: '#2a4455', color: '#8a9aaa' }}>ทั้งหมด</button>
                  <button onClick={() => setPeriodOpen(false)}
                    className="px-3 py-1 rounded-lg text-xs font-bold" style={{ background: '#d4962a', color: '#fff' }}>Done</button>
                </div>
                <div className="grid grid-cols-3 gap-1.5">
                  {allMonthKeys.map(k => (
                    <button key={k} onClick={() => toggleMonth(k)}
                      className="py-1.5 rounded-lg text-xs font-bold transition-all"
                      style={selectedMonths.has(k)
                        ? { background: '#d4962a', color: '#1a2d3a', border: '1px solid #d4962a' }
                        : { background: 'transparent', color: '#8a9aaa', border: '1px solid #2a4455' }}>
                      {mLabel(k)}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>
          {[...selectedMonths].sort().map(k => (
            <span key={k} className="flex items-center gap-1 px-3 py-1 rounded-full text-xs font-bold"
              style={{ background: '#d4962a', color: '#fff' }}>
              {mLabel(k)}
              <button onClick={() => toggleMonth(k)} className="ml-1 opacity-75 hover:opacity-100">×</button>
            </span>
          ))}
          {selectedMonths.size > 0 && (
            <span className="ml-auto text-xs text-gray-400">{filteredUploads.length} PO</span>
          )}
        </div>

        {grandItemCount === 0 ? (
          <div className="bg-white rounded-2xl border border-amber-100 p-12 text-center">
            <p className="text-sm" style={{ color: '#c0a060' }}>ไม่มีข้อมูล PO ในช่วงเวลาที่เลือก</p>
          </div>
        ) : (<>

          {/* Stats row */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
            {[
              { label: 'Item Code', value: fmt(grandItemCount), color: '#3d8b82' },
              { label: 'กลุ่มสินค้า', value: fmt(groupData.length), color: '#d4962a' },
              { label: 'Suppliers', value: fmt(new Set(filteredItems.map(i => i.supplier)).size), color: '#6b5ea8' },
              { label: 'FOB THB รวม', value: grandPoThb > 0 ? grandPoThb.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—', color: '#c85a3a' },
            ].map(s => (
              <div key={s.label} className="bg-white rounded-xl border border-amber-100 shadow-sm px-4 py-3">
                <p className="text-xs" style={{ color: '#8a7a6a' }}>{s.label}</p>
                <p className="text-xl font-bold mt-0.5" style={{ color: s.color }}>{s.value}</p>
              </div>
            ))}
          </div>

          {/* Overview chart */}
          <div className="bg-white rounded-2xl border border-amber-100 shadow-sm p-6 mb-6">
            <h2 className="text-sm font-bold mb-5" style={{ color: '#3a2a1a' }}>สัดส่วนตามกลุ่มสินค้า</h2>
            <div className="flex flex-col lg:flex-row gap-8 items-start">
              {/* Two donuts */}
              <div className="shrink-0 flex gap-6 items-start">
                <div className="flex flex-col items-center gap-1">
                  <DonutChart slices={overviewSlices} total={grandItemCount} size={150}
                    centerLabel={String(grandItemCount)} centerSub="item code" />
                  <span className="text-xs font-semibold mt-1" style={{ color: '#8a7a6a' }}>by Item Code</span>
                </div>
                <div className="flex flex-col items-center gap-1">
                  <DonutChart
                    slices={fobDonutSlices}
                    total={grandGroupFobThb} size={150}
                    centerLabel={grandGroupFobThb >= 1000000 ? `${(grandGroupFobThb/1000000).toFixed(1)}M` : `${(grandGroupFobThb/1000).toFixed(0)}k`}
                    centerSub="FOB THB" />
                  <span className="text-xs font-semibold mt-1" style={{ color: '#8a7a6a' }}>by FOB THB</span>
                </div>
              </div>
              <div className="flex-1 overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-xs font-semibold border-b" style={{ color: '#8a7a6a', borderColor: '#ede8df' }}>
                      <th className="py-2 text-left pr-3">กลุ่มสินค้า</th>
                      <th className="py-2 text-right pr-2">Item Code</th>
                      <th className="py-2 text-right pr-5">% items</th>
                      <th className="py-2 text-right pr-2">FOB THB</th>
                      <th className="py-2 text-right pr-5">% FOB</th>
                      <th className="py-2 text-left">Suppliers</th>
                    </tr>
                  </thead>
                  <tbody>
                    {groupData.map(g => {
                      const fobPct = grandGroupFobThb > 0 ? (g.fobThb / grandGroupFobThb) * 100 : 0
                      return (
                        <tr key={g.group} className="border-b hover:bg-amber-50/50 cursor-pointer transition-colors"
                          style={{ borderColor: '#ede8df' }}
                          onClick={() => {
                            const next = expandedGroup === g.group ? null : g.group
                            setExpandedGroup(next)
                            if (next) setTimeout(() => document.getElementById(`group-card-${next}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 50)
                          }}>
                          <td className="py-2.5 pr-3">
                            <div className="flex items-center gap-2">
                              <span className="w-3 h-3 rounded-full shrink-0" style={{ background: g.color }} />
                              <span className="font-semibold" style={{ color: '#3a2a1a' }}>{g.group}</span>
                            </div>
                          </td>
                          <td className="py-2.5 pr-2 text-right font-medium" style={{ color: '#3a2a1a' }}>{fmt(g.itemCount)}</td>
                          <td className="py-2.5 pr-5 text-right">
                            <div className="flex items-center justify-end gap-1.5">
                              <div className="w-12 h-1.5 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                                <div className="h-full rounded-full" style={{ width: `${Math.min(g.pct, 100)}%`, background: g.color }} />
                              </div>
                              <span className="text-xs font-semibold w-9 text-right" style={{ color: g.color }}>{g.pct.toFixed(1)}%</span>
                            </div>
                          </td>
                          <td className="py-2.5 pr-2 text-right font-medium" style={{ color: '#5a6a68' }}>
                            {g.fobThb >= 1000000 ? `${(g.fobThb/1000000).toFixed(1)}M` : g.fobThb >= 1000 ? `${(g.fobThb/1000).toFixed(0)}k` : fmt(g.fobThb, 0)}
                          </td>
                          <td className="py-2.5 pr-5 text-right">
                            <div className="flex items-center justify-end gap-1.5">
                              <div className="w-12 h-1.5 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                                <div className="h-full rounded-full" style={{ width: `${Math.min(fobPct, 100)}%`, background: g.color, opacity: 0.65 }} />
                              </div>
                              <span className="text-xs font-semibold w-9 text-right" style={{ color: g.color, opacity: 0.8 }}>{fobPct.toFixed(1)}%</span>
                            </div>
                          </td>
                          <td className="py-2.5 text-xs" style={{ color: '#5a6a68' }}>
                            {g.suppliers.map(s => s.supplier).join(', ')}
                          </td>
                        </tr>
                      )
                    })}
                    <tr style={{ borderTop: '2px solid #d4962a' }}>
                      <td className="py-2.5 pr-3 font-bold text-sm" style={{ color: '#3a2a1a' }}>TOTAL</td>
                      <td className="py-2.5 pr-2 text-right font-bold" style={{ color: '#3d8b82' }}>{fmt(grandItemCount)}</td>
                      <td />
                      <td className="py-2.5 pr-2 text-right font-bold" style={{ color: '#3d8b82' }}>
                        {grandGroupFobThb >= 1000000 ? `${(grandGroupFobThb/1000000).toFixed(1)}M` : fmt(grandGroupFobThb, 0)}
                      </td>
                      <td colSpan={2} />
                    </tr>
                    {(groupFobAlloc.byGroup.get('Unknown') ?? 0) > 0 && (
                      <tr>
                        <td colSpan={6} className="py-1.5 text-xs" style={{ color: '#b0a090' }}>
                          * รวม PO ที่ไม่มี item record ({((groupFobAlloc.byGroup.get('Unknown')! / grandGroupFobThb) * 100).toFixed(1)}% → จัดไว้ใน "Unknown") — อัปโหลดใหม่ผ่าน PO Insights เพื่อแบ่งกลุ่มได้ถูกต้อง
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          {/* FOB THB by Supplier */}
          {supplierTotals.length > 0 && grandPoThb > 0 && (
            <div className="bg-white rounded-2xl border border-amber-100 shadow-sm p-6 mb-6">
              <h2 className="text-sm font-bold mb-4" style={{ color: '#3a2a1a' }}>มูลค่า PO รวมแยกตาม Supplier (FOB THB)</h2>
              <div className="flex flex-col lg:flex-row gap-6 items-start">
                <div className="shrink-0">
                  <DonutChart
                    slices={supplierTotals.map((s, i) => ({ label: s.supplier, value: s.thb, color: PALETTE[i % PALETTE.length] }))}
                    total={grandPoThb} size={140}
                    centerLabel={grandPoThb >= 1000000 ? `${(grandPoThb / 1000000).toFixed(1)}M` : `${(grandPoThb/1000).toFixed(0)}k`}
                    centerSub="FOB THB"
                  />
                </div>
                <div className="flex-1 space-y-2.5">
                  {supplierTotals.map((s, i) => {
                    const pct = grandPoThb > 0 ? (s.thb / grandPoThb) * 100 : 0
                    const color = PALETTE[i % PALETTE.length]
                    return (
                      <div key={s.supplier}>
                        <div className="flex items-center justify-between mb-1">
                          <div className="flex items-center gap-2">
                            <span className="w-2.5 h-2.5 rounded-full" style={{ background: color }} />
                            <span className="text-sm font-medium" style={{ color: '#3a2a1a' }}>{s.supplier}</span>
                          </div>
                          <div className="flex items-center gap-3 text-xs">
                            <span style={{ color: '#6a6a6a' }}>{fmt(s.thb)} THB</span>
                            <span className="font-semibold w-12 text-right" style={{ color }}>{pct.toFixed(1)}%</span>
                          </div>
                        </div>
                        <div className="w-full h-1.5 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                          <div className="h-full rounded-full" style={{ width: `${pct}%`, background: color }} />
                        </div>
                      </div>
                    )
                  })}
                  <p className="text-xs pt-1" style={{ color: '#b0a090' }}>
                    ใช้ exchange rate ต่อ PO (PO ที่ยังไม่ได้ใส่ rate ใช้ estimate CNY × {cnyRate}, USD × {usdRate})
                  </p>
                </div>
              </div>
            </div>
          )}

          {/* Per-group cards */}
          <h2 className="text-sm font-bold mb-3" style={{ color: '#3a2a1a' }}>รายละเอียดแยกตามกลุ่มสินค้า</h2>
          <div className="space-y-3">
            {groupData.map(g => {
              const isExpanded = expandedGroup === g.group
              return (
                <div key={g.group} id={`group-card-${g.group}`} className="bg-white rounded-2xl border shadow-sm overflow-hidden"
                  style={{ borderColor: isExpanded ? g.color : '#ede8df' }}>

                  {/* Card header */}
                  <div className="px-5 py-4 flex items-center justify-between cursor-pointer"
                    onClick={() => setExpandedGroup(isExpanded ? null : g.group)}>
                    <div className="flex items-center gap-3 flex-wrap">
                      <span className="w-3 h-3 rounded-full" style={{ background: g.color }} />
                      <span className="font-bold text-base" style={{ color: '#3a2a1a' }}>{g.group}</span>
                      <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: g.color + '22', color: g.color }}>
                        {g.pct.toFixed(1)}% (items)
                      </span>
                      {grandGroupFobThb > 0 && (
                        <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: g.color + '15', color: g.color, opacity: 0.85 }}>
                          {((g.fobThb / grandGroupFobThb) * 100).toFixed(1)}% (FOB)
                        </span>
                      )}
                      <span className="text-xs" style={{ color: '#8a7a6a' }}>{fmt(g.itemCount)} item codes · {g.suppSet.size} supplier{g.suppSet.size > 1 ? 's' : ''}</span>
                    </div>
                    <span className="text-gray-400 text-sm ml-4 shrink-0">{isExpanded ? '▲' : '▼'}</span>
                  </div>

                  {/* Supplier bars */}
                  <div className="px-5 pb-4 border-t" style={{ borderColor: '#f5f0e8' }}>
                    <div className="flex flex-col lg:flex-row gap-4 pt-4 items-start">
                      <div className="shrink-0">
                        <DonutChart
                          slices={g.suppliers.map(s => ({ label: s.supplier, value: s.suppFob, color: s.color }))}
                          total={g.fobThb} size={100}
                          centerLabel={String(g.itemCount)} centerSub="item code"
                        />
                      </div>
                      <div className="flex-1 space-y-2.5 pt-1 min-w-0">
                        {g.suppliers.map(s => (
                          <div key={s.supplier}>
                            <div className="flex items-center gap-1.5 min-w-0 mb-1">
                              <span className="w-2 h-2 rounded-full shrink-0" style={{ background: s.color }} />
                              <span className="text-xs font-medium truncate" style={{ color: '#3a2a1a' }}>{s.supplier}</span>
                            </div>
                            {/* Item code bar */}
                            <div className="flex items-center gap-2 mb-0.5">
                              <span className="text-xs w-16 shrink-0" style={{ color: '#9a8a7a' }}>item code</span>
                              <div className="flex-1 h-1.5 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                                <div className="h-full rounded-full" style={{ width: `${s.pct}%`, background: s.color }} />
                              </div>
                              <span className="text-xs font-semibold w-8 text-right shrink-0" style={{ color: s.color }}>{s.pct.toFixed(0)}%</span>
                              <span className="text-xs w-14 text-right shrink-0" style={{ color: '#8a7a6a' }}>{s.count} codes</span>
                            </div>
                            {/* FOB THB bar */}
                            <div className="flex items-center gap-2">
                              <span className="text-xs w-16 shrink-0" style={{ color: '#9a8a7a' }}>FOB THB</span>
                              <div className="flex-1 h-1.5 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                                <div className="h-full rounded-full" style={{ width: `${s.fobPct}%`, background: s.color, opacity: 0.55 }} />
                              </div>
                              <span className="text-xs font-semibold w-8 text-right shrink-0" style={{ color: s.color, opacity: 0.8 }}>{s.fobPct.toFixed(0)}%</span>
                              <span className="text-xs w-14 text-right shrink-0" style={{ color: '#8a7a6a' }}>
                                {s.suppFob >= 1000000 ? `${(s.suppFob/1000000).toFixed(1)}M` : `${(s.suppFob/1000).toFixed(0)}k`}
                              </span>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Expanded: item detail table */}
                  {isExpanded && (
                    <div className="border-t" style={{ borderColor: '#ede8df' }}>
                      <div className="px-5 py-3">
                        <p className="text-xs font-semibold mb-2" style={{ color: '#8a7a6a' }}>ITEM CODES ในกลุ่มนี้ ({g.itemCount} รายการ)</p>
                        <div className="overflow-x-auto">
                          <table className="w-full text-xs">
                            <thead>
                              <tr style={{ color: '#9a8a7a', borderBottom: '1px solid #ede8df' }}>
                                <th className="text-left py-2 pr-4 font-semibold">Item Code</th>
                                <th className="text-left py-2 pr-4 font-semibold">Description</th>
                                <th className="text-left py-2 pr-3 font-semibold">PO</th>
                                <th className="text-left py-2 pr-4 font-semibold">Supplier</th>
                                <th className="text-right py-2 pr-3 font-semibold">QTY</th>
                                <th className="text-right py-2 pr-3 font-semibold">Unit Price</th>
                                <th className="text-right py-2 pr-3 font-semibold">THB/unit</th>
                                <th className="text-right py-2 font-semibold">FOB THB รวม</th>
                              </tr>
                            </thead>
                            <tbody>
                              {g.sortedItems.map((item) => (
                                item.supplierEntries.map((e, ei) => (
                                  <tr key={`${item.item_code}-${e.supplier}`}
                                    className="border-t"
                                    style={{ borderColor: ei === 0 ? '#ddd8ce' : '#f5f0e8' }}>
                                    {ei === 0 ? (
                                      <>
                                        <td className="py-2 pr-4 font-mono font-semibold text-xs align-top" rowSpan={item.supplierEntries.length} style={{ color: '#3a2a1a' }}>{item.item_code}</td>
                                        <td className="py-2 pr-4 text-xs align-top" rowSpan={item.supplierEntries.length} style={{ color: '#5a5a5a', maxWidth: '220px' }}>{item.description || '—'}</td>
                                      </>
                                    ) : null}
                                    <td className="py-1.5 pr-3 text-xs">
                                      <Link href={`/po-builder/${e.poId}`} className="hover:underline font-mono" style={{ color: '#d4962a' }}>{e.poLabel}</Link>
                                    </td>
                                    <td className="py-1.5 pr-4 text-xs font-semibold" style={{ color: '#3d8b82' }}>{e.supplier}</td>
                                    <td className="py-1.5 pr-3 text-right text-xs tabular-nums" style={{ color: '#3a2a1a' }}>{fmt(e.qty)}</td>
                                    <td className="py-1.5 pr-3 text-right text-xs tabular-nums" style={{ color: '#5a5a5a' }}>{e.currency === 'USD' ? '$' : '¥'}{fmt(e.unit_price, 2)}</td>
                                    <td className="py-1.5 pr-3 text-right text-xs tabular-nums" style={{ color: '#6a5a4a' }}>{fmt(e.unit_thb, 2)}</td>
                                    <td className="py-1.5 text-right text-xs font-semibold tabular-nums" style={{ color: '#3d8b82' }}>{fmt(e.total_thb, 2)}</td>
                                  </tr>
                                ))
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              )
            })}
          </div>

        </>)}

        <p className="text-xs mt-6 text-center" style={{ color: '#c0b0a0' }}>
          ข้อมูลจาก {filteredItems.length} item codes · {filteredUploads.length} PO
        </p>
      </div>
    </div>
  )
}

'use client'

import { useEffect, useState, useMemo } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'
import NavBar from '@/components/NavBar'

// ── Types ────────────────────────────────────────────────────────────────────
interface POItem {
  project: string
  supplier: string
  item_code: string
  description: string | null
  fob_price: number
  currency: string
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
}

// ── Helpers ──────────────────────────────────────────────────────────────────
const PALETTE = ['#3d8b82','#d4962a','#c85a3a','#6b5ea8','#2a7c9a','#c87a3a','#5a9a6b','#c85a82','#a89a3a','#3a82c8']

function fmt(n: number, dec = 0) {
  return n.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec })
}

import { extractGroup } from '@/lib/po-group'

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
  const [selectedProject, setSelectedProject] = useState<string>('all')
  const [expandedGroup, setExpandedGroup] = useState<string | null>(null)

  useEffect(() => {
    async function load() {
      const [{ data: poItems }, { data: poUploads }, { data: settings }] = await Promise.all([
        supabase.from('po_items').select('project, supplier, item_code, description, fob_price, currency'),
        supabase.from('po_uploads').select('id, supplier, project, currency, total_amount, exchange_rate, po_rbs_ch_no, po_rbs_th_no, po_date, filename, created_at'),
        supabase.from('cost_settings').select('key, value'),
      ])
      setItems((poItems ?? []) as POItem[])
      setUploads((poUploads ?? []) as POUpload[])
      if (settings) {
        const m = Object.fromEntries((settings as { key: string; value: string }[]).map(r => [r.key, r.value]))
        if (m.cny_rate) setCnyRate(parseFloat(m.cny_rate))
        if (m.usd_rate) setUsdRate(parseFloat(m.usd_rate))
      }
      setLoading(false)
    }
    load()
  }, [])

  // Latest po_upload per (supplier, project) for linking
  const poMap = useMemo(() => {
    const map = new Map<string, POUpload>()
    for (const u of uploads) {
      const key = `${u.supplier}|${u.project}`
      const ex = map.get(key)
      if (!ex || u.created_at > ex.created_at) map.set(key, u)
    }
    return map
  }, [uploads])

  const allProjects = useMemo(() =>
    [...new Set(items.map(i => i.project))].filter(Boolean).sort()
  , [items])

  const filteredItems = useMemo(() =>
    selectedProject === 'all' ? items : items.filter(i => i.project === selectedProject)
  , [items, selectedProject])

  const filteredUploads = useMemo(() =>
    selectedProject === 'all' ? uploads : uploads.filter(u => u.project === selectedProject)
  , [uploads, selectedProject])

  const rateFor = (u: POUpload) => u.exchange_rate ?? (u.currency === 'USD' ? usdRate : cnyRate)

  const grandPoThb = useMemo(() =>
    filteredUploads.reduce((s, u) => s + u.total_amount * rateFor(u), 0)
  , [filteredUploads, cnyRate, usdRate])

  // Group items by product category
  const groupData = useMemo(() => {
    const map = new Map<string, { items: POItem[] }>()
    for (const item of filteredItems) {
      const g = extractGroup(item.description, item.item_code)
      if (!map.has(g)) map.set(g, { items: [] })
      map.get(g)!.items.push(item)
    }

    return Array.from(map.entries())
      .map(([group, { items: gItems }], gi) => {
        const itemCount = gItems.length
        const pct = filteredItems.length > 0 ? (itemCount / filteredItems.length) * 100 : 0

        // Supplier breakdown by item count
        const suppMap = new Map<string, number>()
        for (const item of gItems) suppMap.set(item.supplier, (suppMap.get(item.supplier) || 0) + 1)

        const suppliers = Array.from(suppMap.entries())
          .map(([supplier, count], si) => ({
            supplier,
            count,
            pct: itemCount > 0 ? (count / itemCount) * 100 : 0,
            color: PALETTE[si % PALETTE.length],
          }))
          .sort((a, b) => b.count - a.count)

        const suppSet = new Set(gItems.map(i => i.supplier))

        // Unique items sorted by supplier then item_code
        const sortedItems = [...gItems].sort((a, b) =>
          a.supplier.localeCompare(b.supplier) || a.item_code.localeCompare(b.item_code))

        return { group, itemCount, pct, suppliers, suppSet, sortedItems, color: PALETTE[gi % PALETTE.length] }
      })
      .sort((a, b) => b.itemCount - a.itemCount)
  }, [filteredItems])

  // Supplier totals from po_uploads
  const supplierTotals = useMemo(() => {
    const map = new Map<string, number>()
    for (const u of filteredUploads) map.set(u.supplier, (map.get(u.supplier) || 0) + u.total_amount * rateFor(u))
    return Array.from(map.entries()).map(([s, v]) => ({ supplier: s, thb: v })).sort((a, b) => b.thb - a.thb)
  }, [filteredUploads, cnyRate, usdRate])

  const grandItemCount = filteredItems.length
  const overviewSlices = groupData.map(g => ({ label: g.group, value: g.itemCount, color: g.color }))

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
        <div className="flex items-start justify-between flex-wrap gap-4 mb-6">
          <div>
            <h1 className="text-2xl font-bold" style={{ color: '#3a2a1a' }}>PO Summary</h1>
            <p className="text-sm mt-0.5" style={{ color: '#8a7a6a' }}>สัดส่วนรายการสินค้าตามกลุ่มและ Supplier จากทุก PO</p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs font-semibold uppercase tracking-wide" style={{ color: '#8a7a6a' }}>Project</span>
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

        {grandItemCount === 0 ? (
          <div className="bg-white rounded-2xl border border-amber-100 p-12 text-center">
            <p className="text-sm" style={{ color: '#c0a060' }}>ยังไม่มีข้อมูลรายการสินค้า กรุณาอัปโหลด PO ผ่านหน้า PO Insights</p>
          </div>
        ) : (<>

          {/* Stats row */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
            {[
              { label: 'Item Codes ทั้งหมด', value: fmt(grandItemCount), color: '#3d8b82' },
              { label: 'กลุ่มสินค้า', value: fmt(groupData.length), color: '#d4962a' },
              { label: 'Suppliers', value: fmt(new Set(filteredItems.map(i => i.supplier)).size), color: '#6b5ea8' },
              { label: 'FOB THB รวม (est.)', value: grandPoThb > 0 ? `${(grandPoThb / 1000000).toFixed(1)}M` : '—', color: '#c85a3a' },
            ].map(s => (
              <div key={s.label} className="bg-white rounded-xl border border-amber-100 shadow-sm px-4 py-3">
                <p className="text-xs" style={{ color: '#8a7a6a' }}>{s.label}</p>
                <p className="text-xl font-bold mt-0.5" style={{ color: s.color }}>{s.value}</p>
              </div>
            ))}
          </div>

          {/* Overview — item group proportion */}
          <div className="bg-white rounded-2xl border border-amber-100 shadow-sm p-6 mb-6">
            <h2 className="text-sm font-bold mb-5" style={{ color: '#3a2a1a' }}>สัดส่วนตามกลุ่มสินค้า (Item Code Count)</h2>
            <div className="flex flex-col lg:flex-row gap-8 items-start">
              <div className="shrink-0">
                <DonutChart slices={overviewSlices} total={grandItemCount} size={160}
                  centerLabel={grandItemCount >= 1000 ? `${(grandItemCount/1000).toFixed(1)}k` : String(grandItemCount)}
                  centerSub="items" />
              </div>
              <div className="flex-1 overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-xs font-semibold border-b" style={{ color: '#8a7a6a', borderColor: '#ede8df' }}>
                      <th className="py-2 text-left pr-3">กลุ่มสินค้า</th>
                      <th className="py-2 text-right pr-3">Item Codes</th>
                      <th className="py-2 text-right pr-6">% ของรวม</th>
                      <th className="py-2 text-left">Suppliers</th>
                    </tr>
                  </thead>
                  <tbody>
                    {groupData.map(g => (
                      <tr key={g.group} className="border-b hover:bg-amber-50/50 cursor-pointer transition-colors"
                        style={{ borderColor: '#ede8df' }}
                        onClick={() => setExpandedGroup(expandedGroup === g.group ? null : g.group)}>
                        <td className="py-2.5 pr-3">
                          <div className="flex items-center gap-2">
                            <span className="w-3 h-3 rounded-full shrink-0" style={{ background: g.color }} />
                            <span className="font-semibold" style={{ color: '#3a2a1a' }}>{g.group}</span>
                          </div>
                        </td>
                        <td className="py-2.5 pr-3 text-right font-medium" style={{ color: '#3a2a1a' }}>{fmt(g.itemCount)}</td>
                        <td className="py-2.5 pr-6 text-right">
                          <div className="flex items-center justify-end gap-2">
                            <div className="w-16 h-1.5 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                              <div className="h-full rounded-full" style={{ width: `${Math.min(g.pct, 100)}%`, background: g.color }} />
                            </div>
                            <span className="text-xs font-semibold w-10 text-right" style={{ color: g.color }}>{g.pct.toFixed(1)}%</span>
                          </div>
                        </td>
                        <td className="py-2.5 text-xs" style={{ color: '#5a6a68' }}>
                          {g.suppliers.map(s => s.supplier).join(', ')}
                        </td>
                      </tr>
                    ))}
                    <tr style={{ borderTop: '2px solid #d4962a' }}>
                      <td className="py-2.5 pr-3 font-bold text-sm" style={{ color: '#3a2a1a' }}>TOTAL</td>
                      <td className="py-2.5 pr-3 text-right font-bold" style={{ color: '#3d8b82' }}>{fmt(grandItemCount)}</td>
                      <td className="py-2.5 pr-6 text-right font-bold text-xs" style={{ color: '#8a7a6a' }}>100%</td>
                      <td />
                    </tr>
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
                    total={grandPoThb}
                    size={140}
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
                <div key={g.group} className="bg-white rounded-2xl border shadow-sm overflow-hidden"
                  style={{ borderColor: isExpanded ? g.color : '#ede8df' }}>

                  {/* Card header */}
                  <div className="px-5 py-4 flex items-center justify-between cursor-pointer"
                    onClick={() => setExpandedGroup(isExpanded ? null : g.group)}>
                    <div className="flex items-center gap-3 flex-wrap">
                      <span className="w-3 h-3 rounded-full" style={{ background: g.color }} />
                      <span className="font-bold text-base" style={{ color: '#3a2a1a' }}>{g.group}</span>
                      <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: g.color + '22', color: g.color }}>
                        {g.pct.toFixed(1)}% of total
                      </span>
                      <span className="text-xs" style={{ color: '#8a7a6a' }}>{fmt(g.itemCount)} item codes · {g.suppSet.size} supplier{g.suppSet.size > 1 ? 's' : ''}</span>
                    </div>
                    <span className="text-gray-400 text-sm ml-4 shrink-0">{isExpanded ? '▲' : '▼'}</span>
                  </div>

                  {/* Supplier proportion bars */}
                  <div className="px-5 pb-4 border-t" style={{ borderColor: '#f5f0e8' }}>
                    <div className="flex flex-col lg:flex-row gap-4 pt-4 items-start">
                      <div className="shrink-0">
                        <DonutChart
                          slices={g.suppliers.map(s => ({ label: s.supplier, value: s.count, color: s.color }))}
                          total={g.itemCount}
                          size={100}
                          centerLabel={String(g.itemCount)}
                          centerSub="items"
                        />
                      </div>
                      <div className="flex-1 space-y-1.5 pt-1 min-w-0">
                        {g.suppliers.map(s => (
                          <div key={s.supplier}>
                            <div className="flex items-center justify-between mb-0.5">
                              <div className="flex items-center gap-1.5 min-w-0">
                                <span className="w-2 h-2 rounded-full shrink-0" style={{ background: s.color }} />
                                <span className="text-xs font-medium truncate" style={{ color: '#3a2a1a' }}>{s.supplier}</span>
                              </div>
                              <div className="flex items-center gap-2 shrink-0 ml-2">
                                <span className="text-xs" style={{ color: '#8a7a6a' }}>{s.count} items</span>
                                <span className="text-xs font-semibold w-10 text-right" style={{ color: s.color }}>{s.pct.toFixed(1)}%</span>
                              </div>
                            </div>
                            <div className="w-full h-1 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                              <div className="h-full rounded-full" style={{ width: `${s.pct}%`, background: s.color }} />
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Expanded item detail */}
                  {isExpanded && (
                    <div className="border-t" style={{ borderColor: '#ede8df' }}>
                      <div className="px-5 py-3">
                        <p className="text-xs font-semibold mb-2" style={{ color: '#8a7a6a' }}>ITEM CODES ในกลุ่มนี้ ({g.itemCount} รายการ)</p>
                        <div className="overflow-x-auto">
                          <table className="w-full text-xs">
                            <thead>
                              <tr style={{ color: '#9a8a7a', borderBottom: '1px solid #ede8df' }}>
                                <th className="text-left py-2 pr-3">Item Code</th>
                                <th className="text-left py-2 pr-3">Description</th>
                                <th className="text-left py-2 pr-3">Supplier</th>
                                <th className="text-left py-2 pr-3">PO</th>
                                <th className="text-right py-2 pr-3">Unit Price</th>
                                <th className="text-right py-2">Currency</th>
                              </tr>
                            </thead>
                            <tbody>
                              {g.sortedItems.map((item, idx) => {
                                const po = poMap.get(`${item.supplier}|${item.project}`)
                                return (
                                  <tr key={idx} className="border-t" style={{ borderColor: '#f5f0e8' }}>
                                    <td className="py-1.5 pr-3 font-mono font-medium" style={{ color: '#3a2a1a' }}>{item.item_code}</td>
                                    <td className="py-1.5 pr-3" style={{ color: '#5a5a5a' }}>{item.description || <span style={{ color: '#c0b0a0' }}>—</span>}</td>
                                    <td className="py-1.5 pr-3 font-medium" style={{ color: '#3d8b82' }}>{item.supplier}</td>
                                    <td className="py-1.5 pr-3">
                                      {po ? (
                                        <Link href={`/po-builder/${po.id}`}
                                          className="hover:underline font-mono"
                                          style={{ color: '#d4962a' }}>
                                          {po.po_rbs_ch_no || po.filename || po.id.slice(0, 8)}
                                        </Link>
                                      ) : <span style={{ color: '#c0b0a0' }}>—</span>}
                                    </td>
                                    <td className="py-1.5 pr-3 text-right" style={{ color: '#5a5a5a' }}>{fmt(item.fob_price, 2)}</td>
                                    <td className="py-1.5 text-right font-medium" style={{ color: '#8a7a6a' }}>{item.currency}</td>
                                  </tr>
                                )
                              })}
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

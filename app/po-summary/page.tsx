'use client'

import { useEffect, useState, useMemo } from 'react'
import { supabase } from '@/lib/supabase'
import NavBar from '@/components/NavBar'

// ── Types ────────────────────────────────────────────────────────────────────
interface POUpload {
  id: string
  supplier: string
  project: string
  currency: string
  po_date: string | null
  total_amount: number
  rows: { item_code: string; description: string; qty: number; unit_price: number; total: number }[]
}

interface FlatItem {
  group: string
  supplier: string
  qty: number
  fob_orig: number
  fob_thb: number
  currency: string
}

// ── Helpers ──────────────────────────────────────────────────────────────────
const PALETTE = ['#3d8b82','#d4962a','#c85a3a','#6b5ea8','#2a7c9a','#c87a3a','#5a9a6b','#c85a82','#a89a3a','#3a82c8']

function fmt(n: number, dec = 2) {
  return n.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec })
}

function extractGroup(description: string, itemCode: string): string {
  const d = (description || itemCode || '').trim()
  if (!d) return 'Other'
  const lower = d.toLowerCase()
  // Multi-word prefixes
  if (lower.startsWith('h-beam') || lower.startsWith('h beam')) return 'H-Beam'
  if (lower.startsWith('data strip') || lower.startsWith('data-strip')) return 'Data Strip'
  // First segment before "-"
  const first = d.split('-')[0].trim()
  if (!first || first.length <= 1) {
    const firstWord = d.split(' ')[0].trim()
    return firstWord || 'Other'
  }
  return first.charAt(0).toUpperCase() + first.slice(1)
}

// ── DonutChart ───────────────────────────────────────────────────────────────
function DonutChart({ slices, total, size = 140 }: {
  slices: { label: string; value: number; color: string }[]
  total: number
  size?: number
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
  const label = total >= 1000000
    ? `${(total / 1000000).toFixed(1)}M`
    : total >= 1000
    ? `${(total / 1000).toFixed(0)}k`
    : fmt(total, 0)
  return (
    <svg viewBox={`0 0 ${size} ${size}`} width={size} height={size}>
      {arcs.map((arc, i) => (
        <path key={i} d={arc.path} fill={arc.color} opacity={0.88} />
      ))}
      <circle cx={cx} cy={cy} r={r - 3} fill="#faf5ee" />
      <text x={cx} y={cy - 7} textAnchor="middle" fontSize={size * 0.065} fill="#8a7a6a" fontWeight="600">FOB THB</text>
      <text x={cx} y={cy + 9} textAnchor="middle" fontSize={size * 0.1} fill="#3a2a1a" fontWeight="800">{label}</text>
    </svg>
  )
}

// ── Page ─────────────────────────────────────────────────────────────────────
export default function POSummaryPage() {
  const [uploads, setUploads] = useState<POUpload[]>([])
  const [loading, setLoading] = useState(true)
  const [cnyRate, setCnyRate] = useState(4.85)
  const [usdRate, setUsdRate] = useState(33.00)
  const [selectedProject, setSelectedProject] = useState<string>('all')
  const [expandedGroup, setExpandedGroup] = useState<string | null>(null)

  useEffect(() => {
    async function load() {
      const [{ data: pos }, { data: settings }] = await Promise.all([
        supabase.from('po_uploads').select('id, supplier, project, currency, po_date, total_amount, rows').order('po_date', { ascending: false }),
        supabase.from('cost_settings').select('key, value'),
      ])
      setUploads((pos ?? []) as POUpload[])
      if (settings) {
        const m = Object.fromEntries((settings as { key: string; value: string }[]).map(r => [r.key, r.value]))
        if (m.cny_rate) setCnyRate(parseFloat(m.cny_rate))
        if (m.usd_rate) setUsdRate(parseFloat(m.usd_rate))
      }
      setLoading(false)
    }
    load()
  }, [])

  const allProjects = useMemo(() =>
    [...new Set(uploads.map(u => u.project))].sort()
  , [uploads])

  const filtered = useMemo(() =>
    selectedProject === 'all' ? uploads : uploads.filter(u => u.project === selectedProject)
  , [uploads, selectedProject])

  // Flatten all items across filtered POs
  const flatItems = useMemo<FlatItem[]>(() => {
    const items: FlatItem[] = []
    for (const po of filtered) {
      if (!po.rows?.length) continue
      const rate = po.currency === 'USD' ? usdRate : cnyRate
      for (const row of po.rows) {
        if (!row.item_code && !row.description) continue
        items.push({
          group: extractGroup(row.description, row.item_code),
          supplier: po.supplier,
          qty: row.qty || 0,
          fob_orig: row.total || 0,
          fob_thb: (row.total || 0) * rate,
          currency: po.currency,
        })
      }
    }
    return items
  }, [filtered, cnyRate, usdRate])

  // Group by product category
  const groupData = useMemo(() => {
    const map = new Map<string, FlatItem[]>()
    for (const item of flatItems) {
      if (!map.has(item.group)) map.set(item.group, [])
      map.get(item.group)!.push(item)
    }
    const grandThb = flatItems.reduce((s, i) => s + i.fob_thb, 0)

    return Array.from(map.entries())
      .map(([group, items], gi) => {
        const totalThb = items.reduce((s, i) => s + i.fob_thb, 0)
        const totalQty = items.reduce((s, i) => s + i.qty, 0)
        const pct = grandThb > 0 ? (totalThb / grandThb) * 100 : 0

        // Supplier breakdown
        const suppMap = new Map<string, { qty: number; thb: number }>()
        for (const item of items) {
          const s = suppMap.get(item.supplier) || { qty: 0, thb: 0 }
          s.qty += item.qty
          s.thb += item.fob_thb
          suppMap.set(item.supplier, s)
        }
        const suppliers = Array.from(suppMap.entries())
          .map(([supplier, v], si) => ({ supplier, ...v, pct: totalThb > 0 ? (v.thb / totalThb) * 100 : 0, color: PALETTE[si % PALETTE.length] }))
          .sort((a, b) => b.thb - a.thb)

        return { group, items, totalThb, totalQty, pct, suppliers, color: PALETTE[gi % PALETTE.length] }
      })
      .sort((a, b) => b.totalThb - a.totalThb)
  }, [flatItems])

  const grandThb = flatItems.reduce((s, i) => s + i.fob_thb, 0)
  const grandQty = flatItems.reduce((s, i) => s + i.qty, 0)

  // Overview donut slices
  const overviewSlices = groupData.map(g => ({ label: g.group, value: g.totalThb, color: g.color }))

  if (loading) {
    return (
      <div className="min-h-screen bg-amber-50 flex flex-col">
        <NavBar />
        <div className="flex-1 flex items-center justify-center">
          <p className="text-sm text-amber-600">กำลังโหลด...</p>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen flex flex-col" style={{ background: '#faf5ee' }}>
      <NavBar />

      <div className="max-w-7xl mx-auto w-full px-6 py-8">

        {/* Header */}
        <div className="flex items-center justify-between flex-wrap gap-3 mb-6">
          <div>
            <h1 className="text-2xl font-bold" style={{ color: '#3a2a1a' }}>PO Summary</h1>
            <p className="text-sm mt-0.5" style={{ color: '#8a7a6a' }}>สรุปสัดส่วนสินค้าตามกลุ่มและ Supplier จากทุก PO</p>
          </div>
          {/* Project filter */}
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs font-semibold uppercase tracking-wide" style={{ color: '#8a7a6a' }}>Project</span>
            <div className="flex gap-1 flex-wrap">
              <button onClick={() => setSelectedProject('all')}
                className={`px-3 py-1 rounded-full text-xs font-medium border transition-colors ${selectedProject === 'all' ? 'border-amber-500 text-amber-900' : 'border-amber-200 text-amber-700 hover:border-amber-400'}`}
                style={selectedProject === 'all' ? { background: '#d4962a', color: '#fff' } : { background: '#fdf8f0' }}>
                ทั้งหมด
              </button>
              {allProjects.map(p => (
                <button key={p} onClick={() => setSelectedProject(p)}
                  className={`px-3 py-1 rounded-full text-xs font-medium border transition-colors`}
                  style={selectedProject === p
                    ? { background: '#3d8b82', color: '#fff', borderColor: '#3d8b82' }
                    : { background: '#fdf8f0', color: '#5a7a78', borderColor: '#b8d8d4' }}>
                  {p}
                </button>
              ))}
            </div>
          </div>
        </div>

        {flatItems.length === 0 ? (
          <div className="bg-white rounded-2xl border border-amber-100 p-12 text-center">
            <p className="text-amber-600 text-sm">ยังไม่มีข้อมูล PO หรือยังไม่มีรายการสินค้าในระบบ</p>
          </div>
        ) : (<>

          {/* ── Overview ──────────────────────────────────────────────── */}
          <div className="bg-white rounded-2xl border border-amber-100 shadow-sm p-6 mb-6">
            <h2 className="text-sm font-bold mb-5" style={{ color: '#3a2a1a' }}>สัดส่วนตามกลุ่มสินค้า (Estimated FOB THB)</h2>
            <div className="flex flex-col lg:flex-row gap-8 items-start">
              {/* Donut */}
              <div className="shrink-0">
                <DonutChart slices={overviewSlices} total={grandThb} size={160} />
                <p className="text-center text-xs mt-1" style={{ color: '#8a7a6a' }}>Grand Total</p>
              </div>

              {/* Overview table */}
              <div className="flex-1 overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-xs font-semibold border-b" style={{ color: '#8a7a6a', borderColor: '#ede8df' }}>
                      <th className="py-2 text-left pr-3">กลุ่มสินค้า</th>
                      <th className="py-2 text-right pr-3">QTY รวม</th>
                      <th className="py-2 text-right pr-3">Estimated FOB THB</th>
                      <th className="py-2 text-right pr-3">% ของรวม</th>
                      <th className="py-2 text-left">Supplier หลัก</th>
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
                        <td className="py-2.5 pr-3 text-right" style={{ color: '#5a5a5a' }}>{fmt(g.totalQty, 0)}</td>
                        <td className="py-2.5 pr-3 text-right font-medium" style={{ color: '#3a2a1a' }}>{fmt(g.totalThb)}</td>
                        <td className="py-2.5 pr-3 text-right">
                          <div className="flex items-center justify-end gap-2">
                            <div className="w-16 h-1.5 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                              <div className="h-full rounded-full" style={{ width: `${Math.min(g.pct, 100)}%`, background: g.color }} />
                            </div>
                            <span className="text-xs font-semibold" style={{ color: g.color }}>{g.pct.toFixed(1)}%</span>
                          </div>
                        </td>
                        <td className="py-2.5 text-xs" style={{ color: '#6a6a6a' }}>
                          {g.suppliers[0]?.supplier}{g.suppliers.length > 1 ? ` +${g.suppliers.length - 1}` : ''}
                        </td>
                      </tr>
                    ))}
                    <tr style={{ borderTop: '2px solid #d4962a' }}>
                      <td className="py-2.5 pr-3 font-bold text-sm" style={{ color: '#3a2a1a' }}>TOTAL</td>
                      <td className="py-2.5 pr-3 text-right font-bold" style={{ color: '#3a2a1a' }}>{fmt(grandQty, 0)}</td>
                      <td className="py-2.5 pr-3 text-right font-bold" style={{ color: '#3d8b82' }}>{fmt(grandThb)}</td>
                      <td className="py-2.5 pr-3 text-right font-bold text-xs" style={{ color: '#8a7a6a' }}>100%</td>
                      <td></td>
                    </tr>
                  </tbody>
                </table>
                <p className="text-xs mt-3" style={{ color: '#b0a090' }}>
                  Estimated FOB THB คำนวณจาก ESTIMATE RATES (CNY × {cnyRate}, USD × {usdRate})
                </p>
              </div>
            </div>
          </div>

          {/* ── Per-group cards ───────────────────────────────────────── */}
          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
            {groupData.map(g => {
              const isExpanded = expandedGroup === g.group
              return (
                <div key={g.group} className="bg-white rounded-2xl border shadow-sm overflow-hidden"
                  style={{ borderColor: isExpanded ? g.color : '#ede8df' }}>
                  {/* Card header */}
                  <div className="px-5 pt-4 pb-3 flex items-center justify-between cursor-pointer"
                    onClick={() => setExpandedGroup(isExpanded ? null : g.group)}>
                    <div className="flex items-center gap-2">
                      <span className="w-3 h-3 rounded-full" style={{ background: g.color }} />
                      <span className="font-bold" style={{ color: '#3a2a1a' }}>{g.group}</span>
                      <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: g.color + '22', color: g.color }}>
                        {g.pct.toFixed(1)}%
                      </span>
                    </div>
                    <span className="text-gray-400 text-xs">{isExpanded ? '▲' : '▼'}</span>
                  </div>

                  {/* Mini stats */}
                  <div className="px-5 pb-4 grid grid-cols-2 gap-2 text-xs border-b" style={{ borderColor: '#ede8df' }}>
                    <div>
                      <p style={{ color: '#b0a090' }}>Estimated FOB THB</p>
                      <p className="font-bold" style={{ color: '#3a2a1a' }}>{fmt(g.totalThb)}</p>
                    </div>
                    <div>
                      <p style={{ color: '#b0a090' }}>QTY รวม</p>
                      <p className="font-bold" style={{ color: '#3a2a1a' }}>{fmt(g.totalQty, 0)} ชิ้น</p>
                    </div>
                  </div>

                  {/* Supplier breakdown */}
                  <div className="px-5 py-4">
                    <div className="flex gap-4 items-start">
                      <DonutChart
                        slices={g.suppliers.map(s => ({ label: s.supplier, value: s.thb, color: s.color }))}
                        total={g.totalThb}
                        size={100}
                      />
                      <div className="flex-1 space-y-1.5 pt-1">
                        {g.suppliers.map(s => (
                          <div key={s.supplier}>
                            <div className="flex items-center justify-between mb-0.5">
                              <div className="flex items-center gap-1.5">
                                <span className="w-2 h-2 rounded-full shrink-0" style={{ background: s.color }} />
                                <span className="text-xs font-medium" style={{ color: '#3a2a1a' }}>{s.supplier}</span>
                              </div>
                              <span className="text-xs font-semibold" style={{ color: s.color }}>{s.pct.toFixed(1)}%</span>
                            </div>
                            <div className="w-full h-1 rounded-full overflow-hidden" style={{ background: '#ede8df' }}>
                              <div className="h-full rounded-full" style={{ width: `${s.pct}%`, background: s.color }} />
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Expanded: supplier detail table */}
                  {isExpanded && (
                    <div className="px-5 pb-4 border-t" style={{ borderColor: '#ede8df' }}>
                      <table className="w-full text-xs mt-3">
                        <thead>
                          <tr style={{ color: '#8a7a6a' }}>
                            <th className="text-left py-1">Supplier</th>
                            <th className="text-right py-1">QTY</th>
                            <th className="text-right py-1">Estimated THB</th>
                            <th className="text-right py-1">%</th>
                          </tr>
                        </thead>
                        <tbody>
                          {g.suppliers.map(s => (
                            <tr key={s.supplier} className="border-t" style={{ borderColor: '#f0ebe3' }}>
                              <td className="py-1.5 font-medium" style={{ color: '#3a2a1a' }}>{s.supplier}</td>
                              <td className="py-1.5 text-right" style={{ color: '#5a5a5a' }}>{fmt(s.qty, 0)}</td>
                              <td className="py-1.5 text-right font-medium" style={{ color: '#3a2a1a' }}>{fmt(s.thb)}</td>
                              <td className="py-1.5 text-right font-semibold" style={{ color: s.color }}>{s.pct.toFixed(1)}%</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              )
            })}
          </div>

        </>)}

        <p className="text-xs mt-6 text-center" style={{ color: '#c0b0a0' }}>
          ข้อมูลจาก {filtered.length} PO · {flatItems.length} รายการสินค้า
        </p>
      </div>
    </div>
  )
}

'use client'

import { useEffect, useState, useMemo, useCallback } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'
import { exportQCSummaryPptx, type QCSlideData } from '@/lib/qc-pptx'

// ─── Types ────────────────────────────────────────────────────────────────────
interface QCRow {
  id: string
  report_no: string
  supplier_company: string | null
  issue_found_date: string | null
  status: string
  issue_types: string[] | null
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function mKey(d: string | null): string | null {
  if (!d) return null
  return d.slice(0, 7)
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

const PALETTE = ['#3d8b82','#d4962a','#c85a3a','#6b5ea8','#2a7c9a','#c87a3a','#5a9a6b','#c85a82']
const ISSUE_PALETTE: Record<string, string> = {
  'Packing Quality': '#c87a3a',
  'Working Process': '#6b5ea8',
  'Product Quality': '#c85a3a',
  'Loading Process': '#3d8b82',
}

// ─── DonutChart ───────────────────────────────────────────────────────────────
function DonutChart({ slices, total, label }: { slices: { label: string; value: number; color: string }[]; total: number; label: string }) {
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
      <text x={cx} y={cy + 9} textAnchor="middle" fontSize={18} fill="#3a2a1a" fontWeight="800">{total}</text>
      <text x={cx} y={cy + 23} textAnchor="middle" fontSize={9} fill="#aaa">{label}</text>
    </svg>
  )
}

// ─── BarChart (monthly) ───────────────────────────────────────────────────────
function MonthlyBarChart({ months, values }: { months: string[]; values: number[] }) {
  const max = Math.max(...values, 1)
  const W = 520, H = 130, padL = 24, padR = 8, padT = 18, padB = 22
  const cW = W - padL - padR
  const barW = Math.max(8, cW / months.length - 4)
  return (
    <svg viewBox={`0 0 ${W} ${H}`} className="w-full" style={{ maxHeight: 140 }}>
      {[0, 0.5, 1].map(frac => {
        const y = padT + (H - padT - padB) * (1 - frac)
        return (
          <g key={frac}>
            <line x1={padL} y1={y} x2={W - padR} y2={y} stroke="#e8dcc8" strokeWidth={1} strokeDasharray={frac === 0 ? '0' : '3,3'} />
            {frac > 0 && <text x={padL - 3} y={y + 4} textAnchor="end" fontSize={8} fill="#bbb">{Math.round(max * frac)}</text>}
          </g>
        )
      })}
      {months.map((m, i) => {
        const bH = Math.max(2, (values[i] / max) * (H - padT - padB))
        const x = padL + (i / months.length) * cW + (cW / months.length - barW) / 2
        const y = H - padB - bH
        return (
          <g key={m}>
            <rect x={x} y={y} width={barW} height={bH} fill="#d4962a" rx={2} opacity={0.85} />
            {values[i] > 0 && (
              <text x={x + barW / 2} y={y - 3} textAnchor="middle" fontSize={8} fill="#d4962a" fontWeight="700">{values[i]}</text>
            )}
            <text x={x + barW / 2} y={H - 4} textAnchor="middle" fontSize={7} fill="#aaa">{mLabel(m).slice(0, 3)}</text>
          </g>
        )
      })}
    </svg>
  )
}

const NAV = [
  { href: '/', label: 'PO Matching' }, { href: '/dashboard', label: 'Dashboard' },
  { href: '/calendar', label: 'Calendar' }, { href: '/report', label: 'Report' },
  { href: '/compare', label: 'Cost Compare' },
  { href: '/po-builder', label: 'PO Builder' }, { href: '/guide', label: 'Guide' },
]

// ─── Page ─────────────────────────────────────────────────────────────────────
export default function QCSummaryPage() {
  const [reports, setReports] = useState<QCRow[]>([])
  const [loading, setLoading] = useState(true)
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(new Set())
  const [periodOpen, setPeriodOpen] = useState(false)
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set())
  const [exporting, setExporting] = useState(false)

  const allMonthKeys = useMemo(() => generateMonthKeys(18), [])

  useEffect(() => {
    supabase.from('qc_reports')
      .select('id, report_no, supplier_company, issue_found_date, status, issue_types')
      .order('issue_found_date', { ascending: false })
      .then(({ data }) => {
        if (data) setReports(data as QCRow[])
        setLoading(false)
      })
  }, [])

  async function handleExportPptx() {
    if (selectedIds.size === 0) return
    setExporting(true)
    const { data } = await supabase.from('qc_reports')
      .select('report_no, supplier_company, issue_found_date, issue_types, description, corrective_actions, corrective_action_comment, root_cause, preventive_action, photo_urls')
      .in('id', [...selectedIds])
      .order('issue_found_date', { ascending: true })
    if (data) {
      const slides: QCSlideData[] = data.map(r => ({
        report_no: r.report_no || '',
        supplier_company: r.supplier_company,
        issue_found_date: r.issue_found_date,
        issue_types: r.issue_types ?? [],
        description: r.description || '',
        corrective_actions: r.corrective_actions ?? [],
        corrective_action_comment: r.corrective_action_comment || '',
        root_cause: r.root_cause || '',
        preventive_action: r.preventive_action || '',
        photo_urls: r.photo_urls ?? [],
      }))
      await exportQCSummaryPptx(slides, 'QC-Report')
    }
    setExporting(false)
  }

  function toggleSelect(id: string) {
    setSelectedIds(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n })
  }

  function selectAllFiltered() {
    setSelectedIds(new Set(filtered.map(r => r.id)))
  }

  function clearSelection() {
    setSelectedIds(new Set())
  }

  const filtered = useMemo(() =>
    selectedMonths.size === 0
      ? reports
      : reports.filter(r => { const k = mKey(r.issue_found_date); return k ? selectedMonths.has(k) : false }),
    [reports, selectedMonths])

  const chart12Months = useMemo(() => generateMonthKeys(12), [])

  const monthlyValues = useMemo(() =>
    chart12Months.map(k => reports.filter(r => mKey(r.issue_found_date) === k).length),
    [reports, chart12Months])

  const bySupplier = useMemo(() => {
    const map = new Map<string, number>()
    filtered.forEach(r => {
      const key = r.supplier_company || '(ไม่ระบุ)'
      map.set(key, (map.get(key) ?? 0) + 1)
    })
    return [...map.entries()].sort(([, a], [, b]) => b - a)
  }, [filtered])

  const byIssueType = useMemo(() => {
    const map = new Map<string, number>()
    filtered.forEach(r => {
      (r.issue_types ?? []).forEach(t => map.set(t, (map.get(t) ?? 0) + 1))
    })
    return [...map.entries()].sort(([, a], [, b]) => b - a)
  }, [filtered])

  const openCount = filtered.filter(r => r.status === 'open').length
  const closedCount = filtered.filter(r => r.status === 'closed').length

  const toggleMonth = useCallback((k: string) => {
    setSelectedMonths(prev => { const n = new Set(prev); n.has(k) ? n.delete(k) : n.add(k); return n })
  }, [])

  const ISSUE_COLORS: Record<string, string> = {
    'Packing Quality': '#f97316',
    'Working Process': '#8b5cf6',
    'Product Quality': '#ef4444',
    'Loading Process': '#3b82f6',
  }

  return (
    <div style={{ background: '#ede5d4', minHeight: '100vh', fontFamily: 'system-ui,-apple-system,sans-serif' }}>

      {/* Nav */}
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-6 text-sm sticky top-0 z-20 shadow-sm flex-wrap">
        <span className="font-bold text-gray-900">Import PO</span>
        {NAV.map(n => (
          <Link key={n.href} href={n.href} className="text-gray-500 hover:text-gray-800 transition-colors">{n.label}</Link>
        ))}
        <div className="relative group">
          <span className="text-blue-600 cursor-default">Summary ▾</span>
          <div className="absolute left-0 top-full pt-1 hidden group-hover:block z-50">
            <div className="bg-white border border-gray-200 rounded-lg shadow-lg py-1 min-w-[150px]">
              <Link href="/summary" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">Item Summary</Link>
              <Link href="/qc/summary" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">QC Summary</Link>
            </div>
          </div>
        </div>
        <Link href="/qc" className="text-gray-500 hover:text-gray-800 transition-colors">QC Report</Link>
      </nav>

      {/* Hero header */}
      <div style={{ background: '#1e3340', padding: '24px 32px' }}>
        <div style={{ maxWidth: 1200, margin: '0 auto', display: 'flex', alignItems: 'center', gap: 24, flexWrap: 'wrap' }}>
          <div style={{ flex: 1, minWidth: 200 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: '#3d8b82', textTransform: 'uppercase', letterSpacing: '0.12em', marginBottom: 4 }}>QC Report</div>
            <div style={{ fontSize: 30, fontWeight: 900, color: '#d4962a', lineHeight: 1.1 }}>QC SUMMARY</div>
            <div style={{ fontSize: 13, color: '#7a9aaa', marginTop: 4 }}>สรุปภาพรวม Quality Claim Report</div>
          </div>
          {[
            { icon: '📋', value: String(filtered.length), label: 'Total Reports', bg: '#d4962a' },
            { icon: '🔴', value: String(openCount), label: 'Open', bg: '#c85a3a' },
            { icon: '✅', value: String(closedCount), label: 'Closed', bg: '#3d8b82' },
            { icon: '🏭', value: String(bySupplier.length), label: 'Suppliers', bg: '#6b5ea8' },
          ].map(card => (
            <div key={card.label} style={{ background: card.bg, borderRadius: 14, padding: '14px 20px', minWidth: 120, display: 'flex', alignItems: 'center', gap: 12 }}>
              <span style={{ fontSize: 24 }}>{card.icon}</span>
              <div>
                <div style={{ fontSize: 24, fontWeight: 900, color: '#fff', lineHeight: 1.1 }}>{card.value}</div>
                <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.75)', fontWeight: 600, marginTop: 2 }}>{card.label}</div>
              </div>
            </div>
          ))}
          <button
            onClick={handleExportPptx}
            disabled={selectedIds.size === 0 || exporting}
            style={{
              background: selectedIds.size > 0 ? 'rgba(212,150,42,0.25)' : 'rgba(255,255,255,0.08)',
              border: `1.5px solid ${selectedIds.size > 0 ? '#d4962a' : '#3d5060'}`,
              color: selectedIds.size > 0 ? '#d4962a' : '#4a6a7a',
              borderRadius: 10, padding: '10px 18px', fontSize: 12, fontWeight: 700, cursor: selectedIds.size > 0 ? 'pointer' : 'default',
              transition: 'all 0.2s',
            }}
          >
            {exporting ? '⏳ กำลัง Export...' : selectedIds.size > 0 ? `↓ Export PPT (${selectedIds.size})` : '↓ Export PPT'}
          </button>
        </div>
      </div>

      {/* Period bar */}
      <div style={{ background: '#18303c', padding: '10px 32px', position: 'relative', zIndex: 30 }}>
        <div style={{ maxWidth: 1200, margin: '0 auto', display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 10, fontWeight: 800, color: '#5a8a9a', textTransform: 'uppercase', letterSpacing: '0.12em' }}>Period:</span>
          <div style={{ position: 'relative' }}>
            <button onClick={() => setPeriodOpen(o => !o)} style={{
              display: 'flex', alignItems: 'center', gap: 8,
              padding: '6px 14px', borderRadius: 8, cursor: 'pointer',
              background: '#1e3a4a', border: '1px solid #2e5060',
              color: '#d4c8a8', fontSize: 11, fontWeight: 700, minWidth: 200,
            }}>
              <span style={{ flex: 1, textAlign: 'left' }}>
                {selectedMonths.size === 0 ? 'ทั้งหมด'
                  : selectedMonths.size === 1 ? mLabel([...selectedMonths][0])
                  : `${selectedMonths.size} เดือน`}
              </span>
              <span style={{ fontSize: 9, color: '#5a8a9a' }}>{periodOpen ? '▲' : '▼'}</span>
            </button>
            {periodOpen && (
              <div style={{
                position: 'absolute', top: 'calc(100% + 6px)', left: 0,
                background: '#1a2e3c', border: '1px solid #2e5060', borderRadius: 12,
                boxShadow: '0 8px 32px rgba(0,0,0,0.4)', padding: 12, minWidth: 280, zIndex: 100,
              }}>
                <div style={{ display: 'flex', gap: 8, marginBottom: 10, paddingBottom: 8, borderBottom: '1px solid #2a4455' }}>
                  <button onClick={() => setSelectedMonths(new Set())}
                    style={{ flex: 1, padding: '4px 0', borderRadius: 6, border: 'none', background: '#3d8b82', color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    ทั้งหมด
                  </button>
                  <button onClick={() => setSelectedMonths(new Set(allMonthKeys))}
                    style={{ flex: 1, padding: '4px 0', borderRadius: 6, border: 'none', background: '#2a4455', color: '#8a9aaa', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    18 เดือน
                  </button>
                  <button onClick={() => setPeriodOpen(false)}
                    style={{ padding: '4px 10px', borderRadius: 6, border: 'none', background: '#d4962a', color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    Done
                  </button>
                </div>
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
          {selectedMonths.size > 0 && [...selectedMonths].sort().map(k => (
            <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 4, background: '#d4962a', borderRadius: 6, padding: '3px 8px' }}>
              <span style={{ fontSize: 10, fontWeight: 700, color: '#1a2d3a' }}>{mLabel(k)}</span>
              <button onClick={() => toggleMonth(k)} style={{ background: 'none', border: 'none', color: '#1a2d3a', fontSize: 11, cursor: 'pointer', padding: 0, lineHeight: 1, opacity: 0.7 }}>×</button>
            </div>
          ))}
          <span style={{ color: '#4a7a8a', fontSize: 10, marginLeft: 'auto' }}>
            {filtered.length} reports
          </span>
        </div>
      </div>

      {loading ? (
        <div style={{ textAlign: 'center', padding: '80px 0', color: '#aaa', fontSize: 14 }}>กำลังโหลด...</div>
      ) : (
        <div style={{ maxWidth: 1200, margin: '0 auto', padding: '20px 24px', display: 'grid', gridTemplateColumns: '260px 1fr', gap: 16 }}>

          {/* Left column */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>

            {/* Donut: by supplier */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, padding: '16px 16px 12px' }}>
              <div style={{ fontSize: 11, fontWeight: 800, color: '#5a4a3a', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10, textAlign: 'center' }}>
                Reports by Supplier
              </div>
              <div style={{ display: 'flex', justifyContent: 'center' }}>
                <DonutChart
                  slices={bySupplier.slice(0, 8).map(([name, count], i) => ({ label: name, value: count, color: PALETTE[i % PALETTE.length] }))}
                  total={filtered.length}
                  label="reports"
                />
              </div>
              <div style={{ marginTop: 8, display: 'flex', flexDirection: 'column', gap: 5 }}>
                {bySupplier.slice(0, 6).map(([name, count], i) => (
                  <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
                    <div style={{ width: 9, height: 9, borderRadius: '50%', background: PALETTE[i % PALETTE.length], flexShrink: 0 }} />
                    <span style={{ fontSize: 11, color: '#6a5a4a', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{name}</span>
                    <span style={{ fontSize: 10, fontWeight: 800, color: '#4a3a2a' }}>
                      {filtered.length > 0 ? ((count / filtered.length) * 100).toFixed(0) : 0}%
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
              <div style={{ maxHeight: 300, overflowY: 'auto' }}>
                {bySupplier.length === 0 ? (
                  <div style={{ padding: 20, textAlign: 'center', color: '#bbb', fontSize: 12 }}>ไม่มีข้อมูล</div>
                ) : bySupplier.map(([name, count], i) => (
                  <div key={name} style={{ padding: '9px 14px', borderBottom: '1px solid #f0ebe0', display: 'flex', alignItems: 'center', gap: 9 }}>
                    <div style={{ width: 10, height: 10, borderRadius: '50%', background: PALETTE[i % PALETTE.length], flexShrink: 0 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: '#3a2a1a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{name}</div>
                    </div>
                    <div style={{ fontSize: 16, fontWeight: 900, color: '#d4962a', flexShrink: 0 }}>{count}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Right column */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>

            {/* Monthly bar chart */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, padding: '16px 20px' }}>
              <div style={{ fontSize: 12, fontWeight: 800, color: '#5a4a3a', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 10 }}>
                จำนวน Report รายเดือน (12 เดือนล่าสุด)
              </div>
              <MonthlyBarChart months={chart12Months} values={monthlyValues} />
            </div>

            {/* Issue type breakdown */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, overflow: 'hidden' }}>
              <div style={{ padding: '10px 16px', borderBottom: '1px solid #e2d8c8', fontSize: 11, fontWeight: 800, color: '#5a4a3a', textTransform: 'uppercase', letterSpacing: '0.08em' }}>
                ประเภทปัญหา
              </div>
              {byIssueType.length === 0 ? (
                <div style={{ padding: 20, textAlign: 'center', color: '#bbb', fontSize: 12 }}>ยังไม่มีข้อมูลประเภทปัญหา</div>
              ) : (
                <div style={{ padding: '12px 16px', display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {byIssueType.map(([type, count]) => {
                    const max = byIssueType[0][1]
                    const pct = max > 0 ? (count / max) * 100 : 0
                    const color = ISSUE_COLORS[type] ?? '#6b5ea8'
                    return (
                      <div key={type}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, marginBottom: 4 }}>
                          <span style={{ fontWeight: 700, color: '#3a2a1a' }}>{type}</span>
                          <span style={{ fontWeight: 800, color }}>{count} reports</span>
                        </div>
                        <div style={{ height: 8, background: '#ede8df', borderRadius: 4, overflow: 'hidden' }}>
                          <div style={{ height: '100%', width: `${pct}%`, background: color, borderRadius: 4, transition: 'width 0.4s' }} />
                        </div>
                      </div>
                    )
                  })}
                </div>
              )}
            </div>

            {/* Recent reports table */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, overflow: 'hidden' }}>
              <div style={{ padding: '10px 16px', borderBottom: '1px solid #e2d8c8', display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ fontSize: 11, fontWeight: 800, color: '#5a4a3a', textTransform: 'uppercase', letterSpacing: '0.08em', flex: 1 }}>
                  รายการ ({filtered.length})
                </span>
                {selectedIds.size > 0 && (
                  <span style={{ fontSize: 11, color: '#d4962a', fontWeight: 700 }}>เลือก {selectedIds.size} รายการ</span>
                )}
                <button onClick={selectAllFiltered}
                  style={{ fontSize: 10, fontWeight: 700, color: '#3d8b82', background: 'none', border: '1px solid #3d8b82', borderRadius: 6, padding: '3px 10px', cursor: 'pointer' }}>
                  เลือกทั้งหมด
                </button>
                {selectedIds.size > 0 && (
                  <button onClick={clearSelection}
                    style={{ fontSize: 10, fontWeight: 700, color: '#8a7a6a', background: 'none', border: '1px solid #ccc', borderRadius: 6, padding: '3px 10px', cursor: 'pointer' }}>
                    ยกเลิก
                  </button>
                )}
              </div>
              <div style={{ overflowX: 'auto', maxHeight: 400, overflowY: 'auto' }}>
                <table style={{ width: '100%', fontSize: 11, borderCollapse: 'collapse', minWidth: 600 }}>
                  <thead>
                    <tr style={{ background: '#f0ebe0', position: 'sticky', top: 0, zIndex: 1 }}>
                      <th style={{ padding: '8px 10px', width: 36, borderBottom: '1px solid #e2d8c8' }} />
                      {['Report No.', 'Supplier', 'Issue Date', 'ประเภทปัญหา', 'สถานะ'].map(h => (
                        <th key={h} style={{ padding: '8px 12px', textAlign: 'left', fontSize: 10, fontWeight: 800, color: '#8a7a6a', whiteSpace: 'nowrap', borderBottom: '1px solid #e2d8c8' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((r, i) => {
                      const isSelected = selectedIds.has(r.id)
                      return (
                        <tr key={r.id}
                          style={{ borderBottom: '1px solid #f5efe8', background: isSelected ? '#fff8ec' : i % 2 === 0 ? '#fff' : '#faf7f2' }}
                          onMouseEnter={e => { if (!isSelected) e.currentTarget.style.background = '#f0ece4' }}
                          onMouseLeave={e => { if (!isSelected) e.currentTarget.style.background = i % 2 === 0 ? '#fff' : '#faf7f2' }}>
                          <td style={{ padding: '8px 10px', textAlign: 'center' }}>
                            <input type="checkbox" checked={isSelected}
                              onChange={() => toggleSelect(r.id)}
                              style={{ width: 14, height: 14, cursor: 'pointer', accentColor: '#d4962a' }} />
                          </td>
                          <td style={{ padding: '8px 12px', fontFamily: 'monospace', fontWeight: 700, color: '#3d8b82', whiteSpace: 'nowrap', cursor: 'pointer' }}
                            onClick={() => window.location.href = `/qc/${r.id}`}>{r.report_no}</td>
                          <td style={{ padding: '8px 12px', color: '#5a4a3a' }}>{r.supplier_company || '—'}</td>
                          <td style={{ padding: '8px 12px', color: '#8a7a6a', whiteSpace: 'nowrap' }}>
                            {r.issue_found_date ? new Date(r.issue_found_date).toLocaleDateString('th-TH') : '—'}
                          </td>
                          <td style={{ padding: '8px 12px' }}>
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                              {(r.issue_types ?? []).map(t => (
                                <span key={t} style={{
                                  padding: '2px 8px', borderRadius: 99, fontSize: 10, fontWeight: 700,
                                  background: ISSUE_COLORS[t] ? `${ISSUE_COLORS[t]}22` : '#f0ece4',
                                  color: ISSUE_COLORS[t] ?? '#5a4a3a',
                                }}>{t}</span>
                              ))}
                              {(r.issue_types ?? []).length === 0 && <span style={{ color: '#bbb', fontSize: 10 }}>—</span>}
                            </div>
                          </td>
                          <td style={{ padding: '8px 12px' }}>
                            <span style={{
                              display: 'inline-flex', alignItems: 'center', padding: '2px 8px', borderRadius: 99, fontSize: 10, fontWeight: 700,
                              background: r.status === 'closed' ? '#dcfce7' : '#ffedd5',
                              color: r.status === 'closed' ? '#15803d' : '#c2410c',
                            }}>
                              {r.status === 'closed' ? '✓ ปิดจบ' : '● เปิดอยู่'}
                            </span>
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        </div>
      )}
    </div>
  )
}

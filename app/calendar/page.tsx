'use client'

import { useEffect, useState, useMemo } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'
import LockButton from '@/components/LockButton'

interface CalInvoice {
  id: string
  invoice_no: string
  status: string | null
  estimated_arrival: string | null
  estimated_arrival_end: string | null
  eta_date: string | null
  supplier: string | null
}

const MONTHS_TH = [
  'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
  'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม',
]
const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
const DAY_TH = ['จ.','อ.','พ.','พฤ.','ศ.','ส.','อา.']

const STATUS_CFG: Record<string, { bg: string; text: string; border: string; barBg: string }> = {
  'อยู่ที่จีน':    { bg: 'bg-amber-100',  text: 'text-amber-900',  border: 'border-amber-300',  barBg: '#fef3c7' },
  'On board':      { bg: 'bg-blue-100',   text: 'text-blue-900',   border: 'border-blue-300',   barBg: '#dbeafe' },
  'กำลังเข้าคลัง': { bg: 'bg-orange-100', text: 'text-orange-900', border: 'border-orange-300', barBg: '#fed7aa' },
  'เข้าคลังแล้ว': { bg: 'bg-green-100',  text: 'text-green-900',  border: 'border-green-300',  barBg: '#bbf7d0' },
}
const STATUS_BORDER_HEX: Record<string, string> = {
  'อยู่ที่จีน': '#fcd34d', 'On board': '#93c5fd', 'กำลังเข้าคลัง': '#fb923c', 'เข้าคลังแล้ว': '#4ade80',
}

function computeStatus(
  status: string | null, arrival: string | null,
  arrivalEnd: string | null, etaDate?: string | null,
): string {
  const base = status || 'อยู่ที่จีน'
  const norm = (base === 'ถึงคลัง' || base === 'ถึงไทย กำลังเข้าคลัง') ? 'กำลังเข้าคลัง' : base
  const today = new Date(); today.setHours(0, 0, 0, 0)
  if (norm === 'On board' && etaDate) {
    const eta = new Date(etaDate + 'T00:00:00'); eta.setHours(0, 0, 0, 0)
    if (today >= eta) return 'กำลังเข้าคลัง'
  }
  if (!arrival || norm === 'อยู่ที่จีน') return norm
  const cd = new Date((arrivalEnd || arrival) + 'T00:00:00'); cd.setHours(0, 0, 0, 0)
  if (today > cd) return 'เข้าคลังแล้ว'
  return norm
}

function getWeekNum(date: Date): number {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()))
  const day = d.getUTCDay() || 7
  d.setUTCDate(d.getUTCDate() + 4 - day)
  const y = new Date(Date.UTC(d.getUTCFullYear(), 0, 1))
  return Math.ceil(((d.getTime() - y.getTime()) / 86400000 + 1) / 7)
}

// Monday-first day index: Mon=0 … Sun=6
function dow(d: Date) { return (d.getDay() + 6) % 7 }

function ds(d: Date) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

function pd(s: string) { return new Date(s + 'T00:00:00') }

// Grid column number (1-indexed): col 1 = week label, col 2 = Mon … col 8 = Sun
function dayCol(dayIndex: number) { return dayIndex + 2 }

export default function CalendarPage() {
  const [invoices, setInvoices] = useState<CalInvoice[]>([])
  const [loading, setLoading] = useState(true)
  const [curDate, setCurDate] = useState(() => {
    const d = new Date()
    return new Date(d.getFullYear(), d.getMonth(), 1)
  })

  const today = useMemo(() => { const d = new Date(); d.setHours(0, 0, 0, 0); return d }, [])
  const todayStr = ds(today)

  useEffect(() => { load() }, [])

  async function load() {
    const { data } = await supabase
      .from('invoices')
      .select('id, invoice_no, status, estimated_arrival, estimated_arrival_end, eta_date, supplier')
      .order('estimated_arrival', { ascending: true })
    if (data) setInvoices(data as CalInvoice[])
    setLoading(false)
  }

  const year = curDate.getFullYear()
  const month = curDate.getMonth()

  // All weeks that touch the current month (Mon–Sun, Monday-first)
  const weeks = useMemo(() => {
    const firstDay = new Date(year, month, 1)
    const lastDay = new Date(year, month + 1, 0)
    const firstMon = new Date(firstDay); firstMon.setDate(firstDay.getDate() - dow(firstDay))
    const lastSun = new Date(lastDay); lastSun.setDate(lastDay.getDate() + (6 - dow(lastDay)))

    const result: { mon: Date; sun: Date; wn: number; days: Date[] }[] = []
    let mon = new Date(firstMon)
    while (mon <= lastSun) {
      const sun = new Date(mon); sun.setDate(mon.getDate() + 6)
      const days = Array.from({ length: 7 }, (_, i) => {
        const d = new Date(mon); d.setDate(mon.getDate() + i); return d
      })
      result.push({ mon: new Date(mon), sun: new Date(sun), wn: getWeekNum(mon), days })
      mon.setDate(mon.getDate() + 7)
    }
    return result
  }, [year, month])

  const monthStr = `${year}-${String(month + 1).padStart(2, '0')}`
  const monthArrivals = invoices.filter(inv => inv.estimated_arrival?.startsWith(monthStr))
  const monthEtas = invoices.filter(inv => inv.eta_date?.startsWith(monthStr))

  // Per-week: which invoices overlap, with computed status
  const weeksData = useMemo(() => weeks.map(w => {
    const wMon = ds(w.mon); const wSun = ds(w.sun)
    const inv = invoices
      .filter(inv => {
        if (!inv.estimated_arrival) return false
        const end = inv.estimated_arrival_end || inv.estimated_arrival
        return inv.estimated_arrival <= wSun && end >= wMon
      })
      .map(inv => ({
        ...inv,
        displayStatus: computeStatus(inv.status, inv.estimated_arrival, inv.estimated_arrival_end, inv.eta_date),
      }))
    return { ...w, inv }
  }), [weeks, invoices])

  // Grid template: week-label column + 7 day columns
  const GRID = '88px repeat(7, 1fr)'

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-gray-50 to-slate-100">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4 sticky top-0 z-20 shadow-sm">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/report" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/compare" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Builder</Link>
        <Link href="/calendar" className="text-sm text-blue-600 font-semibold">ปฏิทิน</Link>
        <Link href="/guide" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">คู่มือ</Link>
        <div className="ml-auto"><LockButton /></div>
      </nav>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 py-8">
        {/* Header + month nav */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
          <div>
            <h1 className="text-2xl font-bold text-gray-900 tracking-tight">ปฏิทินเข้าคลัง</h1>
            <p className="text-sm text-gray-400 mt-0.5">ไทม์ไลน์วันที่ประมาณการเข้าคลังของแต่ละ Invoice</p>
          </div>
          <div className="flex items-center gap-1">
            <button
              onClick={() => setCurDate(new Date(year, month - 1, 1))}
              className="w-9 h-9 flex items-center justify-center rounded-xl text-gray-500 hover:bg-gray-100 transition-colors text-xl font-light"
            >‹</button>
            <div className="min-w-[172px] text-center select-none">
              <span className="text-xl font-bold text-gray-800">{MONTHS_TH[month]}</span>
              <span className="text-lg font-semibold text-gray-400 ml-2">{year + 543}</span>
            </div>
            <button
              onClick={() => setCurDate(new Date(year, month + 1, 1))}
              className="w-9 h-9 flex items-center justify-center rounded-xl text-gray-500 hover:bg-gray-100 transition-colors text-xl font-light"
            >›</button>
            <button
              onClick={() => setCurDate(new Date(today.getFullYear(), today.getMonth(), 1))}
              className="ml-2 px-3 py-1.5 text-xs rounded-lg bg-white border border-gray-200 shadow-sm hover:bg-gray-50 text-gray-600 font-medium transition-colors"
            >วันนี้</button>
          </div>
        </div>

        {/* Stats strip */}
        <div className="grid grid-cols-3 gap-3 mb-5">
          <div className="rounded-2xl border border-orange-100 bg-gradient-to-br from-orange-50 to-amber-50 px-4 py-3 shadow-sm">
            <div className="text-2xl font-bold text-orange-600">{loading ? '—' : monthArrivals.length}</div>
            <div className="text-xs text-orange-500 font-medium mt-0.5">เข้าคลังเดือนนี้</div>
          </div>
          <div className="rounded-2xl border border-blue-100 bg-gradient-to-br from-blue-50 to-sky-50 px-4 py-3 shadow-sm">
            <div className="text-2xl font-bold text-blue-600">{loading ? '—' : monthEtas.length}</div>
            <div className="text-xs text-blue-500 font-medium mt-0.5">ETA ถึงท่าเรือเดือนนี้</div>
          </div>
          <div className="rounded-2xl border border-gray-100 bg-white px-4 py-3 shadow-sm">
            <div className="text-2xl font-bold text-gray-700">{loading ? '—' : invoices.length}</div>
            <div className="text-xs text-gray-400 font-medium mt-0.5">Invoice ทั้งหมด</div>
          </div>
        </div>

        {/* Gantt timeline */}
        {loading ? (
          <div className="bg-white rounded-2xl border border-gray-100 shadow-sm p-16 text-center text-gray-300">
            <div className="text-3xl mb-2">⏳</div>
            กำลังโหลด...
          </div>
        ) : (
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
            {weeksData.map(({ mon, sun, wn, days, inv: weekInvs }, wi) => {
              const todayInWeek = ds(today) >= ds(mon) && ds(today) <= ds(sun)

              return (
                <div key={ds(mon)} className={wi > 0 ? 'border-t-2 border-gray-100' : ''}>

                  {/* ── Day header row ── */}
                  <div className="grid border-b border-gray-200" style={{ gridTemplateColumns: GRID }}>
                    {/* Week label */}
                    <div className="bg-gradient-to-b from-amber-50 to-yellow-50 border-r-2 border-amber-200 px-2.5 py-2.5 flex flex-col justify-center">
                      <div className="text-[11px] font-black text-amber-700 tracking-widest uppercase">
                        Week {wn}
                      </div>
                      <div className="text-[10px] text-amber-500 font-semibold mt-0.5">
                        {MONTHS_SHORT[mon.getMonth()]}-{mon.getFullYear()}
                      </div>
                    </div>

                    {/* Day columns */}
                    {days.map((d, di) => {
                      const dStr = ds(d)
                      const isToday = dStr === todayStr
                      const inMonth = d.getMonth() === month
                      const isSun = di === 6
                      return (
                        <div
                          key={di}
                          className={`border-r last:border-r-0 border-gray-100 py-2 text-center ${
                            isToday ? 'bg-blue-50' : isSun ? 'bg-red-50/30' : ''
                          }`}
                        >
                          <div className={`text-[10px] font-bold uppercase tracking-widest ${
                            isSun ? 'text-red-400' : 'text-gray-400'
                          }`}>{DAY_TH[di]}</div>
                          <div className={`mt-1 mx-auto w-7 h-7 flex items-center justify-center rounded-full text-[13px] font-bold ${
                            isToday ? 'bg-blue-600 text-white shadow-md shadow-blue-200' :
                            !inMonth ? 'text-gray-200' :
                            isSun ? 'text-red-400' :
                            'text-gray-700'
                          }`}>{d.getDate()}</div>
                        </div>
                      )
                    })}
                  </div>

                  {/* ── Invoice bar rows ── */}
                  {weekInvs.length === 0 ? (
                    // Empty week placeholder
                    <div
                      className="grid"
                      style={{ gridTemplateColumns: GRID, minHeight: '36px' }}
                    >
                      <div className="bg-amber-50/30 border-r-2 border-amber-100" />
                      {days.map((d, di) => (
                        <div
                          key={di}
                          className={`border-r last:border-r-0 border-gray-50 ${
                            ds(d) === todayStr ? 'bg-blue-50/20' : di === 6 ? 'bg-red-50/10' : ''
                          }`}
                        />
                      ))}
                    </div>
                  ) : weekInvs.map((inv, ii) => {
                    const arrStart = pd(inv.estimated_arrival!)
                    const arrEnd = pd(inv.estimated_arrival_end || inv.estimated_arrival!)
                    const cfg = STATUS_CFG[inv.displayStatus] || STATUS_CFG['อยู่ที่จีน']
                    const borderHex = STATUS_BORDER_HEX[inv.displayStatus] || STATUS_BORDER_HEX['อยู่ที่จีน']

                    // Clamp bar to week boundaries
                    const barStartDay = arrStart < mon ? 0 : dow(arrStart)
                    const barEndDay = arrEnd > sun ? 6 : dow(arrEnd)
                    const continuesPrev = arrStart < mon
                    const continuesNext = arrEnd > sun

                    // Bar grid columns (1-indexed: 1=label, 2=Mon…8=Sun)
                    const colStart = dayCol(barStartDay)
                    const colEnd = dayCol(barEndDay) + 1 // exclusive

                    // Border radius: flat edge if bar continues, pill if it starts/ends
                    const rTL = continuesPrev ? '3px' : '20px'
                    const rBL = continuesPrev ? '3px' : '20px'
                    const rTR = continuesNext ? '3px' : '20px'
                    const rBR = continuesNext ? '3px' : '20px'

                    // ETA marker
                    const etaStr = inv.eta_date
                    const etaInWeek = !!etaStr && etaStr >= ds(mon) && etaStr <= ds(sun)
                    const etaDayIdx = etaInWeek ? dow(pd(etaStr!)) : -1
                    const etaInBar = etaDayIdx >= barStartDay && etaDayIdx <= barEndDay
                    const etaGridCol = etaDayIdx >= 0 ? dayCol(etaDayIdx) : -1

                    const isEven = ii % 2 === 0

                    return (
                      <div
                        key={inv.id + ds(mon)}
                        className={`grid ${isEven ? 'bg-white' : 'bg-gray-50/50'}`}
                        style={{ gridTemplateColumns: GRID, minHeight: '44px' }}
                      >
                        {/* Week label placeholder */}
                        <div
                          className="border-r-2 border-amber-100"
                          style={{ background: 'rgba(254,243,199,0.25)', gridRow: 1 }}
                        />

                        {/* Today column tint (full-height) */}
                        {todayInWeek && days.map((d, di) => {
                          if (ds(d) !== todayStr) return null
                          return (
                            <div
                              key={'today-tint-' + di}
                              style={{ gridColumn: dayCol(di), gridRow: 1 }}
                              className="bg-blue-50/30 pointer-events-none"
                            />
                          )
                        })}

                        {/* ── The bar ── */}
                        <Link
                          href={`/dashboard/${inv.id}`}
                          className="flex items-center overflow-hidden cursor-pointer hover:brightness-95 transition-all"
                          style={{
                            gridColumn: `${colStart} / ${colEnd}`,
                            gridRow: 1,
                            margin: '6px 3px',
                            minHeight: '32px',
                            borderRadius: `${rTL} ${rTR} ${rBR} ${rBL}`,
                            background: cfg.barBg,
                            border: `1.5px solid ${borderHex}`,
                            zIndex: 2,
                          }}
                          title={`${inv.invoice_no}${inv.supplier ? ' · ' + inv.supplier : ''} · ${inv.displayStatus}`}
                        >
                          {/* Left-continues arrow */}
                          {continuesPrev && (
                            <span className="text-[11px] opacity-40 pl-1.5 shrink-0">◀</span>
                          )}

                          {/* Invoice name + supplier */}
                          <span className="text-[11px] font-bold truncate px-2 flex-1">
                            {inv.invoice_no}
                            {inv.supplier && (
                              <span className="font-normal opacity-50 ml-1 text-[10px]">{inv.supplier}</span>
                            )}
                          </span>

                          {/* Right-continues arrow */}
                          {continuesNext && (
                            <span className="text-[11px] opacity-40 pr-1.5 shrink-0">▶</span>
                          )}
                        </Link>

                        {/* ── Truck icon at ETA column (overlaps the bar) ── */}
                        {etaInWeek && (
                          <div
                            className="flex items-center justify-start pointer-events-none"
                            style={{
                              gridColumn: `${etaGridCol} / ${etaGridCol + 1}`,
                              gridRow: 1,
                              zIndex: 3,
                              margin: '6px 0',
                              paddingLeft: '4px',
                            }}
                          >
                            <span
                              className="text-sm drop-shadow-sm"
                              title={`ETA: ${etaStr}`}
                            >
                              🚛
                            </span>
                          </div>
                        )}

                        {/* ETA day column subtle highlight (when ETA is outside bar) */}
                        {etaInWeek && !etaInBar && (
                          <div
                            className="bg-blue-50/50 pointer-events-none"
                            style={{ gridColumn: `${etaGridCol} / ${etaGridCol + 1}`, gridRow: 1 }}
                          />
                        )}
                      </div>
                    )
                  })}
                </div>
              )
            })}
          </div>
        )}

        {/* Legend */}
        <div className="mt-4 flex flex-wrap items-center gap-x-5 gap-y-2">
          <span className="text-[11px] text-gray-400 font-bold uppercase tracking-widest">สัญลักษณ์</span>
          {Object.entries(STATUS_CFG).map(([status, cfg]) => (
            <div
              key={status}
              className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full border text-[11px] font-semibold ${cfg.bg} ${cfg.text} ${cfg.border}`}
            >
              {status}
            </div>
          ))}
          <div className="flex items-center gap-1.5">
            <span className="text-sm">🚛</span>
            <span className="text-xs text-gray-500">ETA ถึงท่าเรือไทย</span>
          </div>
          <div className="flex items-center gap-1.5 text-xs text-gray-400">
            <span>◀▶</span>
            <span>ต่อเนื่องจากสัปดาห์อื่น</span>
          </div>
        </div>
      </div>
    </div>
  )
}

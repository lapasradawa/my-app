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

// Status config – solid colors for crisp bars
const S: Record<string, { fill: string; stroke: string; textColor: string; tag: string }> = {
  'อยู่ที่จีน':    { fill: '#fef9c3', stroke: '#eab308', textColor: '#713f12', tag: 'bg-yellow-100 text-yellow-800 border-yellow-300' },
  'On board':      { fill: '#dbeafe', stroke: '#3b82f6', textColor: '#1e3a8a', tag: 'bg-blue-100 text-blue-800 border-blue-300' },
  'กำลังเข้าคลัง': { fill: '#ffedd5', stroke: '#f97316', textColor: '#7c2d12', tag: 'bg-orange-100 text-orange-800 border-orange-300' },
  'เข้าคลังแล้ว': { fill: '#d1fae5', stroke: '#10b981', textColor: '#065f46', tag: 'bg-emerald-100 text-emerald-800 border-emerald-300' },
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

function dow(d: Date) { return (d.getDay() + 6) % 7 } // Mon=0 … Sun=6

function ds(d: Date) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

function pd(s: string) { return new Date(s + 'T00:00:00') }

// Grid column (1-indexed): col 1 = week label, col 2 = Mon … col 8 = Sun
function gc(dayIdx: number) { return dayIdx + 2 }

const LABEL_W = 80 // px width of week-label column

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

  // All weeks touching current month (Monday-first)
  const weeks = useMemo(() => {
    const firstDay = new Date(year, month, 1)
    const lastDay = new Date(year, month + 1, 0)
    const firstMon = new Date(firstDay); firstMon.setDate(firstDay.getDate() - dow(firstDay))
    const lastSun = new Date(lastDay); lastSun.setDate(lastDay.getDate() + (6 - dow(lastDay)))

    const result: { mon: Date; sun: Date; wn: number; days: Date[] }[] = []
    let m = new Date(firstMon)
    while (m <= lastSun) {
      const s = new Date(m); s.setDate(m.getDate() + 6)
      const days = Array.from({ length: 7 }, (_, i) => {
        const d = new Date(m); d.setDate(m.getDate() + i); return d
      })
      result.push({ mon: new Date(m), sun: new Date(s), wn: getWeekNum(m), days })
      m.setDate(m.getDate() + 7)
    }
    return result
  }, [year, month])

  const monthStr = `${year}-${String(month + 1).padStart(2, '0')}`
  const monthArrivals = invoices.filter(inv => inv.estimated_arrival?.startsWith(monthStr))
  const monthEtas = invoices.filter(inv => inv.eta_date?.startsWith(monthStr))

  const weeksData = useMemo(() => weeks.map(w => {
    const wM = ds(w.mon); const wS = ds(w.sun)
    const inv = invoices
      .filter(inv => {
        if (!inv.estimated_arrival) return false
        return inv.estimated_arrival <= wS && (inv.estimated_arrival_end || inv.estimated_arrival) >= wM
      })
      .map(inv => ({
        ...inv,
        st: computeStatus(inv.status, inv.estimated_arrival, inv.estimated_arrival_end, inv.eta_date),
      }))
    return { ...w, inv }
  }), [weeks, invoices])

  const GRID = `${LABEL_W}px repeat(7, minmax(0, 1fr))`

  return (
    <div className="h-screen overflow-hidden flex flex-col" style={{ background: '#f1f5f9' }}>
      {/* Nav */}
      <nav className="shrink-0 bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4 z-20 shadow-sm">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/calendar" className="text-sm text-blue-600 font-semibold">Calendar</Link>
        <Link href="/report" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/compare" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Builder</Link>
        <Link href="/guide" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Guide</Link>
        <div className="ml-auto"><LockButton /></div>
      </nav>

      {/* Main content — fills remaining height, no page scroll */}
      <div className="flex-1 overflow-hidden flex flex-col px-5 pt-3 pb-2 gap-2 max-w-screen-xl w-full mx-auto">

        {/* ── Compact top bar: title + stats + month nav ── */}
        <div className="shrink-0 flex items-center gap-4">
          {/* Title */}
          <div className="shrink-0">
            <h1 className="text-base font-black text-gray-900 leading-none">ปฏิทินเข้าคลัง</h1>
            <p className="text-[10px] text-gray-400 mt-0.5">ไทม์ไลน์ประมาณการเข้าคลัง</p>
          </div>

          {/* Stats pills */}
          <div className="flex items-center gap-2">
            {[
              { n: loading ? '—' : monthArrivals.length, label: 'เข้าคลังเดือนนี้', color: 'bg-orange-100 text-orange-700 border-orange-200' },
              { n: loading ? '—' : monthEtas.length, label: 'ETA เดือนนี้', color: 'bg-blue-100 text-blue-700 border-blue-200' },
              { n: loading ? '—' : invoices.length, label: 'Invoice ทั้งหมด', color: 'bg-gray-100 text-gray-600 border-gray-200' },
            ].map(s => (
              <div key={s.label} className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full border text-xs font-semibold ${s.color}`}>
                <span className="font-black">{s.n}</span>
                <span className="font-medium opacity-80">{s.label}</span>
              </div>
            ))}
          </div>

          {/* Month nav — pushed to right */}
          <div className="ml-auto flex items-center gap-1.5">
            <button
              onClick={() => setCurDate(new Date(year, month - 1, 1))}
              className="w-7 h-7 flex items-center justify-center rounded-lg bg-white border border-gray-200 text-gray-500 hover:bg-gray-50 shadow-sm transition text-base"
            >‹</button>
            <div className="bg-white border border-gray-200 shadow-sm rounded-lg px-4 py-1.5 min-w-[160px] text-center select-none">
              <span className="text-sm font-bold text-gray-800">{MONTHS_TH[month]}</span>
              <span className="text-xs font-semibold text-gray-400 ml-1.5">{year + 543}</span>
            </div>
            <button
              onClick={() => setCurDate(new Date(year, month + 1, 1))}
              className="w-7 h-7 flex items-center justify-center rounded-lg bg-white border border-gray-200 text-gray-500 hover:bg-gray-50 shadow-sm transition text-base"
            >›</button>
            <button
              onClick={() => setCurDate(new Date(today.getFullYear(), today.getMonth(), 1))}
              className="px-3 py-1.5 text-xs font-semibold rounded-lg bg-blue-600 text-white hover:bg-blue-700 shadow-sm transition"
            >วันนี้</button>
          </div>
        </div>

        {/* ── Timeline ── flex-1 so it fills remaining height */}
        {loading ? (
          <div className="flex-1 bg-white rounded-xl border border-gray-100 shadow-sm flex items-center justify-center text-gray-300">
            <div className="text-center">
              <div className="text-3xl mb-2 animate-pulse">⏳</div>
              <div className="text-sm">กำลังโหลด...</div>
            </div>
          </div>
        ) : (
          <div className="flex-1 overflow-y-auto bg-white rounded-xl shadow-sm border border-gray-100">
            {weeksData.map(({ mon, sun, wn, days, inv: weekInvs }, wi) => {
              const todayInWeek = todayStr >= ds(mon) && todayStr <= ds(sun)
              const todayColIdx = todayInWeek ? dow(today) : -1

              return (
                <div
                  key={ds(mon)}
                  className={wi > 0 ? 'border-t border-slate-200' : ''}
                >
                  {/* ── Day-header row (compact, sticky) ── */}
                  <div className="grid sticky top-0 z-10" style={{ gridTemplateColumns: GRID, background: '#fff' }}>
                    {/* Week label */}
                    <div
                      className="flex items-center gap-1.5 px-2.5 border-r border-slate-700"
                      style={{ background: '#1e293b' }}
                    >
                      <div className="text-lg font-black text-white leading-none">{wn}</div>
                      <div>
                        <div className="text-[8px] font-black tracking-widest text-slate-400 uppercase leading-none">Week</div>
                        <div className="text-[8px] font-semibold text-slate-500 leading-none mt-0.5">
                          {MONTHS_SHORT[mon.getMonth()]}
                        </div>
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
                          className="border-r last:border-r-0 border-gray-100 py-1.5 text-center"
                          style={isSun ? { backgroundColor: 'rgba(0,0,0,0.025)' }
                            : isToday ? { backgroundColor: 'rgba(219,234,254,0.6)' } : undefined}
                        >
                          <div className={`text-[8px] font-black uppercase tracking-wider leading-none mb-0.5 ${
                            isToday ? 'text-blue-500' : isSun ? 'text-gray-300' : 'text-slate-400'
                          }`}>
                            {DAY_TH[di]}
                          </div>
                          <div className={`w-6 h-6 mx-auto flex items-center justify-center rounded-full text-xs font-bold ${
                            isToday
                              ? 'bg-blue-600 text-white shadow-md shadow-blue-300'
                              : !inMonth ? 'text-gray-200'
                              : isSun ? 'text-gray-300'
                              : 'text-gray-700'
                          }`}>
                            {d.getDate()}
                          </div>
                          {isSun && (
                            <div className="text-[7px] text-gray-300 mt-0.5 leading-none">หยุด</div>
                          )}
                        </div>
                      )
                    })}
                  </div>

                  {/* ── Invoice rows ── */}
                  <div className="relative">
                    {/* Sunday column — striped "closed" overlay */}
                    <div
                      className="absolute top-0 bottom-0 pointer-events-none z-0"
                      style={{
                        left: `calc(${LABEL_W}px + 6 * (100% - ${LABEL_W}px) / 7)`,
                        width: `calc((100% - ${LABEL_W}px) / 7)`,
                        backgroundColor: 'rgba(0,0,0,0.025)',
                        borderLeft: '1px solid rgba(0,0,0,0.04)',
                      }}
                    />

                    {/* Today column highlight spanning all rows of this week */}
                    {todayColIdx >= 0 && (
                      <div
                        className="absolute top-0 bottom-0 pointer-events-none z-0"
                        style={{
                          left: `calc(${LABEL_W}px + ${todayColIdx} * (100% - ${LABEL_W}px) / 7)`,
                          width: `calc((100% - ${LABEL_W}px) / 7)`,
                          background: 'rgba(219,234,254,0.35)',
                          borderLeft: '1px solid rgba(147,197,253,0.4)',
                          borderRight: '1px solid rgba(147,197,253,0.4)',
                        }}
                      />
                    )}

                    {weekInvs.length === 0 ? (
                      <div className="grid" style={{ gridTemplateColumns: GRID }}>
                        <div className="border-r border-slate-700/20 bg-slate-800/5" style={{ height: '28px' }} />
                        {days.map((_, di) => (
                          <div
                            key={di}
                            className="border-r last:border-r-0 border-gray-50"
                            style={{ height: '28px' }}
                          />
                        ))}
                      </div>
                    ) : (
                      weekInvs.map((inv, ii) => {
                        const arrStart = pd(inv.estimated_arrival!)
                        const arrEnd = pd(inv.estimated_arrival_end || inv.estimated_arrival!)
                        const cfg = S[inv.st] || S['อยู่ที่จีน']

                        // Sunday (dow=6) is closed — bars never enter that column.
                        // Clamp barStartDay to Mon (0) if source date falls on Sunday.
                        const rawStart = arrStart < mon ? 0 : dow(arrStart)
                        const barStartDay = rawStart === 6 ? 0 : rawStart // Sun start → treat as Mon next week (handled by nextWeek row)

                        // Clamp barEndDay to Sat (5) — never render into Sunday column.
                        const rawEnd = arrEnd > sun ? 5 : Math.min(dow(arrEnd), 5)
                        const barEndDay = rawEnd

                        const prevCont = arrStart < mon || dow(arrStart) === 6
                        // nextCont: extends past this week's Saturday (into Mon+)
                        const nextCont = arrEnd > sun || (arrEnd <= sun && dow(arrEnd) === 6)

                        // Border-radius: pill if starts/ends here, square if continues
                        const rl = prevCont ? '4px' : '18px'
                        const rr = nextCont ? '4px' : '18px'
                        const radius = `${rl} ${rr} ${rr} ${rl}`

                        // Left border accent: show only when bar starts this week
                        const leftBorderW = prevCont ? '1.5px' : '4px'

                        // CSS grid columns (1-indexed; 1=label, 2=Mon…8=Sun)
                        const colStart = gc(barStartDay)
                        const colEnd = gc(barEndDay) + 1

                        // Arriving today = today falls within the arrival date range
                        const arrivesToday =
                          todayStr >= inv.estimated_arrival! &&
                          todayStr <= (inv.estimated_arrival_end || inv.estimated_arrival!)


                        return (
                          <div
                            key={inv.id + ds(mon)}
                            className={`grid relative ${ii % 2 === 1 ? 'bg-slate-50/60' : 'bg-white'}`}
                            style={{ gridTemplateColumns: GRID, height: arrivesToday ? '48px' : '40px', zIndex: 1 }}
                          >
                            {/* Week label spacer */}
                            <div
                              className="border-r"
                              style={{
                                background: 'rgba(30,41,59,0.04)',
                                borderRightColor: 'rgba(30,41,59,0.12)',
                                gridRow: 1,
                              }}
                            />

                            {/* ── The bar ── */}
                            <Link
                              href={`/dashboard/${inv.id}`}
                              className={`group flex items-center overflow-hidden cursor-pointer relative z-10 transition-all duration-150 ${
                                arrivesToday
                                  ? 'hover:brightness-95'
                                  : 'hover:brightness-[0.96] hover:shadow-lg'
                              }`}
                              style={{
                                gridColumn: `${colStart} / ${colEnd}`,
                                gridRow: 1,
                                margin: '5px 4px',
                                height: arrivesToday ? 'calc(100% - 10px)' : 'calc(100% - 10px)',
                                borderRadius: radius,
                                background: arrivesToday ? cfg.stroke : cfg.fill,
                                border: arrivesToday
                                  ? `2px solid ${cfg.stroke}`
                                  : undefined,
                                borderTop: arrivesToday ? undefined : `1.5px solid ${cfg.stroke}50`,
                                borderRight: arrivesToday ? undefined : `1.5px solid ${cfg.stroke}50`,
                                borderBottom: arrivesToday ? undefined : `1.5px solid ${cfg.stroke}50`,
                                borderLeft: arrivesToday ? undefined : `${leftBorderW} solid ${cfg.stroke}`,
                                boxShadow: arrivesToday
                                  ? `0 0 0 3px ${cfg.stroke}30, 0 4px 16px ${cfg.stroke}50`
                                  : undefined,
                              }}
                              title={`${inv.invoice_no}${inv.supplier ? ' · ' + inv.supplier : ''} · ${inv.st}`}
                            >
                              {prevCont && (
                                <span
                                  className="text-[10px] pl-1.5 pr-0.5 shrink-0 opacity-60"
                                  style={{ color: arrivesToday ? 'white' : cfg.textColor }}
                                >◀</span>
                              )}
                              <div className="flex flex-col justify-center px-2 overflow-hidden flex-1 min-w-0">
                                <div className="flex items-center gap-1 min-w-0">
                                  <span
                                    className="truncate leading-none font-bold text-[11px]"
                                    style={{ color: arrivesToday ? 'white' : cfg.textColor }}
                                  >
                                    {inv.invoice_no}
                                  </span>
                                  {arrivesToday && (
                                    <span className="shrink-0 text-[8px] font-black bg-white/25 text-white rounded-full px-1.5 py-0.5 leading-none whitespace-nowrap">
                                      เข้าวันนี้!
                                    </span>
                                  )}
                                </div>
                                {inv.supplier && (
                                  <span
                                    className="text-[9px] truncate leading-none mt-0.5"
                                    style={{ color: arrivesToday ? 'rgba(255,255,255,0.75)' : cfg.textColor, opacity: arrivesToday ? 1 : 0.6 }}
                                  >
                                    {inv.supplier}
                                  </span>
                                )}
                              </div>
                              {nextCont && (
                                <span
                                  className="text-[10px] pr-1.5 pl-0.5 shrink-0 opacity-60"
                                  style={{ color: arrivesToday ? 'white' : cfg.textColor }}
                                >▶</span>
                              )}
                            </Link>

                          </div>
                        )
                      })
                    )}
                  </div>
                </div>
              )
            })}
          </div>
        )}

        {/* Legend — compact strip */}
        <div className="shrink-0 flex flex-wrap items-center gap-x-4 gap-y-1 px-1 py-1">
          <span className="text-[9px] font-black text-gray-400 uppercase tracking-widest">สถานะ</span>
          {Object.entries(S).map(([status, cfg]) => (
            <div key={status} className="flex items-center gap-1">
              <div
                className="w-3 h-3 rounded"
                style={{ background: cfg.fill, borderLeft: `3px solid ${cfg.stroke}` }}
              />
              <span className="text-[10px] font-semibold" style={{ color: cfg.textColor }}>{status}</span>
            </div>
          ))}
          <div className="flex items-center gap-1">
            <div className="w-5 h-3 rounded bg-blue-50 border border-blue-200" />
            <span className="text-[10px] text-gray-400">วันนี้</span>
          </div>
          <div className="flex items-center gap-1 text-[10px] text-gray-400">
            <span className="font-bold">◀▶</span>
            <span>ต่อเนื่องสัปดาห์อื่น</span>
          </div>
        </div>

      </div>
    </div>
  )
}

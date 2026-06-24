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
const DAYS_TH = ['จ','อ','พ','พฤ','ศ','ส','อา']

const STATUS_STYLE: Record<string, { pill: string; dot: string; label: string }> = {
  'อยู่ที่จีน':    { pill: 'bg-yellow-50 text-yellow-700 border-yellow-200',  dot: 'bg-yellow-400', label: 'อยู่ที่จีน' },
  'On board':      { pill: 'bg-blue-50 text-blue-700 border-blue-200',        dot: 'bg-blue-400',   label: 'On board' },
  'กำลังเข้าคลัง': { pill: 'bg-orange-50 text-orange-700 border-orange-200',  dot: 'bg-orange-400', label: 'กำลังเข้าคลัง' },
  'เข้าคลังแล้ว': { pill: 'bg-green-50 text-green-700 border-green-200',     dot: 'bg-green-400',  label: 'เข้าคลังแล้ว' },
}

function computeStatus(
  status: string | null,
  arrival: string | null,
  arrivalEnd: string | null,
  etaDate?: string | null,
): string {
  const base = status || 'อยู่ที่จีน'
  const normalized = (base === 'ถึงคลัง' || base === 'ถึงไทย กำลังเข้าคลัง') ? 'กำลังเข้าคลัง' : base
  const today = new Date(); today.setHours(0, 0, 0, 0)
  if (normalized === 'On board' && etaDate) {
    const eta = new Date(etaDate + 'T00:00:00'); eta.setHours(0, 0, 0, 0)
    if (today >= eta) return 'กำลังเข้าคลัง'
  }
  if (!arrival || normalized === 'อยู่ที่จีน') return normalized
  const checkDate = new Date(arrivalEnd || arrival); checkDate.setHours(0, 0, 0, 0)
  if (today > checkDate) return 'เข้าคลังแล้ว'
  return normalized
}

function toDateStr(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

export default function CalendarPage() {
  const [invoices, setInvoices] = useState<CalInvoice[]>([])
  const [loading, setLoading] = useState(true)
  const [curDate, setCurDate] = useState(() => {
    const d = new Date()
    return new Date(d.getFullYear(), d.getMonth(), 1)
  })
  const [selectedDay, setSelectedDay] = useState<string | null>(null)

  const today = useMemo(() => {
    const d = new Date(); d.setHours(0, 0, 0, 0); return d
  }, [])
  const todayStr = toDateStr(today)

  useEffect(() => { loadInvoices() }, [])

  async function loadInvoices() {
    const { data } = await supabase
      .from('invoices')
      .select('id, invoice_no, status, estimated_arrival, estimated_arrival_end, eta_date, supplier')
      .order('estimated_arrival', { ascending: true })
    if (data) setInvoices(data as CalInvoice[])
    setLoading(false)
  }

  const year = curDate.getFullYear()
  const month = curDate.getMonth()

  // Build calendar cells (Monday-first)
  const calendarCells = useMemo(() => {
    const firstDay = new Date(year, month, 1)
    const daysInMonth = new Date(year, month + 1, 0).getDate()
    const startOffset = (firstDay.getDay() + 6) % 7 // Mon=0, Sun=6
    const totalCells = Math.ceil((startOffset + daysInMonth) / 7) * 7
    return Array.from({ length: totalCells }, (_, i) => {
      const d = i - startOffset + 1
      return d >= 1 && d <= daysInMonth ? new Date(year, month, d) : null
    })
  }, [year, month])

  // Map dateStr → arrivals and ETAs
  const { arrivalMap, etaMap } = useMemo(() => {
    const arrivalMap: Record<string, { inv: CalInvoice; status: string }[]> = {}
    const etaMap: Record<string, CalInvoice[]> = {}
    for (const inv of invoices) {
      if (inv.estimated_arrival) {
        const key = inv.estimated_arrival
        if (!arrivalMap[key]) arrivalMap[key] = []
        arrivalMap[key].push({
          inv,
          status: computeStatus(inv.status, inv.estimated_arrival, inv.estimated_arrival_end, inv.eta_date),
        })
      }
      if (inv.eta_date) {
        if (!etaMap[inv.eta_date]) etaMap[inv.eta_date] = []
        etaMap[inv.eta_date].push(inv)
      }
    }
    return { arrivalMap, etaMap }
  }, [invoices])

  // Stats for current month
  const monthStr = `${year}-${String(month + 1).padStart(2, '0')}`
  const monthArrivals = invoices.filter(inv => inv.estimated_arrival?.startsWith(monthStr))
  const monthEtas = invoices.filter(inv => inv.eta_date?.startsWith(monthStr))

  // Selected-day panel data
  const selectedArrivals = selectedDay ? (arrivalMap[selectedDay] || []) : []
  const selectedEtas = selectedDay ? (etaMap[selectedDay] || []) : []

  function prevMonth() { setCurDate(new Date(year, month - 1, 1)); setSelectedDay(null) }
  function nextMonth() { setCurDate(new Date(year, month + 1, 1)); setSelectedDay(null) }
  function goToday() { setCurDate(new Date(today.getFullYear(), today.getMonth(), 1)); setSelectedDay(todayStr) }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-gray-50 to-slate-100">
      {/* Nav */}
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
        <div className="ml-auto">
          <LockButton />
        </div>
      </nav>

      <div className="max-w-6xl mx-auto px-4 sm:px-6 py-8">
        {/* Page header + month navigation */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
          <div>
            <h1 className="text-2xl font-bold text-gray-900 tracking-tight">ปฏิทินเข้าคลัง</h1>
            <p className="text-sm text-gray-400 mt-0.5">วันที่ประมาณการเข้าคลังของแต่ละ Invoice</p>
          </div>
          <div className="flex items-center gap-1">
            <button
              onClick={prevMonth}
              className="w-9 h-9 flex items-center justify-center rounded-xl text-gray-500 hover:bg-gray-100 transition-colors text-lg font-light"
            >
              ‹
            </button>
            <div className="min-w-[168px] text-center select-none">
              <span className="text-xl font-bold text-gray-800">{MONTHS_TH[month]}</span>
              <span className="text-lg font-semibold text-gray-400 ml-2">{year + 543}</span>
            </div>
            <button
              onClick={nextMonth}
              className="w-9 h-9 flex items-center justify-center rounded-xl text-gray-500 hover:bg-gray-100 transition-colors text-lg font-light"
            >
              ›
            </button>
            <button
              onClick={goToday}
              className="ml-2 px-3 py-1.5 text-xs rounded-lg bg-white border border-gray-200 shadow-sm hover:bg-gray-50 text-gray-600 font-medium transition-colors"
            >
              วันนี้
            </button>
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
            <div className="text-xs text-blue-500 font-medium mt-0.5">ETA ถึงท่าเรือไทยเดือนนี้</div>
          </div>
          <div className="rounded-2xl border border-gray-100 bg-white px-4 py-3 shadow-sm">
            <div className="text-2xl font-bold text-gray-700">{loading ? '—' : invoices.length}</div>
            <div className="text-xs text-gray-400 font-medium mt-0.5">Invoice ทั้งหมด</div>
          </div>
        </div>

        {/* Calendar + side panel */}
        <div className="flex gap-4 items-start">
          {/* Calendar grid */}
          <div className="flex-1 min-w-0 bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
            {/* Day-of-week header */}
            <div className="grid grid-cols-7 border-b border-gray-100 bg-gray-50/70">
              {DAYS_TH.map((d, i) => (
                <div
                  key={d}
                  className={`py-2.5 text-center text-[11px] font-bold uppercase tracking-widest ${
                    i === 6 ? 'text-red-400' : 'text-gray-400'
                  }`}
                >
                  {d}
                </div>
              ))}
            </div>

            {/* Day cells */}
            <div className="grid grid-cols-7">
              {calendarCells.map((day, i) => {
                const isSunday = i % 7 === 6
                const borderClasses = `border-b border-gray-100 ${!isSunday ? 'border-r border-gray-100' : ''}`

                if (!day) {
                  return <div key={i} className={`min-h-28 bg-gray-50/50 ${borderClasses}`} />
                }

                const dateStr = toDateStr(day)
                const arrivals = arrivalMap[dateStr] || []
                const etas = etaMap[dateStr] || []
                const isToday = dateStr === todayStr
                const isPast = day < today
                const isSelected = dateStr === selectedDay
                const hasEvents = arrivals.length > 0 || etas.length > 0
                const MAX_SHOW = etas.length > 0 ? 1 : 2

                return (
                  <div
                    key={i}
                    onClick={() => setSelectedDay(isSelected ? null : dateStr)}
                    className={`min-h-28 p-1.5 cursor-pointer transition-all group ${borderClasses} ${
                      isSelected
                        ? 'bg-blue-50/60 ring-1 ring-inset ring-blue-200'
                        : isToday
                        ? 'bg-blue-50/20'
                        : isPast && !hasEvents
                        ? 'bg-gray-50/40'
                        : 'hover:bg-gray-50/70'
                    }`}
                  >
                    {/* Day number */}
                    <div className={`flex ${hasEvents ? 'justify-between' : 'justify-end'} items-start mb-1`}>
                      {hasEvents && (
                        <div className="flex gap-0.5 pt-0.5">
                          {etas.length > 0 && <span className="w-1.5 h-1.5 rounded-full bg-blue-400 mt-0.5" />}
                          {arrivals.length > 0 && (
                            <span className={`w-1.5 h-1.5 rounded-full mt-0.5 ${
                              STATUS_STYLE[arrivals[0].status]?.dot ?? 'bg-gray-400'
                            }`} />
                          )}
                        </div>
                      )}
                      <span
                        className={`text-[11px] font-bold w-5 h-5 flex items-center justify-center rounded-full transition-colors ${
                          isToday
                            ? 'bg-blue-600 text-white'
                            : isSunday
                            ? 'text-red-400 group-hover:bg-red-50'
                            : isPast
                            ? 'text-gray-300'
                            : 'text-gray-600 group-hover:bg-gray-100'
                        }`}
                      >
                        {day.getDate()}
                      </span>
                    </div>

                    {/* ETA chip */}
                    {etas.slice(0, 1).map(inv => (
                      <div
                        key={'eta-' + inv.id}
                        onClick={e => e.stopPropagation()}
                        className="mb-0.5"
                      >
                        <Link
                          href={`/dashboard/${inv.id}`}
                          className="flex items-center gap-1 px-1.5 py-0.5 rounded-md bg-blue-50 border border-blue-100 text-blue-700 text-[10px] truncate hover:bg-blue-100 transition-colors"
                          title={`ETA: ${inv.invoice_no}`}
                        >
                          <span className="shrink-0">🚢</span>
                          <span className="truncate font-semibold">{inv.invoice_no}</span>
                        </Link>
                      </div>
                    ))}
                    {etas.length > 1 && (
                      <div className="text-[10px] text-blue-400 pl-1 mb-0.5">+{etas.length - 1} ETA</div>
                    )}

                    {/* Arrival chips */}
                    {arrivals.slice(0, MAX_SHOW).map(({ inv, status }) => {
                      const st = STATUS_STYLE[status] || STATUS_STYLE['อยู่ที่จีน']
                      const hasRange = inv.estimated_arrival_end && inv.estimated_arrival_end !== inv.estimated_arrival
                      return (
                        <div
                          key={inv.id}
                          onClick={e => e.stopPropagation()}
                          className="mb-0.5"
                        >
                          <Link
                            href={`/dashboard/${inv.id}`}
                            className={`flex items-center gap-1 px-1.5 py-0.5 rounded-md border text-[10px] truncate hover:opacity-80 transition-opacity ${st.pill}`}
                            title={`${inv.invoice_no} · ${status}${hasRange ? ` · ถึง ${inv.estimated_arrival_end}` : ''}${inv.supplier ? ` · ${inv.supplier}` : ''}`}
                          >
                            <span className={`w-1.5 h-1.5 rounded-full shrink-0 ${st.dot}`} />
                            <span className="truncate font-semibold">{inv.invoice_no}</span>
                            {hasRange && <span className="opacity-50 shrink-0 text-[9px]">›</span>}
                          </Link>
                        </div>
                      )
                    })}

                    {/* Overflow */}
                    {arrivals.length > MAX_SHOW && (
                      <div className="text-[10px] text-gray-400 text-right pr-0.5 mt-0.5">
                        +{arrivals.length - MAX_SHOW} รายการ
                      </div>
                    )}
                  </div>
                )
              })}
            </div>
          </div>

          {/* Side panel: selected day details */}
          <div className={`w-72 shrink-0 transition-all duration-200 ${selectedDay ? 'opacity-100' : 'opacity-0 pointer-events-none'}`}>
            {selectedDay && (
              <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden sticky top-20">
                <div className="px-4 py-3 border-b border-gray-100 bg-gray-50/70 flex items-center justify-between">
                  <div>
                    <div className="text-xs text-gray-400 font-medium uppercase tracking-wider">รายละเอียดวันที่</div>
                    <div className="text-sm font-bold text-gray-800 mt-0.5">
                      {new Date(selectedDay + 'T00:00:00').toLocaleDateString('th-TH', {
                        weekday: 'long', day: 'numeric', month: 'long', year: 'numeric',
                      })}
                    </div>
                  </div>
                  <button
                    onClick={() => setSelectedDay(null)}
                    className="w-6 h-6 rounded-full hover:bg-gray-200 flex items-center justify-center text-gray-400 text-xs transition-colors"
                  >
                    ✕
                  </button>
                </div>

                <div className="px-4 py-3 space-y-4 max-h-[70vh] overflow-y-auto">
                  {/* ETA section */}
                  {selectedEtas.length > 0 && (
                    <div>
                      <div className="flex items-center gap-1.5 mb-2">
                        <span className="text-sm">🚢</span>
                        <span className="text-xs font-bold text-blue-700 uppercase tracking-wider">ETA ถึงท่าเรือ</span>
                      </div>
                      <div className="space-y-1.5">
                        {selectedEtas.map(inv => (
                          <Link
                            key={inv.id}
                            href={`/dashboard/${inv.id}`}
                            className="block p-2.5 rounded-xl bg-blue-50 border border-blue-100 hover:bg-blue-100 transition-colors"
                          >
                            <div className="font-bold text-sm text-blue-800">{inv.invoice_no}</div>
                            {inv.supplier && (
                              <div className="text-xs text-blue-500 mt-0.5 truncate">{inv.supplier}</div>
                            )}
                          </Link>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Arrival section */}
                  {selectedArrivals.length > 0 && (
                    <div>
                      <div className="flex items-center gap-1.5 mb-2">
                        <span className="text-sm">📦</span>
                        <span className="text-xs font-bold text-gray-500 uppercase tracking-wider">เข้าคลัง</span>
                      </div>
                      <div className="space-y-1.5">
                        {selectedArrivals.map(({ inv, status }) => {
                          const st = STATUS_STYLE[status] || STATUS_STYLE['อยู่ที่จีน']
                          const hasRange = inv.estimated_arrival_end && inv.estimated_arrival_end !== inv.estimated_arrival
                          return (
                            <Link
                              key={inv.id}
                              href={`/dashboard/${inv.id}`}
                              className={`block p-2.5 rounded-xl border hover:opacity-80 transition-opacity ${st.pill}`}
                            >
                              <div className="flex items-center justify-between gap-2">
                                <span className="font-bold text-sm">{inv.invoice_no}</span>
                                <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded-full border ${st.pill}`}>
                                  {status}
                                </span>
                              </div>
                              {inv.supplier && (
                                <div className="text-[11px] opacity-70 mt-0.5 truncate">{inv.supplier}</div>
                              )}
                              {hasRange && (
                                <div className="text-[11px] opacity-60 mt-0.5">
                                  ถึง {new Date(inv.estimated_arrival_end! + 'T00:00:00').toLocaleDateString('th-TH', { day: 'numeric', month: 'short' })}
                                </div>
                              )}
                            </Link>
                          )
                        })}
                      </div>
                    </div>
                  )}

                  {selectedArrivals.length === 0 && selectedEtas.length === 0 && (
                    <div className="text-center py-8 text-gray-300">
                      <div className="text-3xl mb-2">📅</div>
                      <div className="text-sm">ไม่มีรายการวันนี้</div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Legend */}
        <div className="mt-4 flex flex-wrap items-center gap-x-5 gap-y-2">
          <span className="text-[11px] text-gray-400 font-bold uppercase tracking-widest">สัญลักษณ์</span>
          {Object.values(STATUS_STYLE).map(({ dot, label }) => (
            <div key={label} className="flex items-center gap-1.5">
              <span className={`w-2 h-2 rounded-full ${dot}`} />
              <span className="text-xs text-gray-500">{label}</span>
            </div>
          ))}
          <div className="flex items-center gap-1.5">
            <span className="text-xs">🚢</span>
            <span className="text-xs text-gray-500">ETA ถึงท่าเรือไทย</span>
          </div>
          <div className="flex items-center gap-1.5">
            <span className="text-xs">›</span>
            <span className="text-xs text-gray-500">มีช่วงหลายวัน</span>
          </div>
        </div>
      </div>
    </div>
  )
}

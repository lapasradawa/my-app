'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'
import type { ResultRow } from '@/lib/excel-parser'
import { isUnlocked } from '@/lib/auth'
import LockButton from '@/components/LockButton'
import PasswordModal from '@/components/PasswordModal'

const STATUSES = ['อยู่ที่จีน', 'On board', 'กำลังเข้าคลัง'] as const
type Status = typeof STATUSES[number]

const STATUS_STYLE: Record<string, string> = {
  'อยู่ที่จีน': 'bg-yellow-100 text-yellow-800 border-yellow-300',
  'On board': 'bg-blue-100 text-blue-800 border-blue-300',
  'กำลังเข้าคลัง': 'bg-orange-100 text-orange-800 border-orange-300',
  'ถึงไทย กำลังเข้าคลัง': 'bg-orange-100 text-orange-800 border-orange-300', // legacy
  'ถึงคลัง': 'bg-orange-100 text-orange-800 border-orange-300', // legacy
  'เข้าคลังแล้ว': 'bg-green-100 text-green-800 border-green-300',
}

const STATUS_BADGE: Record<string, string> = {
  'อยู่ที่จีน': 'bg-yellow-100 text-yellow-800',
  'On board': 'bg-blue-100 text-blue-800',
  'กำลังเข้าคลัง': 'bg-orange-100 text-orange-800',
  'ถึงไทย กำลังเข้าคลัง': 'bg-orange-100 text-orange-800', // legacy
  'ถึงคลัง': 'bg-orange-100 text-orange-800', // legacy
  'เข้าคลังแล้ว': 'bg-green-100 text-green-800',
}

// คำนวณ display status โดยอัตโนมัติ
// - ถ้า On board และพ้น ETA แล้ว → กำลังเข้าคลัง
// - ถ้าพ้นวันประมาณการเข้าคลังแล้ว → เข้าคลังแล้ว
function computeStatus(status: string | null, arrival: string | null, arrivalEnd: string | null, etaDate?: string | null): string {
  const base = status || 'อยู่ที่จีน'
  const normalized = (base === 'ถึงคลัง' || base === 'ถึงไทย กำลังเข้าคลัง') ? 'กำลังเข้าคลัง' : base
  const today = new Date(); today.setHours(0, 0, 0, 0)
  // Check arrival date FIRST — if today is past the start of the arrival window, goods are in warehouse
  if (arrival && normalized !== 'อยู่ที่จีน') {
    const checkDate = new Date(arrival + 'T00:00:00'); checkDate.setHours(0, 0, 0, 0)
    if (today > checkDate) return 'เข้าคลังแล้ว'
  }
  // On board + past ETA → กำลังเข้าคลัง
  if (normalized === 'On board' && etaDate) {
    const eta = new Date(etaDate + 'T00:00:00'); eta.setHours(0, 0, 0, 0)
    if (today >= eta) return 'กำลังเข้าคลัง'
  }
  if (!arrival || normalized === 'อยู่ที่จีน') return normalized
  return normalized
}

interface Invoice {
  id: string
  invoice_no: string
  filename: string
  created_at: string
  status: Status | null
  estimated_arrival: string | null
  estimated_arrival_end: string | null
  eta_date: string | null
  bl_date: string | null
  payment_status: string | null
  payment_date: string | null
  supplier: string | null
}

function computeDueDate(blDate: string | null): Date | null {
  if (!blDate) return null
  const d = new Date(blDate)
  d.setDate(d.getDate() + 30)
  return d
}

function getPaymentLabel(payment_status: string | null, bl_date: string | null): { label: string; colorClass: string } {
  if (payment_status === 'paid') return { label: 'จ่ายแล้ว', colorClass: 'bg-green-100 text-green-800' }
  const due = computeDueDate(bl_date)
  if (!due) return { label: '—', colorClass: 'bg-gray-100 text-gray-400' }
  const today = new Date(); today.setHours(0, 0, 0, 0)
  const dueDay = new Date(due); dueDay.setHours(0, 0, 0, 0)
  const diff = Math.ceil((dueDay.getTime() - today.getTime()) / 86400000)
  if (diff < 0) return { label: 'Overdue', colorClass: 'bg-red-100 text-red-700' }
  if (diff <= 3) return { label: 'ใกล้ถึง due date', colorClass: 'bg-yellow-100 text-yellow-800' }
  return { label: 'ยังไม่จ่าย', colorClass: 'bg-gray-100 text-gray-500' }
}

interface LocalEdit {
  status: Status
  estimated_arrival: string
  estimated_arrival_end: string
  eta_date: string
  supplier: string
}

interface SearchResult {
  id: string
  invoice_no: string
  status: string
  estimated_arrival: string | null
  estimated_arrival_end: string | null
  totalQty: number
  matchingRows: { no: number; code: string; po: string; qty: number }[]
}

function fmtDate(d: string | null): string {
  if (!d) return ''
  return new Date(d + 'T00:00:00').toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
}

function fmtRange(start: string | null, end: string | null): string {
  if (!start) return '—'
  if (!end || end === start) return fmtDate(start)
  return `${fmtDate(start)} – ${fmtDate(end)}`
}

export default function DashboardPage() {
  const [invoices, setInvoices] = useState<Invoice[]>([])
  const [edits, setEdits] = useState<Record<string, LocalEdit>>({})
  const [saving, setSaving] = useState<Record<string, boolean>>({})
  const [saved, setSaved] = useState<Record<string, boolean>>({})
  const [loading, setLoading] = useState(true)
  const [searchCode, setSearchCode] = useState('')
  const [searchResults, setSearchResults] = useState<SearchResult[] | null>(null)
  const [searching, setSearching] = useState(false)
  const [unlocked, setUnlocked] = useState(false)
  const [showPasswordModal, setShowPasswordModal] = useState(false)
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null)

  useEffect(() => { loadInvoices(); setUnlocked(isUnlocked()) }, [])

  function requireUnlock(action: () => void) {
    if (isUnlocked()) { action() }
    else { setPendingAction(() => action); setShowPasswordModal(true) }
  }

  async function loadInvoices() {
    setLoading(true)
    const { data } = await supabase
      .from('invoices')
      .select('id, invoice_no, filename, created_at, status, estimated_arrival, estimated_arrival_end, eta_date, bl_date, payment_status, payment_date, supplier')
      .order('estimated_arrival', { ascending: false, nullsFirst: false })
    if (data) {
      setInvoices(data as Invoice[])
      const initial: Record<string, LocalEdit> = {}
      for (const inv of data) {
        initial[inv.id] = {
          status: (inv.status as Status) || 'อยู่ที่จีน',
          estimated_arrival: inv.estimated_arrival || '',
          estimated_arrival_end: inv.estimated_arrival_end || '',
          eta_date: inv.eta_date || '',
          supplier: inv.supplier || '',
        }
      }
      setEdits(initial)
    }
    setLoading(false)
  }

  function setField<K extends keyof LocalEdit>(id: string, key: K, value: LocalEdit[K]) {
    setEdits(e => ({ ...e, [id]: { ...e[id], [key]: value } }))
  }

  function isDirty(inv: Invoice) {
    const e = edits[inv.id]
    if (!e) return false
    return (
      e.status !== ((inv.status as Status) || 'อยู่ที่จีน') ||
      e.estimated_arrival !== (inv.estimated_arrival || '') ||
      e.estimated_arrival_end !== (inv.estimated_arrival_end || '') ||
      e.eta_date !== (inv.eta_date || '') ||
      e.supplier !== (inv.supplier || '')
    )
  }

  async function saveRow(id: string) {
    const e = edits[id]
    if (!e) return
    setSaving(s => ({ ...s, [id]: true }))
    const noDate = e.status === 'อยู่ที่จีน'
    await supabase.from('invoices').update({
      status: e.status,
      estimated_arrival: noDate ? null : (e.estimated_arrival || null),
      estimated_arrival_end: noDate ? null : (e.estimated_arrival_end || null),
      eta_date: e.eta_date || null,
      supplier: e.supplier || null,
    }).eq('id', id)
    setInvoices(prev => prev.map(inv => inv.id === id ? {
      ...inv,
      status: e.status,
      estimated_arrival: noDate ? null : (e.estimated_arrival || null),
      estimated_arrival_end: noDate ? null : (e.estimated_arrival_end || null),
      eta_date: e.eta_date || null,
      supplier: e.supplier || null,
    } : inv))
    setSaving(s => ({ ...s, [id]: false }))
    setSaved(s => ({ ...s, [id]: true }))
    setTimeout(() => setSaved(s => ({ ...s, [id]: false })), 2000)
  }

  const formatDate = (iso: string) => new Date(iso).toLocaleDateString('en-GB', {
    day: '2-digit', month: 'short', year: 'numeric'
  })

  async function handleSearch() {
    const term = searchCode.trim()
    if (!term) return
    setSearching(true)
    setSearchResults(null)
    const { data } = await supabase
      .from('invoices')
      .select('id, invoice_no, status, estimated_arrival, estimated_arrival_end, eta_date, rows')
    if (data) {
      const results: SearchResult[] = []
      for (const inv of data) {
        const matching = (inv.rows as ResultRow[]).filter(r =>
          r.code.toLowerCase().includes(term.toLowerCase())
        )
        if (matching.length > 0) {
          results.push({
            id: inv.id,
            invoice_no: inv.invoice_no,
            status: computeStatus(inv.status, inv.estimated_arrival, inv.estimated_arrival_end, inv.eta_date),
            estimated_arrival: inv.estimated_arrival,
            estimated_arrival_end: inv.estimated_arrival_end,
            totalQty: matching.reduce((s, r) => s + r.qty, 0),
            matchingRows: matching.map(r => ({ no: r.no, code: r.code, po: r.po, qty: r.qty })),
          })
        }
      }
      setSearchResults(results)
    }
    setSearching(false)
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-blue-600 font-semibold">Dashboard</Link>
        <Link href="/calendar" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Calendar</Link>
        <Link href="/report" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/summary" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Summary</Link>
        <Link href="/compare" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Builder</Link>
        <Link href="/qc" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">QC Report</Link>
        <Link href="/guide" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Guide</Link>
        <div className="ml-auto">
          <LockButton onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />
        </div>
      </nav>
      {showPasswordModal && (
        <PasswordModal
          onSuccess={() => { setUnlocked(true); setShowPasswordModal(false); pendingAction?.(); setPendingAction(null) }}
          onCancel={() => { setShowPasswordModal(false); setPendingAction(null) }}
        />
      )}

      <div className="max-w-6xl mx-auto px-6 py-8">
        <div className="mb-6 flex items-end justify-between">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">Dashboard</h1>
            <p className="text-sm text-gray-500 mt-1">ติดตามสถานะและวันที่ประมาณการเข้าคลัง</p>
          </div>
          <p className="text-sm text-gray-500">
            วันนี้: <span className="font-semibold text-gray-700">
              {new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}
            </span>
          </p>
        </div>

        {/* Item Code Search */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 mb-6">
          <p className="text-sm font-semibold text-gray-700 mb-3">ค้นหา Item Code</p>
          <div className="flex gap-2">
            <input
              type="text"
              placeholder="พิมพ์ code สินค้า เช่น P46P090..."
              value={searchCode}
              onChange={e => setSearchCode(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && handleSearch()}
              className="flex-1 border border-gray-300 rounded-lg px-3 py-2 text-sm outline-none focus:border-blue-400"
            />
            <button
              onClick={handleSearch}
              disabled={searching || !searchCode.trim()}
              className="px-4 py-2 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700 disabled:opacity-50 transition-colors"
            >
              {searching ? 'กำลังค้น...' : 'ค้นหา'}
            </button>
            {searchResults !== null && (
              <button
                onClick={() => { setSearchResults(null); setSearchCode('') }}
                className="px-3 py-2 text-gray-500 hover:text-gray-700 text-sm border border-gray-200 rounded-lg"
              >
                ล้าง
              </button>
            )}
          </div>

          {searchResults !== null && (
            <div className="mt-4">
              {searchResults.length === 0 ? (
                <p className="text-sm text-gray-400 text-center py-4">ไม่พบ code นี้ในระบบ</p>
              ) : (
                <>
                  <p className="text-xs text-gray-500 mb-3">พบใน {searchResults.length} Invoice</p>
                  <div className="space-y-3">
                    {searchResults.map(r => (
                      <div key={r.id} className="border border-gray-100 rounded-lg p-3 bg-gray-50">
                        <div className="flex items-center gap-3 mb-2 flex-wrap">
                          <Link href={`/dashboard/${r.id}`} className="font-semibold text-blue-600 hover:underline text-sm">
                            {r.invoice_no}
                          </Link>
                          <span className={`text-xs font-medium px-2 py-0.5 rounded-full ${STATUS_BADGE[r.status] || 'bg-gray-100 text-gray-700'}`}>
                            {r.status}
                          </span>
                          {r.estimated_arrival && (
                            <span className="text-xs text-gray-500">
                              เข้าคลัง: <strong>{fmtRange(r.estimated_arrival, r.estimated_arrival_end)}</strong>
                            </span>
                          )}
                          <span className="text-xs text-gray-500 ml-auto">รวม QTY: <strong>{r.totalQty.toLocaleString()}</strong></span>
                        </div>
                        <table className="w-full text-xs text-gray-600">
                          <thead>
                            <tr className="text-gray-400">
                              <th className="text-left pb-1">Code</th>
                              <th className="text-left pb-1">PO</th>
                              <th className="text-right pb-1">QTY</th>
                            </tr>
                          </thead>
                          <tbody>
                            {r.matchingRows.map(row => (
                              <tr key={row.no}>
                                <td className="font-mono pr-4 py-0.5">{row.code}</td>
                                <td className="pr-4 py-0.5 text-gray-500">{row.po || '-'}</td>
                                <td className="text-right font-medium">{row.qty.toLocaleString()}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>
          )}
        </div>

        {/* Invoice Table */}
        {loading ? (
          <div className="text-center py-12 text-gray-400">กำลังโหลด...</div>
        ) : invoices.length === 0 ? (
          <div className="text-center py-12 text-gray-400">
            ยังไม่มี Invoice — <Link href="/" className="text-blue-600 underline">ไปอัปโหลด</Link>
          </div>
        ) : (
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-auto max-h-[calc(100vh-320px)]">
            <table className="w-full text-sm table-fixed">
              <thead className="sticky top-0 z-10">
                <tr className="bg-gray-50 text-gray-600 border-b border-gray-200">
                  <th className="px-4 py-3 text-left font-medium w-[13%]">Invoice No.</th>
                  <th className="px-4 py-3 text-left font-medium w-[10%]">Supplier</th>
                  <th className="px-4 py-3 text-left font-medium w-[11%]">สถานะ</th>
                  <th className="px-4 py-3 text-left font-medium w-[11%]"><span className="whitespace-nowrap">วันที่ถึงท่าเรือไทย</span><br/>(ETA)</th>
                  <th className="px-4 py-3 text-left font-medium w-[18%]">ประมาณการเข้าคลัง</th>
                  <th className="px-4 py-3 text-left font-medium w-[11%]">วันที่บันทึก</th>
                  <th className="px-4 py-3 text-left font-medium w-[10%]">Due Date</th>
                  <th className="px-4 py-3 text-left font-medium w-[11%]">สถานะจ่าย</th>
                  <th className="px-4 py-3 w-[8%]"></th>
                </tr>
              </thead>
              <tbody>
                {invoices.map((inv) => {
                  const e = edits[inv.id] || { status: 'อยู่ที่จีน' as Status, estimated_arrival: '', estimated_arrival_end: '' }
                  const showDate = e.status === 'On board' || e.status === 'กำลังเข้าคลัง' || e.status === 'ถึงไทย กำลังเข้าคลัง' || e.status === 'ถึงคลัง'
                  const displaySt = computeStatus(inv.status, inv.estimated_arrival, inv.estimated_arrival_end, inv.eta_date)
                  return (
                    <tr key={inv.id} className="border-b border-gray-100 last:border-0 hover:bg-gray-50/60 transition-colors">
                      <td className="px-4 py-3">
                        <Link href={`/dashboard/${inv.id}`} className="font-semibold text-blue-600 hover:text-blue-800 hover:underline">
                          {inv.invoice_no || '-'}
                        </Link>
                      </td>
                      <td className="px-4 py-3">
                        {unlocked ? (
                          <input
                            type="text"
                            value={edits[inv.id]?.supplier ?? ''}
                            onChange={ev => setField(inv.id, 'supplier', ev.target.value)}
                            placeholder="ใส่ชื่อ supplier"
                            className="text-xs border border-gray-200 rounded px-2 py-1 outline-none focus:border-blue-400 w-36 text-gray-700"
                          />
                        ) : (
                          <span className="text-xs text-gray-700">{inv.supplier || <span className="text-gray-300">—</span>}</span>
                        )}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        {unlocked && displaySt !== 'เข้าคลังแล้ว' ? (
                          <select
                            value={e.status === 'ถึงคลัง' || e.status === 'ถึงไทย กำลังเข้าคลัง' ? 'กำลังเข้าคลัง' : e.status}
                            onChange={ev => setField(inv.id, 'status', ev.target.value as Status)}
                            className={`text-xs font-medium px-2 py-1 rounded-full border cursor-pointer outline-none ${STATUS_STYLE[e.status] || ''}`}
                          >
                            {STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
                          </select>
                        ) : (
                          <span className={`text-xs font-medium px-2 py-1 rounded-full border whitespace-nowrap ${STATUS_STYLE[displaySt] || 'bg-gray-100 text-gray-700 border-gray-300'}`}>
                            {displaySt}
                          </span>
                        )}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        {unlocked ? (
                          <input
                            type="date"
                            value={edits[inv.id]?.eta_date ?? ''}
                            onChange={ev => setField(inv.id, 'eta_date', ev.target.value)}
                            className="text-xs border border-gray-300 rounded px-2 py-1 outline-none focus:border-blue-400 text-gray-700"
                          />
                        ) : (
                          <span className="text-xs text-gray-600">{fmtDate(inv.eta_date) || <span className="text-gray-300">—</span>}</span>
                        )}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        {unlocked && showDate ? (
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <input
                              type="date"
                              value={e.estimated_arrival}
                              onChange={ev => setField(inv.id, 'estimated_arrival', ev.target.value)}
                              className="text-xs border border-gray-300 rounded px-2 py-1 outline-none focus:border-blue-400 text-gray-700"
                            />
                            <span className="text-xs text-gray-400">–</span>
                            <input
                              type="date"
                              value={e.estimated_arrival_end}
                              min={e.estimated_arrival || undefined}
                              onChange={ev => setField(inv.id, 'estimated_arrival_end', ev.target.value)}
                              className="text-xs border border-gray-300 rounded px-2 py-1 outline-none focus:border-blue-400 text-gray-700"
                            />
                            {e.estimated_arrival_end && (
                              <button onClick={() => setField(inv.id, 'estimated_arrival_end', '')} className="text-gray-300 hover:text-gray-500 text-xs">✕</button>
                            )}
                          </div>
                        ) : (
                          <span className="text-xs text-gray-600 whitespace-nowrap">
                            {fmtRange(inv.estimated_arrival, inv.estimated_arrival_end) || '—'}
                          </span>
                        )}
                      </td>
                      <td className="px-4 py-3 text-gray-500 text-xs whitespace-nowrap">{formatDate(inv.created_at)}</td>
                      <td className="px-4 py-3 whitespace-nowrap text-xs text-gray-600">
                        {(() => {
                          const due = computeDueDate(inv.bl_date)
                          if (!due) return <span className="text-gray-300">—</span>
                          return due.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
                        })()}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        {(() => {
                          const p = getPaymentLabel(inv.payment_status, inv.bl_date)
                          if (p.label === '—') return <span className="text-gray-300 text-xs">—</span>
                          return (
                            <span className={`text-xs font-medium px-2 py-0.5 rounded-full whitespace-nowrap ${p.colorClass}`}>
                              {p.label}
                            </span>
                          )
                        })()}
                      </td>
                      <td className="px-4 py-3">
                        {unlocked && (
                          saved[inv.id] ? (
                            <span className="text-xs text-green-600 font-medium">✓ บันทึกแล้ว</span>
                          ) : (
                            <button
                              onClick={() => requireUnlock(() => saveRow(inv.id))}
                              disabled={saving[inv.id] || !isDirty(inv)}
                              className={`text-xs px-3 py-1 rounded-lg transition-colors ${
                                isDirty(inv) ? 'bg-blue-600 text-white hover:bg-blue-700' : 'bg-gray-100 text-gray-400 cursor-default'
                              }`}
                            >
                              {saving[inv.id] ? 'บันทึก...' : 'บันทึก'}
                            </button>
                          )
                        )}
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* Summary cards */}
        {invoices.length > 0 && (
          <div className="mt-6 grid grid-cols-4 gap-4">
            {([...STATUSES, 'เข้าคลังแล้ว'] as string[]).map(s => {
              const count = invoices.filter(inv =>
                computeStatus(inv.status, inv.estimated_arrival, inv.estimated_arrival_end, inv.eta_date) === s
              ).length
              return (
                <div key={s} className="rounded-xl p-4 border border-gray-200 bg-white">
                  <p className="text-2xl font-bold text-gray-800">{count}</p>
                  <p className="text-sm text-gray-600 mt-1">{s}</p>
                </div>
              )
            })}
          </div>
        )}
      </div>
    </div>
  )
}

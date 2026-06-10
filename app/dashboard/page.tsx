'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'

const STATUSES = ['อยู่ที่จีน', 'On board', 'ถึงคลัง'] as const
type Status = typeof STATUSES[number]

const STATUS_STYLE: Record<Status, string> = {
  'อยู่ที่จีน': 'bg-yellow-100 text-yellow-800 border-yellow-300',
  'On board': 'bg-blue-100 text-blue-800 border-blue-300',
  'ถึงคลัง': 'bg-green-100 text-green-800 border-green-300',
}

interface Invoice {
  id: string
  invoice_no: string
  filename: string
  created_at: string
  status: Status | null
  estimated_arrival: string | null
}

interface LocalEdit {
  status: Status
  estimated_arrival: string
}

export default function DashboardPage() {
  const [invoices, setInvoices] = useState<Invoice[]>([])
  const [edits, setEdits] = useState<Record<string, LocalEdit>>({})
  const [saving, setSaving] = useState<Record<string, boolean>>({})
  const [saved, setSaved] = useState<Record<string, boolean>>({})
  const [loading, setLoading] = useState(true)

  useEffect(() => { loadInvoices() }, [])

  async function loadInvoices() {
    setLoading(true)
    const { data } = await supabase
      .from('invoices')
      .select('id, invoice_no, filename, created_at, status, estimated_arrival')
      .order('created_at', { ascending: false })
    if (data) {
      setInvoices(data as Invoice[])
      const initial: Record<string, LocalEdit> = {}
      for (const inv of data) {
        initial[inv.id] = {
          status: (inv.status as Status) || 'อยู่ที่จีน',
          estimated_arrival: inv.estimated_arrival || '',
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
    const origStatus = (inv.status as Status) || 'อยู่ที่จีน'
    const origDate = inv.estimated_arrival || ''
    return e.status !== origStatus || e.estimated_arrival !== origDate
  }

  async function saveRow(id: string) {
    const e = edits[id]
    if (!e) return
    setSaving(s => ({ ...s, [id]: true }))
    await supabase.from('invoices').update({
      status: e.status,
      estimated_arrival: e.status === 'อยู่ที่จีน' ? null : (e.estimated_arrival || null),
    }).eq('id', id)
    setInvoices(prev => prev.map(inv => inv.id === id
      ? { ...inv, status: e.status, estimated_arrival: e.status === 'อยู่ที่จีน' ? null : (e.estimated_arrival || null) }
      : inv
    ))
    setSaving(s => ({ ...s, [id]: false }))
    setSaved(s => ({ ...s, [id]: true }))
    setTimeout(() => setSaved(s => ({ ...s, [id]: false })), 2000)
  }

  const formatDate = (iso: string) => new Date(iso).toLocaleDateString('th-TH', {
    day: '2-digit', month: 'short', year: 'numeric'
  })

  const formatArrival = (d: string | null) => {
    if (!d) return '-'
    const [y, m, day] = d.split('-')
    return `${day}/${m}/${y}`
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-blue-600 font-semibold">Dashboard</Link>
      </nav>

      <div className="max-w-6xl mx-auto px-6 py-8">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">Dashboard</h1>
          <p className="text-sm text-gray-500 mt-1">ติดตามสถานะและวันที่ประมาณการเข้าคลัง</p>
        </div>

        {loading ? (
          <div className="text-center py-12 text-gray-400">กำลังโหลด...</div>
        ) : invoices.length === 0 ? (
          <div className="text-center py-12 text-gray-400">
            ยังไม่มี Invoice —{' '}
            <Link href="/" className="text-blue-600 underline">ไปอัปโหลด</Link>
          </div>
        ) : (
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-gray-50 text-gray-600 border-b border-gray-200">
                  <th className="px-4 py-3 text-left font-medium">Invoice No.</th>
                  <th className="px-4 py-3 text-left font-medium">ไฟล์</th>
                  <th className="px-4 py-3 text-left font-medium">วันที่บันทึก</th>
                  <th className="px-4 py-3 text-left font-medium">สถานะ</th>
                  <th className="px-4 py-3 text-left font-medium">ประมาณการเข้าคลัง</th>
                  <th className="px-4 py-3 text-left font-medium w-28"></th>
                </tr>
              </thead>
              <tbody>
                {invoices.map((inv) => {
                  const e = edits[inv.id] || { status: 'อยู่ที่จีน' as Status, estimated_arrival: '' }
                  const showDate = e.status === 'On board' || e.status === 'ถึงคลัง'
                  return (
                    <tr key={inv.id} className="border-b border-gray-100 last:border-0 hover:bg-gray-50/60 transition-colors">
                      <td className="px-4 py-3">
                        <Link
                          href={`/dashboard/${inv.id}`}
                          className="font-semibold text-blue-600 hover:text-blue-800 hover:underline"
                        >
                          {inv.invoice_no || '-'}
                        </Link>
                      </td>
                      <td className="px-4 py-3 text-gray-500 text-xs max-w-[160px] truncate">{inv.filename || '-'}</td>
                      <td className="px-4 py-3 text-gray-500 text-xs whitespace-nowrap">{formatDate(inv.created_at)}</td>
                      <td className="px-4 py-3">
                        <select
                          value={e.status}
                          onChange={ev => setField(inv.id, 'status', ev.target.value as Status)}
                          className={`text-xs font-medium px-2 py-1 rounded-full border cursor-pointer outline-none ${STATUS_STYLE[e.status]}`}
                        >
                          {STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
                        </select>
                      </td>
                      <td className="px-4 py-3">
                        {showDate ? (
                          <input
                            type="date"
                            value={e.estimated_arrival}
                            onChange={ev => setField(inv.id, 'estimated_arrival', ev.target.value)}
                            className="text-xs border border-gray-300 rounded px-2 py-1 outline-none focus:border-blue-400 text-gray-700"
                          />
                        ) : (
                          <span className="text-xs text-gray-400 italic">—</span>
                        )}
                      </td>
                      <td className="px-4 py-3">
                        {saved[inv.id] ? (
                          <span className="text-xs text-green-600 font-medium">✓ บันทึกแล้ว</span>
                        ) : (
                          <button
                            onClick={() => saveRow(inv.id)}
                            disabled={saving[inv.id] || !isDirty(inv)}
                            className={`text-xs px-3 py-1 rounded-lg transition-colors ${
                              isDirty(inv)
                                ? 'bg-blue-600 text-white hover:bg-blue-700'
                                : 'bg-gray-100 text-gray-400 cursor-default'
                            }`}
                          >
                            {saving[inv.id] ? 'บันทึก...' : 'บันทึก'}
                          </button>
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
          <div className="mt-6 grid grid-cols-3 gap-4">
            {STATUSES.map(s => {
              const count = invoices.filter(inv => ((inv.status as Status) || 'อยู่ที่จีน') === s).length
              return (
                <div key={s} className={`rounded-xl p-4 border ${STATUS_STYLE[s].replace('text-', 'border-').split(' ')[2]} bg-white`}>
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

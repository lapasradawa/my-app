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
}

export default function DashboardPage() {
  const [invoices, setInvoices] = useState<Invoice[]>([])
  const [localStatus, setLocalStatus] = useState<Record<string, Status>>({})
  const [saving, setSaving] = useState<Record<string, boolean>>({})
  const [saved, setSaved] = useState<Record<string, boolean>>({})
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    loadInvoices()
  }, [])

  async function loadInvoices() {
    setLoading(true)
    const { data } = await supabase
      .from('invoices')
      .select('id, invoice_no, filename, created_at, status')
      .order('created_at', { ascending: false })
    if (data) {
      setInvoices(data as Invoice[])
      const initial: Record<string, Status> = {}
      for (const inv of data) {
        initial[inv.id] = (inv.status as Status) || 'อยู่ที่จีน'
      }
      setLocalStatus(initial)
    }
    setLoading(false)
  }

  async function saveStatus(id: string) {
    setSaving(s => ({ ...s, [id]: true }))
    await supabase.from('invoices').update({ status: localStatus[id] }).eq('id', id)
    setInvoices(prev => prev.map(inv => inv.id === id ? { ...inv, status: localStatus[id] } : inv))
    setSaving(s => ({ ...s, [id]: false }))
    setSaved(s => ({ ...s, [id]: true }))
    setTimeout(() => setSaved(s => ({ ...s, [id]: false })), 2000)
  }

  const formatDate = (iso: string) => new Date(iso).toLocaleDateString('th-TH', {
    day: '2-digit', month: 'short', year: 'numeric'
  })

  const statusOf = (inv: Invoice): Status => localStatus[inv.id] || 'อยู่ที่จีน'
  const isDirty = (inv: Invoice) => localStatus[inv.id] !== (inv.status || 'อยู่ที่จีน')

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Nav */}
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-blue-600 font-semibold">Dashboard</Link>
      </nav>

      <div className="max-w-5xl mx-auto px-6 py-8">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">Dashboard</h1>
          <p className="text-sm text-gray-500 mt-1">ติดตามสถานะ Invoice ทั้งหมด</p>
        </div>

        {loading ? (
          <div className="text-center py-12 text-gray-400">กำลังโหลด...</div>
        ) : invoices.length === 0 ? (
          <div className="text-center py-12 text-gray-400">
            ยังไม่มี Invoice — ไปที่{' '}
            <Link href="/" className="text-blue-600 underline">PO Matching</Link>{' '}
            เพื่ออัปโหลดและบันทึก
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
                  <th className="px-4 py-3 text-left font-medium"></th>
                </tr>
              </thead>
              <tbody>
                {invoices.map((inv, i) => (
                  <tr key={inv.id} className={`border-b border-gray-100 last:border-0 hover:bg-gray-50 transition-colors ${i % 2 === 0 ? '' : 'bg-gray-50/50'}`}>
                    <td className="px-4 py-3">
                      <Link
                        href={`/dashboard/${inv.id}`}
                        className="font-semibold text-blue-600 hover:text-blue-800 hover:underline"
                      >
                        {inv.invoice_no || '-'}
                      </Link>
                    </td>
                    <td className="px-4 py-3 text-gray-500 text-xs">{inv.filename || '-'}</td>
                    <td className="px-4 py-3 text-gray-500 text-xs">{formatDate(inv.created_at)}</td>
                    <td className="px-4 py-3">
                      <select
                        value={statusOf(inv)}
                        onChange={e => setLocalStatus(s => ({ ...s, [inv.id]: e.target.value as Status }))}
                        className={`text-xs font-medium px-2 py-1 rounded-full border cursor-pointer outline-none ${STATUS_STYLE[statusOf(inv)]}`}
                      >
                        {STATUSES.map(s => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </td>
                    <td className="px-4 py-3">
                      {saved[inv.id] ? (
                        <span className="text-xs text-green-600 font-medium">✓ บันทึกแล้ว</span>
                      ) : (
                        <button
                          onClick={() => saveStatus(inv.id)}
                          disabled={saving[inv.id] || !isDirty(inv)}
                          className={`text-xs px-3 py-1 rounded-lg transition-colors ${
                            isDirty(inv)
                              ? 'bg-blue-600 text-white hover:bg-blue-700'
                              : 'bg-gray-100 text-gray-400 cursor-default'
                          }`}
                        >
                          {saving[inv.id] ? 'กำลังบันทึก...' : 'บันทึก'}
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  )
}

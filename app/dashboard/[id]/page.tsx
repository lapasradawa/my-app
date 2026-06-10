'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { useParams } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import type { ResultRow } from '@/lib/excel-parser'

const STATUS_STYLE: Record<string, string> = {
  'อยู่ที่จีน': 'bg-yellow-100 text-yellow-800',
  'On board': 'bg-blue-100 text-blue-800',
  'ถึงคลัง': 'bg-green-100 text-green-800',
}

export default function InvoiceDetailPage() {
  const { id } = useParams<{ id: string }>()
  const [invoice, setInvoice] = useState<{
    invoice_no: string; filename: string; created_at: string
    status: string; rows: ResultRow[]; container_names: string[]
  } | null>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    async function load() {
      const { data } = await supabase.from('invoices').select('*').eq('id', id).single()
      if (data) setInvoice(data as typeof invoice)
      setLoading(false)
    }
    load()
  }, [id])

  const fmt = (n: number) => n === 0 ? '-' : n.toLocaleString()
  const formatDate = (iso: string) => new Date(iso).toLocaleDateString('th-TH', {
    day: '2-digit', month: 'short', year: 'numeric'
  })

  if (loading) return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center text-gray-400">
      กำลังโหลด...
    </div>
  )

  if (!invoice) return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center text-gray-400">
      ไม่พบ Invoice นี้
    </div>
  )

  const status = invoice.status || 'อยู่ที่จีน'

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Nav */}
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <span className="text-gray-300">|</span>
        <span className="text-sm text-blue-600 font-semibold">{invoice.invoice_no}</span>
      </nav>

      <div className="max-w-screen-xl mx-auto px-6 py-8">
        {/* Header */}
        <div className="mb-6 flex items-start justify-between flex-wrap gap-4">
          <div>
            <div className="flex items-center gap-3 mb-1">
              <h1 className="text-2xl font-bold text-gray-900">{invoice.invoice_no}</h1>
              <span className={`text-xs font-semibold px-2.5 py-1 rounded-full ${STATUS_STYLE[status] || 'bg-gray-100 text-gray-700'}`}>
                {status}
              </span>
            </div>
            <p className="text-sm text-gray-500">{invoice.filename} · บันทึกเมื่อ {formatDate(invoice.created_at)}</p>
            <p className="text-sm text-gray-500 mt-0.5">
              {invoice.rows.length} รายการ · {invoice.container_names.length} ตู้ ({invoice.container_names.join(', ')})
            </p>
          </div>
          <Link
            href="/dashboard"
            className="text-sm text-gray-500 hover:text-gray-800 border border-gray-200 px-3 py-1.5 rounded-lg transition-colors"
          >
            ← กลับ Dashboard
          </Link>
        </div>

        {/* Table */}
        <div className="overflow-auto rounded-xl border border-gray-200 bg-white shadow-sm">
          <table className="text-sm w-full border-collapse">
            <thead>
              <tr className="bg-gray-100 text-gray-700">
                <th className="px-3 py-2 text-center border-b border-gray-200 whitespace-nowrap">No.</th>
                <th className="px-3 py-2 text-left border-b border-gray-200 whitespace-nowrap min-w-[180px]">Code</th>
                <th className="px-3 py-2 text-left border-b border-gray-200 min-w-[280px]">Description</th>
                <th className="px-3 py-2 text-left border-b border-gray-200 whitespace-nowrap">PO</th>
                <th className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">QTY</th>
                {invoice.container_names.map(name => (
                  <th key={name} className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">{name}</th>
                ))}
                <th className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">LEFT</th>
              </tr>
            </thead>
            <tbody>
              {invoice.rows.map((row: ResultRow) => (
                <tr key={row.no} className="hover:bg-gray-50 border-b border-gray-100 last:border-0">
                  <td className="px-3 py-2 text-center text-gray-500">{row.no}</td>
                  <td className="px-3 py-2 font-mono text-xs">{row.code}</td>
                  <td className="px-3 py-2 text-gray-700">{row.description}</td>
                  <td className="px-3 py-2 text-gray-600 whitespace-nowrap">{row.po}</td>
                  <td className="px-3 py-2 text-right font-medium">{row.qty.toLocaleString()}</td>
                  {invoice.container_names.map(name => (
                    <td key={name} className="px-3 py-2 text-right text-gray-500">
                      {fmt(row.containers[name] || 0)}
                    </td>
                  ))}
                  <td className={`px-3 py-2 text-right font-medium ${row.left > 0 ? 'text-orange-600' : 'text-gray-400'}`}>
                    {fmt(row.left)}
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="bg-gray-50 font-semibold text-gray-700 border-t-2 border-gray-200">
                <td colSpan={4} className="px-3 py-2 text-right">รวม</td>
                <td className="px-3 py-2 text-right">{invoice.rows.reduce((s, r) => s + r.qty, 0).toLocaleString()}</td>
                {invoice.container_names.map(name => (
                  <td key={name} className="px-3 py-2 text-right">
                    {invoice.rows.reduce((s, r) => s + (r.containers[name] || 0), 0).toLocaleString()}
                  </td>
                ))}
                <td className="px-3 py-2 text-right">{invoice.rows.reduce((s, r) => s + r.left, 0).toLocaleString()}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>
  )
}

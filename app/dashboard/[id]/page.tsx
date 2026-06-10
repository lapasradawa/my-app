'use client'

import { useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import { useParams } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import type { ResultRow } from '@/lib/excel-parser'
import { exportToExcel } from '@/lib/excel-exporter'
import { isUnlocked, tryUnlock } from '@/lib/auth'

const STATUS_STYLE: Record<string, string> = {
  'อยู่ที่จีน': 'bg-yellow-100 text-yellow-800',
  'On board': 'bg-blue-100 text-blue-800',
  'ถึงไทย กำลังเข้าคลัง': 'bg-orange-100 text-orange-800',
  'ถึงคลัง': 'bg-orange-100 text-orange-800',
  'เข้าคลังแล้ว': 'bg-green-100 text-green-800',
}

interface Invoice {
  id: string
  invoice_no: string
  filename: string
  created_at: string
  status: string
  rows: ResultRow[]
  container_names: string[]
  total_amount: number | null
  currency: string | null
  bl_date: string | null
  payment_status: string | null
  payment_date: string | null
  payment_proof_url: string | null
}

function computeDueDate(blDate: string | null): Date | null {
  if (!blDate) return null
  const d = new Date(blDate)
  d.setDate(d.getDate() + 30)
  return d
}

function getPaymentLabel(invoice: Invoice): { label: string; colorClass: string } {
  if (invoice.payment_status === 'paid') {
    return { label: 'จ่ายแล้ว', colorClass: 'bg-green-100 text-green-800' }
  }
  const due = computeDueDate(invoice.bl_date)
  if (!due) return { label: 'ยังไม่จ่าย', colorClass: 'bg-gray-100 text-gray-600' }

  const today = new Date(); today.setHours(0, 0, 0, 0)
  const dueDay = new Date(due); dueDay.setHours(0, 0, 0, 0)
  const diff = Math.ceil((dueDay.getTime() - today.getTime()) / 86400000)

  if (diff < 0) return { label: 'Overdue', colorClass: 'bg-red-100 text-red-700' }
  if (diff <= 3) return { label: 'ใกล้ถึง due date', colorClass: 'bg-yellow-100 text-yellow-800' }
  return { label: 'ยังไม่จ่าย', colorClass: 'bg-gray-100 text-gray-600' }
}

function fmtDate(iso: string | null) {
  if (!iso) return '-'
  return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
}

// Password modal inline
function PasswordGate({ onSuccess, onCancel }: { onSuccess: () => void; onCancel: () => void }) {
  const [pw, setPw] = useState('')
  const [err, setErr] = useState(false)
  function attempt() {
    if (tryUnlock(pw)) { onSuccess() }
    else { setErr(true); setPw('') }
  }
  return (
    <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center">
      <div className="bg-white rounded-2xl shadow-xl p-6 w-80">
        <h3 className="font-semibold text-gray-800 mb-3">ใส่รหัสผ่านเพื่อแก้ไข</h3>
        <input
          autoFocus type="password" value={pw}
          onChange={e => { setPw(e.target.value); setErr(false) }}
          onKeyDown={e => e.key === 'Enter' && attempt()}
          placeholder="รหัสผ่าน"
          className={`w-full border rounded-lg px-3 py-2 text-sm outline-none mb-2 ${err ? 'border-red-400' : 'border-gray-300 focus:border-blue-400'}`}
        />
        {err && <p className="text-red-500 text-xs mb-2">รหัสผ่านไม่ถูกต้อง</p>}
        <div className="flex gap-2 justify-end">
          <button onClick={onCancel} className="px-3 py-1.5 text-sm text-gray-600 hover:bg-gray-100 rounded-lg">ยกเลิก</button>
          <button onClick={attempt} className="px-4 py-1.5 text-sm bg-blue-600 text-white rounded-lg hover:bg-blue-700">ยืนยัน</button>
        </div>
      </div>
    </div>
  )
}

export default function InvoiceDetailPage() {
  const { id } = useParams<{ id: string }>()
  const [invoice, setInvoice] = useState<Invoice | null>(null)
  const [loading, setLoading] = useState(true)
  const [unlocked, setUnlocked] = useState(false)

  // Password gate
  const [showPW, setShowPW] = useState(false)
  const [pwCallback, setPwCallback] = useState<(() => void) | null>(null)

  // Total amount edit
  const [editTotal, setEditTotal] = useState(false)
  const [totalInput, setTotalInput] = useState('')
  const [currencyInput, setCurrencyInput] = useState('')
  const [savingTotal, setSavingTotal] = useState(false)

  // B/L date edit
  const [editBL, setEditBL] = useState(false)
  const [blInput, setBlInput] = useState('')
  const [savingBL, setSavingBL] = useState(false)

  // Payment modal
  const [showPayModal, setShowPayModal] = useState(false)
  const [payDate, setPayDate] = useState('')
  const [payFile, setPayFile] = useState<File | null>(null)
  const [uploadingPay, setUploadingPay] = useState(false)
  const payFileRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    setUnlocked(isUnlocked())
    load()
  }, [id])

  async function load() {
    const { data } = await supabase.from('invoices').select('*').eq('id', id).single()
    if (data) setInvoice(data as Invoice)
    setLoading(false)
  }

  function requireUnlock(action: () => void) {
    if (isUnlocked()) { action() }
    else { setPwCallback(() => action); setShowPW(true) }
  }

  // Save total amount
  async function saveTotalAmount() {
    const num = parseFloat(totalInput.replace(/,/g, ''))
    if (isNaN(num) || !invoice) return
    setSavingTotal(true)
    await supabase.from('invoices').update({ total_amount: num, currency: currencyInput || null }).eq('id', id)
    setInvoice(prev => prev ? { ...prev, total_amount: num, currency: currencyInput || null } : prev)
    setEditTotal(false)
    setSavingTotal(false)
  }

  // Save B/L date
  async function saveBLDate() {
    if (!blInput || !invoice) return
    setSavingBL(true)
    await supabase.from('invoices').update({ bl_date: blInput }).eq('id', id)
    setInvoice(prev => prev ? { ...prev, bl_date: blInput } : prev)
    setEditBL(false)
    setSavingBL(false)
  }

  // Save payment
  async function savePayment() {
    if (!payDate || !invoice) return
    setUploadingPay(true)
    try {
      let proofUrl: string | null = null

      if (payFile) {
        const ext = payFile.name.split('.').pop()
        const path = `${id}/proof-${Date.now()}.${ext}`
        const { error: uploadErr } = await supabase.storage
          .from('payment-proofs')
          .upload(path, payFile, { upsert: true })
        if (!uploadErr) {
          const { data: urlData } = supabase.storage.from('payment-proofs').getPublicUrl(path)
          proofUrl = urlData.publicUrl
        }
      }

      await supabase.from('invoices').update({
        payment_status: 'paid',
        payment_date: payDate,
        payment_proof_url: proofUrl,
      }).eq('id', id)

      setInvoice(prev => prev ? {
        ...prev,
        payment_status: 'paid',
        payment_date: payDate,
        payment_proof_url: proofUrl,
      } : prev)
      setShowPayModal(false)
      setPayDate('')
      setPayFile(null)
    } finally {
      setUploadingPay(false)
    }
  }

  const fmt = (n: number) => n === 0 ? '-' : n.toLocaleString()

  if (loading) return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center text-gray-400">กำลังโหลด...</div>
  )
  if (!invoice) return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center text-gray-400">ไม่พบ Invoice นี้</div>
  )

  const status = invoice.status || 'อยู่ที่จีน'
  const dueDate = computeDueDate(invoice.bl_date)
  const payLabel = getPaymentLabel(invoice)

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

      {/* Password gate */}
      {showPW && (
        <PasswordGate
          onSuccess={() => {
            setUnlocked(true)
            setShowPW(false)
            pwCallback?.()
            setPwCallback(null)
          }}
          onCancel={() => { setShowPW(false); setPwCallback(null) }}
        />
      )}

      {/* Payment modal */}
      {showPayModal && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center">
          <div className="bg-white rounded-2xl shadow-xl p-6 w-96">
            <h3 className="font-semibold text-gray-800 mb-4">บันทึกการจ่ายเงิน</h3>
            <label className="block text-sm text-gray-600 mb-1">วันที่จ่าย</label>
            <input
              type="date" value={payDate}
              onChange={e => setPayDate(e.target.value)}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm mb-3 outline-none focus:border-blue-400"
            />
            <label className="block text-sm text-gray-600 mb-1">แนบหลักฐานการจ่าย (PDF/รูปภาพ)</label>
            <div
              className="border-2 border-dashed border-gray-200 rounded-lg p-4 text-center cursor-pointer hover:border-blue-300 mb-4"
              onClick={() => payFileRef.current?.click()}
            >
              {payFile ? (
                <p className="text-sm text-blue-600">{payFile.name}</p>
              ) : (
                <p className="text-sm text-gray-400">คลิกเพื่อเลือกไฟล์</p>
              )}
              <input
                ref={payFileRef} type="file"
                accept=".pdf,.jpg,.jpeg,.png"
                className="hidden"
                onChange={e => setPayFile(e.target.files?.[0] || null)}
              />
            </div>
            <div className="flex gap-2 justify-between">
              <button
                onClick={async () => {
                  if (!confirm('ล้างข้อมูลการจ่ายเงินทั้งหมด?')) return
                  await supabase.from('invoices').update({
                    payment_status: 'unpaid',
                    payment_date: null,
                    payment_proof_url: null,
                  }).eq('id', id)
                  setInvoice(prev => prev ? { ...prev, payment_status: 'unpaid', payment_date: null, payment_proof_url: null } : prev)
                  setShowPayModal(false); setPayDate(''); setPayFile(null)
                }}
                className="px-3 py-1.5 text-sm text-red-500 hover:bg-red-50 rounded-lg"
              >
                ล้างข้อมูล
              </button>
              <div className="flex gap-2">
                <button
                  onClick={() => { setShowPayModal(false); setPayDate(''); setPayFile(null) }}
                  className="px-3 py-1.5 text-sm text-gray-600 hover:bg-gray-100 rounded-lg"
                >
                  ยกเลิก
                </button>
                <button
                  onClick={savePayment}
                  disabled={!payDate || uploadingPay}
                  className="px-4 py-1.5 text-sm bg-green-600 text-white rounded-lg hover:bg-green-700 disabled:opacity-50"
                >
                  {uploadingPay ? 'กำลังบันทึก...' : 'บันทึก'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-screen-xl mx-auto px-6 py-8">
        {/* Header */}
        <div className="mb-6 flex items-center justify-between flex-wrap gap-4">
          <div className="flex items-center gap-3">
            <h1 className="text-2xl font-bold text-gray-900">{invoice.invoice_no}</h1>
            <span className={`text-xs font-semibold px-2.5 py-1 rounded-full ${STATUS_STYLE[status] || 'bg-gray-100 text-gray-700'}`}>
              {status}
            </span>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => exportToExcel(invoice.rows, invoice.container_names, invoice.invoice_no + '-PO-Matching.xlsx')}
              className="px-4 py-1.5 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 transition-colors"
            >
              Export Excel
            </button>
            <Link
              href="/dashboard"
              className="text-sm text-gray-500 hover:text-gray-800 border border-gray-200 px-3 py-1.5 rounded-lg transition-colors"
            >
              ← กลับ Dashboard
            </Link>
          </div>
        </div>

        {/* Info cards */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          {/* Total amount */}
          <div className="bg-white rounded-xl border border-gray-200 p-4">
            <p className="text-xs text-gray-400 mb-1">ยอดรวม Invoice</p>
            {editTotal ? (
              <div className="flex flex-col gap-1">
                <input
                  autoFocus type="text" value={totalInput}
                  onChange={e => setTotalInput(e.target.value)}
                  placeholder="เช่น 424527.01"
                  className="border border-gray-300 rounded px-2 py-1 text-sm outline-none focus:border-blue-400 w-full"
                />
                <input
                  type="text" value={currencyInput}
                  onChange={e => setCurrencyInput(e.target.value.toUpperCase())}
                  placeholder="สกุลเงิน เช่น CNY"
                  className="border border-gray-300 rounded px-2 py-1 text-sm outline-none focus:border-blue-400 w-full"
                />
                <div className="flex gap-1 mt-1">
                  <button
                    onClick={saveTotalAmount}
                    disabled={savingTotal || !totalInput}
                    className="text-xs px-2 py-1 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:opacity-50"
                  >
                    บันทึก
                  </button>
                  <button onClick={() => setEditTotal(false)} className="text-xs px-2 py-1 text-gray-500 hover:bg-gray-100 rounded">
                    ยกเลิก
                  </button>
                </div>
              </div>
            ) : (
              <>
                {invoice.total_amount ? (
                  <>
                    <p className="text-lg font-bold text-gray-900">
                      {invoice.total_amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </p>
                    <p className="text-xs text-gray-500">{invoice.currency || ''}</p>
                  </>
                ) : (
                  <p className="text-sm text-gray-400">-</p>
                )}
                <button
                  onClick={() => requireUnlock(() => {
                    setTotalInput(invoice.total_amount ? String(invoice.total_amount) : '')
                    setCurrencyInput(invoice.currency || '')
                    setEditTotal(true)
                  })}
                  className="text-xs text-blue-500 hover:text-blue-700 mt-1"
                >
                  {invoice.total_amount ? 'แก้ไข' : '+ ใส่ยอดเงิน'}
                </button>
              </>
            )}
          </div>

          {/* B/L date */}
          <div className="bg-white rounded-xl border border-gray-200 p-4">
            <p className="text-xs text-gray-400 mb-1">B/L Date</p>
            {editBL ? (
              <div className="flex flex-col gap-1">
                <input
                  autoFocus type="date" value={blInput}
                  onChange={e => setBlInput(e.target.value)}
                  className="border border-gray-300 rounded px-2 py-1 text-sm outline-none focus:border-blue-400 w-full"
                />
                <div className="flex gap-1">
                  <button
                    onClick={saveBLDate}
                    disabled={savingBL || !blInput}
                    className="text-xs px-2 py-1 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:opacity-50"
                  >
                    บันทึก
                  </button>
                  <button
                    onClick={() => setEditBL(false)}
                    className="text-xs px-2 py-1 text-gray-500 hover:bg-gray-100 rounded"
                  >
                    ยกเลิก
                  </button>
                </div>
              </div>
            ) : (
              <>
                <p className="text-sm font-semibold text-gray-900">{fmtDate(invoice.bl_date)}</p>
                <button
                  onClick={() => requireUnlock(() => { setBlInput(invoice.bl_date || ''); setEditBL(true) })}
                  className="text-xs text-blue-500 hover:text-blue-700 mt-1"
                >
                  {invoice.bl_date ? 'แก้ไข' : '+ เพิ่ม B/L date'}
                </button>
              </>
            )}
          </div>

          {/* Due date */}
          <div className="bg-white rounded-xl border border-gray-200 p-4">
            <p className="text-xs text-gray-400 mb-1">Due Date (B/L +30 วัน)</p>
            <p className="text-sm font-semibold text-gray-900">
              {dueDate ? dueDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '-'}
            </p>
            {invoice.bl_date && (
              <p className="text-xs text-gray-400 mt-1">
                {(() => {
                  const today = new Date(); today.setHours(0, 0, 0, 0)
                  const d = new Date(dueDate!); d.setHours(0, 0, 0, 0)
                  const diff = Math.ceil((d.getTime() - today.getTime()) / 86400000)
                  if (diff < 0) return <span className="text-red-500">เกิน {Math.abs(diff)} วัน</span>
                  if (diff === 0) return <span className="text-orange-500">ครบกำหนดวันนี้</span>
                  return `อีก ${diff} วัน`
                })()}
              </p>
            )}
          </div>

          {/* Payment status */}
          <div className="bg-white rounded-xl border border-gray-200 p-4">
            <p className="text-xs text-gray-400 mb-1">สถานะการจ่ายเงิน</p>
            <span className={`inline-block text-xs font-semibold px-2.5 py-1 rounded-full mb-2 ${payLabel.colorClass}`}>
              {payLabel.label}
            </span>
            {invoice.payment_status === 'paid' ? (
              <div className="space-y-0.5">
                <p className="text-xs text-gray-500">จ่ายวันที่ {fmtDate(invoice.payment_date)}</p>
                {invoice.payment_proof_url && (
                  <a
                    href={invoice.payment_proof_url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-xs text-blue-500 hover:underline block"
                  >
                    ดูหลักฐาน →
                  </a>
                )}
                <button
                  onClick={() => requireUnlock(() => {
                    setPayDate(invoice.payment_date || '')
                    setShowPayModal(true)
                  })}
                  className="text-xs text-gray-400 hover:text-gray-600 mt-1 block"
                >
                  แก้ไข
                </button>
              </div>
            ) : (
              <button
                onClick={() => requireUnlock(() => {
                  setPayDate('')
                  setPayFile(null)
                  setShowPayModal(true)
                })}
                className="text-xs px-3 py-1.5 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors block"
              >
                Submit Payment
              </button>
            )}
          </div>
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

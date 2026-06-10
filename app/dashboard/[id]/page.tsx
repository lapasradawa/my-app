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
  baikon_url: string | null
  bl_doc_url: string | null
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
  const [payFiles, setPayFiles] = useState<File[]>([])
  const [uploadingPay, setUploadingPay] = useState(false)
  const payFileRef = useRef<HTMLInputElement>(null)

  // Document uploads
  const [uploadingBaikon, setUploadingBaikon] = useState(false)
  const [uploadingBlDoc, setUploadingBlDoc] = useState(false)
  const baikonRef = useRef<HTMLInputElement>(null)
  const blDocRef = useRef<HTMLInputElement>(null)

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

  function parseProofUrls(raw: string | null): string[] {
    if (!raw) return []
    try { const arr = JSON.parse(raw); if (Array.isArray(arr)) return arr } catch {}
    return [raw]
  }

  // Save payment — uploads all selected files and merges with existing proof URLs
  async function savePayment() {
    if (!payDate || !invoice) return
    setUploadingPay(true)
    try {
      const existing = parseProofUrls(invoice.payment_proof_url)
      const newUrls: string[] = []

      for (const file of payFiles) {
        const ext = file.name.split('.').pop()
        const path = `${id}/proof-${Date.now()}-${Math.random().toString(36).slice(2)}.${ext}`
        const { error: uploadErr } = await supabase.storage
          .from('payment-proofs')
          .upload(path, file, { upsert: true })
        if (!uploadErr) {
          const { data: urlData } = supabase.storage.from('payment-proofs').getPublicUrl(path)
          newUrls.push(urlData.publicUrl)
        }
      }

      const allUrls = [...existing, ...newUrls]
      const proofValue = allUrls.length > 0 ? JSON.stringify(allUrls) : null

      await supabase.from('invoices').update({
        payment_status: 'paid',
        payment_date: payDate,
        payment_proof_url: proofValue,
      }).eq('id', id)

      setInvoice(prev => prev ? {
        ...prev,
        payment_status: 'paid',
        payment_date: payDate,
        payment_proof_url: proofValue,
      } : prev)
      setShowPayModal(false)
      setPayDate('')
      setPayFiles([])
    } finally {
      setUploadingPay(false)
    }
  }

  async function uploadDoc(file: File, type: 'baikon' | 'bl') {
    const setUploading = type === 'baikon' ? setUploadingBaikon : setUploadingBlDoc
    setUploading(true)
    try {
      const ext = file.name.split('.').pop()
      const path = `${id}/${type}-${Date.now()}.${ext}`
      const { error } = await supabase.storage.from('payment-proofs').upload(path, file, { upsert: true })
      if (error) return
      const { data: urlData } = supabase.storage.from('payment-proofs').getPublicUrl(path)
      const url = urlData.publicUrl
      const field = type === 'baikon' ? 'baikon_url' : 'bl_doc_url'
      await supabase.from('invoices').update({ [field]: url }).eq('id', id)
      setInvoice(prev => prev ? { ...prev, [field]: url } : prev)
    } finally {
      setUploading(false)
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

  // Fixed column widths — consistent across all invoices, matching YG260051 proportions
  const W = { no: 50, code: 200, desc: 280, po: 250, qty: 80, cntr: 120, left: 80 }
  const L = {
    code: W.no,
    desc: W.no + W.code,
    po:   W.no + W.code + W.desc,
    qty:  W.no + W.code + W.desc + W.po,
  }
  const tableWidth = W.no + W.code + W.desc + W.po + W.qty + invoice.container_names.length * W.cntr + W.left

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
            {payFiles.length > 0 && (
              <ul className="mb-2 space-y-1">
                {payFiles.map((f, i) => (
                  <li key={i} className="flex items-center justify-between text-xs bg-blue-50 rounded px-2 py-1">
                    <span className="text-blue-700 truncate">{f.name}</span>
                    <button
                      type="button"
                      onClick={() => setPayFiles(prev => prev.filter((_, j) => j !== i))}
                      className="ml-2 text-gray-400 hover:text-red-500"
                    >✕</button>
                  </li>
                ))}
              </ul>
            )}
            <div
              className="border-2 border-dashed border-gray-200 rounded-lg p-3 text-center cursor-pointer hover:border-blue-300 mb-4"
              onClick={() => payFileRef.current?.click()}
            >
              <p className="text-sm text-gray-400">+ คลิกเพื่อเพิ่มไฟล์</p>
              <input
                ref={payFileRef} type="file"
                accept=".pdf,.jpg,.jpeg,.png"
                multiple
                className="hidden"
                onChange={e => {
                  const files = Array.from(e.target.files || [])
                  setPayFiles(prev => [...prev, ...files])
                  e.target.value = ''
                }}
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
                  setShowPayModal(false); setPayDate(''); setPayFiles([])
                }}
                className="px-3 py-1.5 text-sm text-red-500 hover:bg-red-50 rounded-lg"
              >
                ล้างข้อมูล
              </button>
              <div className="flex gap-2">
                <button
                  onClick={() => { setShowPayModal(false); setPayDate(''); setPayFiles([]) }}
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
                  const d = new Date(dueDate!); d.setHours(0, 0, 0, 0)
                  if (invoice.payment_status === 'paid' && invoice.payment_date) {
                    const paid = new Date(invoice.payment_date); paid.setHours(0, 0, 0, 0)
                    const late = Math.ceil((paid.getTime() - d.getTime()) / 86400000)
                    if (late > 0) return <span className="text-red-500">เกิน {late} วัน</span>
                    return null
                  }
                  const today = new Date(); today.setHours(0, 0, 0, 0)
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
                {parseProofUrls(invoice.payment_proof_url).map((url, i, arr) => (
                  <a
                    key={i}
                    href={url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-xs text-blue-500 hover:underline block"
                  >
                    ดูหลักฐาน{arr.length > 1 ? ` ${i + 1}` : ''} →
                  </a>
                ))}
                <button
                  onClick={() => requireUnlock(() => {
                    setPayDate(invoice.payment_date || '')
                    setPayFiles([])
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
                  setPayFiles([])
                  setShowPayModal(true)
                })}
                className="text-xs px-3 py-1.5 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors block"
              >
                Submit Payment
              </button>
            )}
          </div>
        </div>

        {/* Documents */}
        <div className="grid grid-cols-2 gap-4 mb-6">
          {/* ใบขน */}
          <div className="bg-white rounded-xl border border-gray-200 p-4">
            <p className="text-xs text-gray-400 mb-2">เอกสารใบขน</p>
            {invoice.baikon_url ? (
              <div className="flex flex-col gap-2">
                <a
                  href={invoice.baikon_url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1.5 text-sm text-blue-600 hover:text-blue-800 font-medium"
                >
                  <span>📄</span> ดู / ดาวน์โหลด ใบขน
                </a>
                <button
                  onClick={() => requireUnlock(() => baikonRef.current?.click())}
                  className="text-xs text-gray-400 hover:text-gray-600 self-start"
                >
                  {uploadingBaikon ? 'กำลังอัปโหลด...' : 'เปลี่ยนไฟล์'}
                </button>
              </div>
            ) : (
              <button
                onClick={() => requireUnlock(() => baikonRef.current?.click())}
                disabled={uploadingBaikon}
                className="flex items-center gap-2 border-2 border-dashed border-gray-200 rounded-lg px-4 py-3 text-sm text-gray-400 hover:border-blue-300 hover:text-blue-500 w-full justify-center transition-colors disabled:opacity-50"
              >
                {uploadingBaikon ? 'กำลังอัปโหลด...' : '+ อัปโหลดใบขน (PDF)'}
              </button>
            )}
            <input
              ref={baikonRef} type="file" accept=".pdf" className="hidden"
              onChange={e => { const f = e.target.files?.[0]; if (f) uploadDoc(f, 'baikon') }}
            />
          </div>

          {/* B/L Document */}
          <div className="bg-white rounded-xl border border-gray-200 p-4">
            <p className="text-xs text-gray-400 mb-2">เอกสาร B/L</p>
            {invoice.bl_doc_url ? (
              <div className="flex flex-col gap-2">
                <a
                  href={invoice.bl_doc_url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1.5 text-sm text-blue-600 hover:text-blue-800 font-medium"
                >
                  <span>📄</span> ดู / ดาวน์โหลด B/L
                </a>
                <button
                  onClick={() => requireUnlock(() => blDocRef.current?.click())}
                  className="text-xs text-gray-400 hover:text-gray-600 self-start"
                >
                  {uploadingBlDoc ? 'กำลังอัปโหลด...' : 'เปลี่ยนไฟล์'}
                </button>
              </div>
            ) : (
              <button
                onClick={() => requireUnlock(() => blDocRef.current?.click())}
                disabled={uploadingBlDoc}
                className="flex items-center gap-2 border-2 border-dashed border-gray-200 rounded-lg px-4 py-3 text-sm text-gray-400 hover:border-blue-300 hover:text-blue-500 w-full justify-center transition-colors disabled:opacity-50"
              >
                {uploadingBlDoc ? 'กำลังอัปโหลด...' : '+ อัปโหลด B/L (PDF)'}
              </button>
            )}
            <input
              ref={blDocRef} type="file" accept=".pdf" className="hidden"
              onChange={e => { const f = e.target.files?.[0]; if (f) uploadDoc(f, 'bl') }}
            />
          </div>
        </div>

        {/* Table — fixed col widths, sticky header/footer, sticky left cols up to QTY */}
        <div className="overflow-auto rounded-xl border border-gray-200 bg-white shadow-sm max-h-[calc(100vh-380px)]">
          <table className="text-sm border-collapse table-fixed" style={{ width: tableWidth }}>
            <colgroup>
              <col style={{ width: W.no }} />
              <col style={{ width: W.code }} />
              <col style={{ width: W.desc }} />
              <col style={{ width: W.po }} />
              <col style={{ width: W.qty }} />
              {invoice.container_names.map(name => <col key={name} style={{ width: W.cntr }} />)}
              <col style={{ width: W.left }} />
            </colgroup>
            <thead>
              <tr className="bg-gray-100 text-gray-700 text-left">
                <th className="px-3 py-2 text-center border-b border-gray-200 sticky top-0 z-30 bg-gray-100" style={{ left: 0 }}>No.</th>
                <th className="px-3 py-2 border-b border-gray-200 sticky top-0 z-30 bg-gray-100 whitespace-nowrap" style={{ left: L.code }}>Code</th>
                <th className="px-3 py-2 border-b border-gray-200 sticky top-0 z-30 bg-gray-100" style={{ left: L.desc }}>Description</th>
                <th className="px-3 py-2 border-b border-gray-200 sticky top-0 z-30 bg-gray-100" style={{ left: L.po }}>PO</th>
                <th className="px-3 py-2 text-right border-b border-gray-200 sticky top-0 z-30 bg-gray-100 shadow-[2px_0_5px_rgba(0,0,0,0.07)]" style={{ left: L.qty }}>QTY</th>
                {invoice.container_names.map(name => (
                  <th key={name} className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap sticky top-0 z-20 bg-gray-100">{name}</th>
                ))}
                <th className="px-3 py-2 text-right border-b border-gray-200 sticky top-0 z-20 bg-gray-100">LEFT</th>
              </tr>
            </thead>
            <tbody>
              {invoice.rows.map((row: ResultRow) => (
                <tr key={row.no} className="group border-b border-gray-100 last:border-0">
                  <td className="px-3 py-2 text-center text-gray-500 sticky z-10 bg-white group-hover:bg-gray-50" style={{ left: 0 }}>{row.no}</td>
                  <td className="px-3 py-2 font-mono text-xs sticky z-10 bg-white group-hover:bg-gray-50" style={{ left: L.code }}>{row.code}</td>
                  <td className="px-3 py-2 text-gray-700 sticky z-10 bg-white group-hover:bg-gray-50" style={{ left: L.desc }}>{row.description}</td>
                  <td className="px-3 py-2 text-gray-600 sticky z-10 bg-white group-hover:bg-gray-50" style={{ left: L.po }}>{row.po}</td>
                  <td className="px-3 py-2 text-right font-medium sticky z-10 bg-white group-hover:bg-gray-50 shadow-[2px_0_5px_rgba(0,0,0,0.07)]" style={{ left: L.qty }}>{row.qty.toLocaleString()}</td>
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
              <tr className="bg-gray-50 font-semibold text-gray-700 border-t-2 border-gray-200 sticky bottom-0 z-20">
                <td className="px-3 py-2 sticky z-20 bg-gray-50" style={{ left: 0 }}></td>
                <td className="px-3 py-2 sticky z-20 bg-gray-50" style={{ left: L.code }}></td>
                <td className="px-3 py-2 sticky z-20 bg-gray-50" style={{ left: L.desc }}></td>
                <td className="px-3 py-2 text-right sticky z-20 bg-gray-50" style={{ left: L.po }}>รวม</td>
                <td className="px-3 py-2 text-right sticky z-20 bg-gray-50 shadow-[2px_0_5px_rgba(0,0,0,0.07)]" style={{ left: L.qty }}>{invoice.rows.reduce((s, r) => s + r.qty, 0).toLocaleString()}</td>
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

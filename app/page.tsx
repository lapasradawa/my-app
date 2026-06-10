'use client'

import { useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import { processExcel, type ResultRow, type ProcessResult } from '@/lib/excel-parser'
import { exportToExcel } from '@/lib/excel-exporter'
import { supabase } from '@/lib/supabase'
import { isUnlocked } from '@/lib/auth'
import LockButton from '@/components/LockButton'
import PasswordModal from '@/components/PasswordModal'

interface HistoryItem {
  id: string
  invoice_no: string
  filename: string
  created_at: string
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [result, setResult] = useState<ProcessResult | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)
  const [filename, setFilename] = useState('')
  const [dragging, setDragging] = useState(false)
  const [history, setHistory] = useState<HistoryItem[]>([])
  const [saving, setSaving] = useState(false)
  const [savedId, setSavedId] = useState<string | null>(null)
  const [invoiceName, setInvoiceName] = useState('')
  const [editingName, setEditingName] = useState(false)
  const [unlocked, setUnlocked] = useState(false)
  const [showPasswordModal, setShowPasswordModal] = useState(false)
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null)

  useEffect(() => {
    setUnlocked(isUnlocked())
    loadHistory()
  }, [])

  async function loadHistory() {
    const { data } = await supabase
      .from('invoices')
      .select('id, invoice_no, filename, created_at')
      .order('created_at', { ascending: false })
      .limit(50)
    if (data) setHistory(data as HistoryItem[])
  }

  async function handleFile(file: File) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      setError('กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls) เท่านั้น')
      return
    }
    setLoading(true)
    setError(null)
    setResult(null)
    setSavedId(null)
    setFilename(file.name)
    try {
      const buffer = await file.arrayBuffer()
      const data = processExcel(buffer)
      setResult(data)
      setInvoiceName(data.invoiceNo || file.name.replace(/\.xlsx?$/, ''))
    } catch (e) {
      setError(e instanceof Error ? e.message : 'เกิดข้อผิดพลาด')
    } finally {
      setLoading(false)
    }
  }

  function requireUnlock(action: () => void) {
    if (isUnlocked()) {
      action()
    } else {
      setPendingAction(() => action)
      setShowPasswordModal(true)
    }
  }

  async function handleSave() {
    if (!result) return
    setSaving(true)
    try {
      if (savedId) {
        // Update existing record (rename)
        await supabase.from('invoices').update({ invoice_no: invoiceName }).eq('id', savedId)
      } else {
        // Insert new record
        const { data, error: err } = await supabase.from('invoices').insert({
          invoice_no: invoiceName,
          filename,
          rows: result.rows,
          container_names: result.containerNames,
        }).select('id').single()
        if (err) throw new Error(err.message)
        if (data) setSavedId(data.id)
      }
      await loadHistory()
      setEditingName(false)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'บันทึกไม่สำเร็จ')
    } finally {
      setSaving(false)
    }
  }

  async function loadFromHistory(item: HistoryItem) {
    const { data } = await supabase.from('invoices').select('*').eq('id', item.id).single()
    if (data) {
      setResult({ rows: data.rows, containerNames: data.container_names, invoiceNo: data.invoice_no })
      setFilename(data.filename)
      setInvoiceName(data.invoice_no)
      setSavedId(data.id)
      setEditingName(false)
    }
  }

  async function deleteHistory(id: string, e: React.MouseEvent) {
    e.stopPropagation()
    await supabase.from('invoices').delete().eq('id', id)
    setHistory(prev => prev.filter(h => h.id !== id))
    if (savedId === id) reset()
  }

  function onFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (file) handleFile(file)
  }

  function onDrop(e: React.DragEvent) {
    e.preventDefault()
    setDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (file) handleFile(file)
  }

  function onDragOver(e: React.DragEvent) { e.preventDefault(); setDragging(true) }
  function onDragLeave() { setDragging(false) }

  function handleExport() {
    if (!result) return
    exportToExcel(result.rows, result.containerNames, invoiceName + '-PO-Matching.xlsx')
  }

  function reset() {
    setResult(null)
    setError(null)
    setFilename('')
    setSavedId(null)
    setInvoiceName('')
    setEditingName(false)
    if (inputRef.current) inputRef.current.value = ''
  }

  const fmt = (n: number) => n === 0 ? '-' : n.toLocaleString()
  const formatDate = (iso: string) => new Date(iso).toLocaleDateString('th-TH', {
    day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit'
  })

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Nav */}
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4 shrink-0">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-blue-600 font-semibold">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
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
      <div className="flex flex-1 overflow-hidden">
      {/* Sidebar */}
      <aside className="w-64 shrink-0 bg-white border-r border-gray-200 flex flex-col">
        <div className="p-4 border-b border-gray-200">
          <h2 className="font-semibold text-gray-800 text-sm">ประวัติ Invoice</h2>
        </div>
        <div className="flex-1 overflow-y-auto">
          {history.length === 0 ? (
            <p className="text-xs text-gray-400 p-4">ยังไม่มีประวัติ</p>
          ) : (
            history.map(item => (
              <button
                key={item.id}
                onClick={() => loadFromHistory(item)}
                className={`w-full text-left px-4 py-3 border-b border-gray-100 group transition-colors ${
                  savedId === item.id ? 'bg-blue-50' : 'hover:bg-gray-50'
                }`}
              >
                <div className="flex items-start justify-between gap-1">
                  <div className="min-w-0">
                    <p className="text-sm font-medium text-gray-800 truncate">{item.invoice_no || '-'}</p>
                    <p className="text-xs text-gray-400 truncate mt-0.5">{item.filename}</p>
                    <p className="text-xs text-gray-400 mt-0.5">{formatDate(item.created_at)}</p>
                  </div>
                  <span
                    onClick={(e) => deleteHistory(item.id, e)}
                    className="text-gray-300 hover:text-red-400 text-xs mt-0.5 shrink-0 opacity-0 group-hover:opacity-100 transition-opacity cursor-pointer"
                  >
                    ✕
                  </span>
                </div>
              </button>
            ))
          )}
        </div>
        <div className="p-4 border-t border-gray-200">
          <button
            onClick={reset}
            className="w-full py-2 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700 transition-colors"
          >
            + อัปโหลดไฟล์ใหม่
          </button>
        </div>
      </aside>

      {/* Main */}
      <div className="flex-1 p-6 overflow-auto">
        <div className="max-w-screen-xl mx-auto">
          <div className="mb-6">
            <h1 className="text-2xl font-bold text-gray-900">CI / Packing List — PO Matching</h1>
            <p className="text-sm text-gray-500 mt-1">อัปโหลดไฟล์ Excel เพื่อ match สินค้าในแต่ละตู้กับ PO</p>
          </div>

          {!result && !loading && (
            <div
              className={`border-2 border-dashed rounded-xl p-12 text-center cursor-pointer transition-colors ${
                dragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 bg-white hover:border-blue-400'
              }`}
              onClick={() => inputRef.current?.click()}
              onDrop={onDrop}
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
            >
              <div className="text-4xl mb-3">📂</div>
              <p className="text-gray-700 font-medium">ลากไฟล์มาวางที่นี่ หรือคลิกเพื่อเลือกไฟล์</p>
              <p className="text-gray-400 text-sm mt-1">รองรับไฟล์ .xlsx และ .xls</p>
              <p className="text-gray-400 text-xs mt-2">ไฟล์ต้องมีชีท CI และชีทตู้ (เช่น UETU2248127)</p>
              <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={onFileChange} />
            </div>
          )}

          {loading && <div className="text-center py-12 text-gray-500">กำลังประมวลผล...</div>}

          {error && (
            <div className="mt-4 bg-red-50 border border-red-200 text-red-700 rounded-lg px-4 py-3 text-sm">
              {error}
              <button onClick={() => setError(null)} className="ml-3 underline">ปิด</button>
            </div>
          )}

          {result && (
            <>
              {/* Invoice name bar */}
              <div className="bg-white border border-gray-200 rounded-xl px-4 py-3 mb-4 flex items-center gap-3 flex-wrap">
                <div className="flex items-center gap-2 flex-1 min-w-0">
                  <span className="text-sm text-gray-500 shrink-0">Invoice:</span>
                  {editingName ? (
                    <input
                      autoFocus
                      value={invoiceName}
                      onChange={e => setInvoiceName(e.target.value)}
                      onKeyDown={e => e.key === 'Enter' && handleSave()}
                      className="border border-blue-400 rounded px-2 py-1 text-sm flex-1 outline-none"
                    />
                  ) : (
                    <span
                      className="text-sm font-semibold text-gray-800 cursor-pointer hover:text-blue-600 truncate"
                      onClick={() => setEditingName(true)}
                      title="คลิกเพื่อแก้ไขชื่อ"
                    >
                      {invoiceName || '-'}
                    </span>
                  )}
                  <button
                    onClick={() => setEditingName(v => !v)}
                    className="text-xs text-gray-400 hover:text-blue-500 shrink-0"
                  >
                    {editingName ? 'ยกเลิก' : 'แก้ไข'}
                  </button>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <span className="text-sm text-gray-400">{result.rows.length} รายการ | {result.containerNames.length} ตู้</span>
                  <button
                    onClick={() => requireUnlock(handleSave)}
                    disabled={saving}
                    className="px-4 py-1.5 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700 disabled:opacity-50 transition-colors"
                  >
                    {saving ? 'กำลังบันทึก...' : savedId ? 'บันทึก (อัปเดต)' : 'บันทึก'}
                  </button>
                  <button
                    onClick={handleExport}
                    className="px-4 py-1.5 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 transition-colors"
                  >
                    Export Excel
                  </button>
                </div>
              </div>

              {/* Table */}
              <div className="overflow-auto rounded-xl border border-gray-200 bg-white shadow-sm">
                <table className="text-sm w-full border-collapse">
                  <thead>
                    <tr className="bg-gray-100 text-gray-700">
                      <th className="px-3 py-2 text-center border-b border-gray-200 whitespace-nowrap">No.</th>
                      <th className="px-3 py-2 text-left border-b border-gray-200 whitespace-nowrap min-w-[180px]">Code</th>
                      <th className="px-3 py-2 text-left border-b border-gray-200 min-w-[300px]">Description</th>
                      <th className="px-3 py-2 text-left border-b border-gray-200 whitespace-nowrap min-w-[200px]">PO</th>
                      <th className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">QTY</th>
                      {result.containerNames.map(name => (
                        <th key={name} className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">{name}</th>
                      ))}
                      {result.containerNames.length > 0 && (
                        <th className="px-3 py-2 text-right border-b border-gray-200 whitespace-nowrap">LEFT</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {result.rows.map((row: ResultRow) => (
                      <tr key={row.no} className="hover:bg-gray-50 border-b border-gray-100 last:border-0">
                        <td className="px-3 py-2 text-center text-gray-500">{row.no}</td>
                        <td className="px-3 py-2 font-mono text-xs">{row.code}</td>
                        <td className="px-3 py-2 text-gray-700">{row.description}</td>
                        <td className="px-3 py-2 text-gray-600 whitespace-nowrap">{row.po}</td>
                        <td className="px-3 py-2 text-right font-medium">{row.qty.toLocaleString()}</td>
                        {result.containerNames.map(name => (
                          <td key={name} className="px-3 py-2 text-right text-gray-500">
                            {fmt(row.containers[name])}
                          </td>
                        ))}
                        {result.containerNames.length > 0 && (
                          <td className={`px-3 py-2 text-right font-medium ${row.left > 0 ? 'text-orange-600' : 'text-gray-400'}`}>
                            {fmt(row.left)}
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-gray-50 font-semibold text-gray-700 border-t-2 border-gray-200">
                      <td colSpan={4} className="px-3 py-2 text-right">รวม</td>
                      <td className="px-3 py-2 text-right">{result.rows.reduce((s, r) => s + r.qty, 0).toLocaleString()}</td>
                      {result.containerNames.map(name => (
                        <td key={name} className="px-3 py-2 text-right">
                          {result.rows.reduce((s, r) => s + (r.containers[name] || 0), 0).toLocaleString()}
                        </td>
                      ))}
                      {result.containerNames.length > 0 && (
                        <td className="px-3 py-2 text-right">{result.rows.reduce((s, r) => s + r.left, 0).toLocaleString()}</td>
                      )}
                    </tr>
                  </tfoot>
                </table>
              </div>
            </>
          )}
        </div>
      </div>
      </div>
    </div>
  )
}

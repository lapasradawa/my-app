'use client'

import { useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import { isUnlocked } from '@/lib/auth'
import LockButton from '@/components/LockButton'
import NavBar from '@/components/NavBar'
import { parsePOUploadExcel } from '@/lib/po-upload-parser'

interface PoRecord {
  id: string
  supplier: string
  project: string
  currency: string
  filename: string | null
  po_date: string | null
  po_rbs_ch_no: string | null
  po_rbs_th_no: string | null
  total_amount: number
  cost_saving: number | null
  cost_saving_pct: number | null
  created_at: string
  rows: { item_code: string; description: string; qty: number; unit_price: number; total: number }[]
}

interface PoBuilderItem {
  item_code: string
  description: string
  fob_price: number | null
  currency: string
  qty: string
}

export default function PoInsightsPage() {
  const fileRef = useRef<HTMLInputElement>(null)
  const [, setUnlocked] = useState(false)

  // ── PO list ──────────────────────────────────────────────────────────
  const [records, setRecords] = useState<PoRecord[]>([])
  const [loadingList, setLoadingList] = useState(true)

  // ── Upload form ───────────────────────────────────────────────────────
  const [showUpload, setShowUpload] = useState(false)
  const [uploadFile, setUploadFile] = useState<File | null>(null)
  const [uploadParsed, setUploadParsed] = useState<ReturnType<typeof parsePOUploadExcel> | null>(null)
  const [uploadError, setUploadError] = useState('')
  const [uploading, setUploading] = useState(false)
  const [uploadSupplier, setUploadSupplier] = useState('')
  const [uploadProject, setUploadProject] = useState('')
  const [uploadDate, setUploadDate] = useState('')
  const [uploadChNo, setUploadChNo] = useState('')
  const [dragging, setDragging] = useState(false)

  // ── Exchange rates (from cost_settings) ─────────────────────────────
  const [cnyRate, setCnyRate] = useState(4.85)
  const [usdRate, setUsdRate] = useState(33.00)

  // ── PO Builder (generate) ─────────────────────────────────────────────
  const [showBuilder, setShowBuilder] = useState(false)
  const [allSuppliers, setAllSuppliers] = useState<string[]>([])
  const [allProjects, setAllProjects] = useState<string[]>([])
  const [poSupplier, setPoSupplier] = useState('')
  const [poItemInput, setPoItemInput] = useState('')
  const [poItems, setPoItems] = useState<PoBuilderItem[]>([])
  const [poAdding, setPoAdding] = useState(false)

  useEffect(() => {
    setUnlocked(isUnlocked())
    loadList()
    loadMeta()
  }, [])

  async function loadList() {
    setLoadingList(true)
    const { data } = await supabase
      .from('po_uploads')
      .select('id, supplier, project, currency, filename, po_date, po_rbs_ch_no, po_rbs_th_no, total_amount, cost_saving, cost_saving_pct, created_at')
      .order('created_at', { ascending: false })
    setRecords((data ?? []) as PoRecord[])
    setLoadingList(false)
  }

  async function loadMeta() {
    const [{ data: items }, { data: settings }] = await Promise.all([
      supabase.from('po_items').select('supplier, project'),
      supabase.from('cost_settings').select('key, value'),
    ])
    if (items) {
      const rows = items as { supplier: string; project: string }[]
      setAllSuppliers([...new Set(rows.map(r => r.supplier))].sort())
      setAllProjects([...new Set(rows.map(r => r.project))].sort())
    }
    if (settings) {
      const m = Object.fromEntries((settings as { key: string; value: string }[]).map(r => [r.key, r.value]))
      if (m.cny_rate) setCnyRate(parseFloat(m.cny_rate))
      if (m.usd_rate) setUsdRate(parseFloat(m.usd_rate))
    }
  }

  // ── Upload handlers ───────────────────────────────────────────────────
  function handleFileSelect(file: File) {
    setUploadError('')
    setUploadParsed(null)
    setUploadFile(file)
    file.arrayBuffer().then(buf => {
      try {
        const result = parsePOUploadExcel(buf)
        setUploadParsed(result)
      } catch (e) {
        setUploadError(e instanceof Error ? e.message : 'parse error')
      }
    })
  }

  async function handleSavePO() {
    if (!uploadParsed || !uploadSupplier || !uploadProject) return
    setUploading(true)
    setUploadError('')
    try {
      // Insert into po_uploads
      const { data: inserted, error } = await supabase
        .from('po_uploads')
        .insert({
          supplier: uploadSupplier,
          project: uploadProject,
          currency: uploadParsed.currency,
          filename: uploadFile?.name ?? null,
          po_date: uploadDate || null,
          po_rbs_ch_no: uploadChNo.trim() || null,
          po_rbs_th_no: null,
          rows: uploadParsed.rows,
          total_amount: uploadParsed.total_amount,
        })
        .select('id')
        .single()
      if (error) throw new Error(error.message)

      // Update po_items (delete old, insert new prices) for Cost Compare
      await supabase
        .from('po_items')
        .delete()
        .eq('project', uploadProject)
        .eq('supplier', uploadSupplier)

      const poItemRows = uploadParsed.rows.map(r => ({
        project: uploadProject,
        supplier: uploadSupplier,
        item_code: r.item_code,
        description: r.description || null,
        fob_price: r.unit_price,
        currency: uploadParsed.currency,
        file_name: uploadFile?.name ?? null,
        uploaded_at: new Date().toISOString(),
      }))
      if (poItemRows.length > 0) {
        await supabase.from('po_items').insert(poItemRows)
      }

      // Reset form
      setUploadFile(null)
      setUploadParsed(null)
      setUploadSupplier('')
      setUploadProject('')
      setUploadDate('')
      setUploadChNo('')
      setShowUpload(false)
      await loadList()
    } catch (e) {
      setUploadError(e instanceof Error ? e.message : 'บันทึกไม่สำเร็จ')
    } finally {
      setUploading(false)
    }
  }

  // ── PO Builder (generate) handlers ───────────────────────────────────
  async function addPoItem() {
    const code = poItemInput.trim()
    if (!code || !poSupplier) return
    setPoAdding(true)
    const { data } = await supabase
      .from('po_items')
      .select('item_code, description, fob_price, currency')
      .ilike('item_code', code)
      .eq('supplier', poSupplier)
      .order('uploaded_at', { ascending: false })
      .limit(1)
    const found = data?.[0] as { item_code: string; description: string | null; fob_price: number; currency: string } | undefined
    setPoItems(prev => [...prev, {
      item_code: found?.item_code ?? code,
      description: found?.description ?? '',
      fob_price: found?.fob_price ?? null,
      currency: found?.currency ?? 'CNY',
      qty: '',
    }])
    setPoItemInput('')
    setPoAdding(false)
  }

  async function changePoSupplier(newSupplier: string) {
    setPoSupplier(newSupplier)
    if (!newSupplier || poItems.length === 0) return
    const updated = await Promise.all(poItems.map(async item => {
      const { data } = await supabase
        .from('po_items').select('fob_price, currency')
        .ilike('item_code', item.item_code)
        .eq('supplier', newSupplier)
        .order('uploaded_at', { ascending: false }).limit(1)
      const found = data?.[0] as { fob_price: number; currency: string } | undefined
      return { ...item, fob_price: found?.fob_price ?? null, currency: found?.currency ?? item.currency }
    }))
    setPoItems(updated)
  }

  function exportPoExcel() {
    const currency = poItems.find(i => i.fob_price !== null)?.currency ?? 'CNY'
    const sheetRows: (string | number)[][] = []
    sheetRows.push([], [], [])
    sheetRows.push(['No.', 'Item Code', 'Description', 'QTY', `UNIT PRICE (${currency}/PC)`, `TOTAL (${currency})`])
    for (let i = 0; i < 30; i++) {
      if (i < poItems.length) {
        const item = poItems[i]
        const qty = parseFloat(item.qty) || 0
        const unit = item.fob_price ?? 0
        sheetRows.push([i + 1, item.item_code, item.description, qty || '', unit || '', qty > 0 && unit > 0 ? qty * unit : ''])
      } else {
        sheetRows.push([i + 1, '', '', '', '', ''])
      }
    }
    const totalQty = poItems.reduce((s, i) => s + (parseFloat(i.qty) || 0), 0)
    const totalAmt = poItems.reduce((s, i) => s + (parseFloat(i.qty) || 0) * (i.fob_price ?? 0), 0)
    sheetRows.push(['TOTAL', '', '', totalQty || '', '', totalAmt || ''])
    sheetRows.push([])
    sheetRows.push(['Remark:'])
    sheetRows.push(['1  Please specify above Purchase Order number in every invoice.'])
    sheetRows.push(['2  Please refer to shipment plan in excel file'])
    sheetRows.push([])
    sheetRows.push(['Issuer', '', '', '', 'Approved by'])
    sheetRows.push([])
    sheetRows.push(['Lapasrada Wanish', '', '', '', 'Piraya Lueprasitsakul'])
    sheetRows.push(['Purchasing and Import Coordinator', '', '', '', 'Purchasing Manager'])
    const today = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
    sheetRows.push([today, '', '', '', today])
    const ws = XLSX.utils.aoa_to_sheet(sheetRows)
    ws['!cols'] = [{ wch: 5 }, { wch: 27 }, { wch: 60 }, { wch: 10 }, { wch: 15 }, { wch: 15 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'PO')
    XLSX.writeFile(wb, `PO_${poSupplier}_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  const fmt = (n: number, d = 2) => n.toLocaleString(undefined, { minimumFractionDigits: d, maximumFractionDigits: d })
  const fmtDate = (iso: string) => new Date(iso).toLocaleDateString('th-TH', { day: '2-digit', month: 'short', year: 'numeric' })

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      <NavBar onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />

      <div className="max-w-6xl mx-auto w-full px-6 py-8">
        {/* Header */}
        <div className="flex items-center justify-between flex-wrap gap-3 mb-6">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">PO Insights</h1>
            <p className="text-sm text-gray-500 mt-0.5">บันทึกและติดตาม Purchase Order ที่เปิดกับ Supplier</p>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => { setShowUpload(v => !v); setShowBuilder(false) }}
              className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700 transition-colors"
            >
              ↑ อัปโหลด PO
            </button>
            <button
              onClick={() => { setShowBuilder(v => !v); setShowUpload(false) }}
              className="flex items-center gap-1.5 px-4 py-2 bg-white border border-gray-300 text-gray-700 text-sm rounded-lg hover:bg-gray-50 transition-colors"
            >
              + สร้าง PO
            </button>
          </div>
        </div>

        {/* ── Upload form ──────────────────────────────────────────────── */}
        {showUpload && (
          <div className="bg-white rounded-xl border border-blue-200 shadow-sm p-6 mb-6">
            <h2 className="text-sm font-bold text-gray-800 mb-4">อัปโหลด PO Excel เข้าระบบ</h2>

            {/* File drop */}
            <div
              className={`border-2 border-dashed rounded-lg p-6 text-center mb-4 cursor-pointer transition-colors ${
                dragging ? 'border-blue-400 bg-blue-50' : uploadParsed ? 'border-green-400 bg-green-50' : 'border-gray-300 hover:border-blue-400'
              }`}
              onClick={() => fileRef.current?.click()}
              onDrop={e => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) handleFileSelect(f) }}
              onDragOver={e => { e.preventDefault(); setDragging(true) }}
              onDragLeave={() => setDragging(false)}
            >
              {uploadParsed ? (
                <div>
                  <p className="text-green-700 font-medium text-sm">✓ {uploadFile?.name}</p>
                  <p className="text-green-600 text-xs mt-1">{uploadParsed.rows.length} รายการ · สกุลเงิน {uploadParsed.currency} · รวม {fmt(uploadParsed.total_amount)} {uploadParsed.currency}</p>
                </div>
              ) : (
                <div>
                  <p className="text-gray-500 text-sm">ลากไฟล์มาวาง หรือคลิกเพื่อเลือกไฟล์ Excel</p>
                  <p className="text-gray-400 text-xs mt-1">คอลัมน์ต้องมี Item Code, Description, QTY, Unit Price</p>
                </div>
              )}
              <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e => { const f = e.target.files?.[0]; if (f) handleFileSelect(f) }} />
            </div>

            {uploadError && <p className="text-red-500 text-xs mb-3">{uploadError}</p>}

            {/* Fields */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Supplier *</label>
                <select
                  value={uploadSupplier}
                  onChange={e => setUploadSupplier(e.target.value)}
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm outline-none focus:border-blue-400"
                >
                  <option value="">— เลือก —</option>
                  {allSuppliers.map(s => <option key={s}>{s}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Project *</label>
                <select
                  value={uploadProject}
                  onChange={e => setUploadProject(e.target.value)}
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm outline-none focus:border-blue-400"
                >
                  <option value="">— เลือก —</option>
                  {allProjects.map(p => <option key={p}>{p}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">วันที่เปิด PO</label>
                <input
                  type="date"
                  value={uploadDate}
                  onChange={e => setUploadDate(e.target.value)}
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm outline-none focus:border-blue-400"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Currency</label>
                <div className="border border-gray-200 rounded-lg px-2 py-1.5 text-sm text-gray-600 bg-gray-50">
                  {uploadParsed?.currency ?? '—'} (auto)
                </div>
              </div>
            </div>
            <div className="grid grid-cols-2 gap-3 mb-4">
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">PO RBS CH No.</label>
                <input
                  type="text"
                  value={uploadChNo}
                  onChange={e => setUploadChNo(e.target.value)}
                  placeholder="เช่น RBSYG01-GEN5"
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm outline-none focus:border-blue-400"
                />
              </div>
              <div className="flex items-end">
                <p className="text-xs text-gray-400 pb-2">PO RBS TH No. เพิ่มได้ในหน้า detail หลังจาก upload แล้ว</p>
              </div>
            </div>

            <div className="flex items-center gap-2 justify-end">
              <button onClick={() => { setShowUpload(false); setUploadParsed(null); setUploadFile(null) }} className="px-4 py-2 text-sm text-gray-500 hover:text-gray-700">ยกเลิก</button>
              <button
                onClick={handleSavePO}
                disabled={!uploadParsed || !uploadSupplier || !uploadProject || uploading}
                className="px-5 py-2 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700 disabled:opacity-40 transition-colors"
              >
                {uploading ? 'กำลังบันทึก...' : 'บันทึก PO + อัปเดต Cost Compare'}
              </button>
            </div>
          </div>
        )}

        {/* ── PO Builder (generate) ────────────────────────────────────── */}
        {showBuilder && (
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mb-6">
            <h2 className="text-sm font-bold text-gray-800 mb-4">สร้าง PO Excel</h2>
            <div className="flex items-center gap-3 mb-4 flex-wrap">
              <label className="text-sm text-gray-600 font-medium whitespace-nowrap">Supplier:</label>
              <select value={poSupplier} onChange={e => changePoSupplier(e.target.value)}
                className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 bg-white min-w-[180px]">
                <option value="">— เลือก Supplier —</option>
                {allSuppliers.map(s => <option key={s}>{s}</option>)}
              </select>
            </div>
            {poSupplier && (
              <>
                <div className="flex items-center gap-2 mb-4">
                  <input type="text" value={poItemInput} onChange={e => setPoItemInput(e.target.value)}
                    onKeyDown={e => e.key === 'Enter' && addPoItem()}
                    placeholder="กรอก Item Code แล้วกด Enter..."
                    className="border border-gray-300 rounded-lg px-3 py-2 text-sm outline-none focus:border-blue-400 flex-1 max-w-sm font-mono" />
                  <button onClick={addPoItem} disabled={!poItemInput.trim() || poAdding}
                    className="px-4 py-2 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700 disabled:opacity-40 transition-colors">
                    {poAdding ? '...' : '+ เพิ่ม'}
                  </button>
                  {poItems.length > 0 && (
                    <button onClick={exportPoExcel}
                      className="ml-auto px-4 py-2 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 transition-colors">
                      ↓ Export Excel
                    </button>
                  )}
                </div>
                {poItems.length > 0 && (
                  <table className="w-full text-sm border-collapse">
                    <thead>
                      <tr className="bg-gray-50 text-gray-500 text-xs border-b border-gray-200">
                        <th className="px-3 py-2 text-left w-10">No.</th>
                        <th className="px-3 py-2 text-left">Item Code</th>
                        <th className="px-3 py-2 text-left">Description</th>
                        <th className="px-3 py-2 text-right">FOB Price</th>
                        <th className="px-3 py-2 text-right w-28">QTY</th>
                        <th className="px-3 py-2 text-right w-32">Total</th>
                        <th className="px-3 py-2 w-8"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {poItems.map((item, i) => {
                        const qty = parseFloat(item.qty) || 0
                        const total = qty * (item.fob_price ?? 0)
                        return (
                          <tr key={i} className="border-b border-gray-100 hover:bg-gray-50/50">
                            <td className="px-3 py-2 text-gray-400 text-xs">{i + 1}</td>
                            <td className="px-3 py-2 font-mono text-gray-800 text-xs">{item.item_code}</td>
                            <td className="px-3 py-2 text-gray-600 text-xs">{item.description || <span className="text-gray-400">—</span>}</td>
                            <td className="px-3 py-2 text-right text-gray-700 text-xs">
                              {item.fob_price !== null ? `${fmt(item.fob_price)} ${item.currency}` : <span className="text-gray-400">ไม่พบ</span>}
                            </td>
                            <td className="px-3 py-2 text-right">
                              <input type="number" min="0" value={item.qty}
                                onChange={e => setPoItems(prev => prev.map((p, j) => j === i ? { ...p, qty: e.target.value } : p))}
                                className="border border-gray-200 rounded px-2 py-1 w-24 text-right text-xs outline-none focus:border-blue-400" placeholder="0" />
                            </td>
                            <td className="px-3 py-2 text-right text-xs">
                              {qty > 0 && item.fob_price !== null ? <span className="text-gray-700">{fmt(total)} {item.currency}</span> : <span className="text-gray-400">—</span>}
                            </td>
                            <td className="px-3 py-2 text-center">
                              <button onClick={() => setPoItems(prev => prev.filter((_, j) => j !== i))} className="text-gray-300 hover:text-red-400 text-xs">✕</button>
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                )}
              </>
            )}
          </div>
        )}

        {/* ── PO list ──────────────────────────────────────────────────── */}
        {loadingList ? (
          <p className="text-sm text-gray-400">กำลังโหลด...</p>
        ) : records.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <p className="text-4xl mb-3">📄</p>
            <p className="text-sm">ยังไม่มี PO ในระบบ — กด "อัปโหลด PO" เพื่อเริ่มต้น</p>
          </div>
        ) : (
          <div className="overflow-x-auto rounded-xl border border-gray-200 bg-white shadow-sm">
            <table className="w-full text-sm border-collapse">
              <thead>
                <tr className="bg-gray-50 text-gray-500 text-xs border-b border-gray-200">
                  <th className="px-4 py-3 text-left">Supplier</th>
                  <th className="px-4 py-3 text-left">Project</th>
                  <th className="px-4 py-3 text-left">PO RBS CH No.</th>
                  <th className="px-4 py-3 text-left">วันที่เปิด PO</th>
                  <th className="px-4 py-3 text-right">Original currency FOB</th>
                  <th className="px-4 py-3 text-right">FOB (THB)</th>
                  <th className="px-4 py-3 text-right">Cost Saving (THB)</th>
                  <th className="px-4 py-3 text-left">อัปโหลดเมื่อ</th>
                  <th className="px-4 py-3"></th>
                </tr>
              </thead>
              <tbody>
                {records.map(rec => (
                  <tr key={rec.id} className="border-b border-gray-100 hover:bg-blue-50/30 transition-colors">
                    <td className="px-4 py-3 font-medium text-gray-900">{rec.supplier}</td>
                    <td className="px-4 py-3 text-gray-600">{rec.project}</td>
                    <td className="px-4 py-3 text-gray-600 text-xs font-mono">{rec.po_rbs_ch_no ?? <span className="text-gray-400">—</span>}</td>
                    <td className="px-4 py-3 text-gray-600">
                      {rec.po_date ? new Date(rec.po_date).toLocaleDateString('th-TH', { day: '2-digit', month: 'short', year: 'numeric' }) : <span className="text-gray-400">—</span>}
                    </td>
                    <td className="px-4 py-3 text-right font-medium text-gray-800">
                      {fmt(rec.total_amount)} <span className="text-gray-400 text-xs">{rec.currency}</span>
                    </td>
                    <td className="px-4 py-3 text-right font-medium text-gray-800">
                      {fmt(rec.total_amount * (rec.currency === 'USD' ? usdRate : cnyRate))}
                      <span className="text-gray-400 text-xs ml-1">THB</span>
                    </td>
                    <td className="px-4 py-3 text-right text-gray-700 text-xs">
                      {rec.cost_saving != null ? fmt(rec.cost_saving) : <span className="text-gray-400">—</span>}
                    </td>
                    <td className="px-4 py-3 text-gray-400 text-xs">{fmtDate(rec.created_at)}</td>
                    <td className="px-4 py-3 text-right">
                      <Link href={`/po-builder/${rec.id}`} className="text-blue-600 hover:text-blue-800 text-xs font-medium whitespace-nowrap">
                        ดูรายละเอียด →
                      </Link>
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

'use client'

import { Fragment, useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import { parsePO } from '@/lib/po-parser'
import { isUnlocked } from '@/lib/auth'
import LockButton from '@/components/LockButton'
import PasswordModal from '@/components/PasswordModal'

interface POItemDB {
  id: string
  project: string
  supplier: string
  item_code: string
  description: string | null
  fob_price: number
  currency: string
  document_no: string | null
  document_date: string | null
  file_name: string | null
  uploaded_at: string
}

interface Settings {
  cny_rate: number
  usd_rate: number
  ddp_multiplier: number
}

interface LatestPrice {
  fob_price: number
  currency: string
  uploaded_at: string
}

interface TableRow {
  item_code: string
  description: string
  prices: Record<string, LatestPrice>
}

export default function ComparePage() {
  const [settings, setSettings] = useState<Settings>({ cny_rate: 4.85, usd_rate: 33.00, ddp_multiplier: 1.11 })
  const [editSettings, setEditSettings] = useState(false)
  const [settingsInput, setSettingsInput] = useState({ cny_rate: '', usd_rate: '', ddp_multiplier: '' })
  const [savingSettings, setSavingSettings] = useState(false)

  const [projects, setProjects] = useState<string[]>([])
  const [selectedProject, setSelectedProject] = useState('')
  const [suppliers, setSuppliers] = useState<string[]>([])
  const [supplierDates, setSupplierDates] = useState<Record<string, string>>({})
  const [tableRows, setTableRows] = useState<TableRow[]>([])
  const [loadingTable, setLoadingTable] = useState(false)

  const [uploadProject, setUploadProject] = useState('')
  const [uploadSupplier, setUploadSupplier] = useState('')
  const [uploading, setUploading] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)

  const [history, setHistory] = useState<{ item_code: string; supplier: string; entries: POItemDB[] } | null>(null)
  const [replaceMode, setReplaceMode] = useState(false)
  const [search, setSearch] = useState('')

  interface UploadBatch { uploaded_at: string; file_name: string | null; count: number }
  const [managingSupplier, setManagingSupplier] = useState<string | null>(null)

  interface PoBuilderItem {
    item_code: string
    description: string
    fob_price: number | null
    currency: string
    qty: string
  }
  const [allSuppliers, setAllSuppliers] = useState<string[]>([])
  const [poSupplier, setPoSupplier] = useState('')
  const [poItemInput, setPoItemInput] = useState('')
  const [poItems, setPoItems] = useState<PoBuilderItem[]>([])
  const [poAdding, setPoAdding] = useState(false)
  const [batches, setBatches] = useState<UploadBatch[]>([])
  const [loadingBatches, setLoadingBatches] = useState(false)
  const [deletingBatch, setDeletingBatch] = useState<string | null>(null)

  const [, setUnlocked] = useState(false)
  const [showPasswordModal, setShowPasswordModal] = useState(false)
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null)

  useEffect(() => { loadSettings(); loadProjects() }, [])

  function requireUnlock(action: () => void) {
    if (isUnlocked()) action()
    else { setPendingAction(() => action); setShowPasswordModal(true) }
  }

  async function loadSettings() {
    const { data } = await supabase.from('cost_settings').select('key, value')
    if (data) {
      const m = Object.fromEntries((data as { key: string; value: string }[]).map(r => [r.key, r.value]))
      setSettings({
        cny_rate: parseFloat(m.cny_rate ?? '4.85'),
        usd_rate: parseFloat(m.usd_rate ?? '33.00'),
        ddp_multiplier: parseFloat(m.ddp_multiplier ?? '1.11'),
      })
    }
  }

  async function saveSettings() {
    setSavingSettings(true)
    const updates = [
      { key: 'cny_rate', value: settingsInput.cny_rate || String(settings.cny_rate) },
      { key: 'usd_rate', value: settingsInput.usd_rate || String(settings.usd_rate) },
      { key: 'ddp_multiplier', value: settingsInput.ddp_multiplier || String(settings.ddp_multiplier) },
    ]
    for (const u of updates) await supabase.from('cost_settings').upsert(u)
    await loadSettings()
    setEditSettings(false)
    setSavingSettings(false)
  }

  async function loadProjects() {
    const { data } = await supabase.from('po_items').select('project, supplier')
    if (data) {
      const rows = data as { project: string; supplier: string }[]
      setProjects([...new Set(rows.map(r => r.project))].sort())
      setAllSuppliers([...new Set(rows.map(r => r.supplier))].sort())
    }
  }

  async function loadProject(project: string) {
    setSelectedProject(project)
    setLoadingTable(true)
    const { data } = await supabase
      .from('po_items').select('*').eq('project', project)
      .order('uploaded_at', { ascending: false })
    if (data) {
      const seen = new Set<string>()
      const latest: POItemDB[] = []
      for (const item of data as POItemDB[]) {
        const key = `${item.supplier}__${item.item_code}`
        if (!seen.has(key)) { seen.add(key); latest.push(item) }
      }
      // Pick the latest description per item_code across all suppliers
      const descMap = new Map<string, { description: string; uploaded_at: string }>()
      for (const item of latest) {
        const cur = descMap.get(item.item_code)
        if (!cur || item.uploaded_at > cur.uploaded_at)
          descMap.set(item.item_code, { description: item.description || '', uploaded_at: item.uploaded_at })
      }

      const supplierSet = new Set<string>()
      const itemMap = new Map<string, TableRow>()
      const dates: Record<string, string> = {}
      for (const item of latest) {
        supplierSet.add(item.supplier)
        if (!itemMap.has(item.item_code)) {
          itemMap.set(item.item_code, {
            item_code: item.item_code,
            description: descMap.get(item.item_code)?.description ?? '',
            prices: {},
          })
        }
        itemMap.get(item.item_code)!.prices[item.supplier] = {
          fob_price: item.fob_price, currency: item.currency, uploaded_at: item.uploaded_at,
        }
        if (!dates[item.supplier] || item.uploaded_at > dates[item.supplier]) dates[item.supplier] = item.uploaded_at
      }
      setSuppliers([...supplierSet].sort())
      setSupplierDates(dates)
      setTableRows([...itemMap.values()].sort((a, b) => a.item_code.localeCompare(b.item_code)))
    }
    setLoadingTable(false)
  }

  async function openManageModal(supplier: string) {
    setManagingSupplier(supplier)
    setLoadingBatches(true)
    const { data } = await supabase
      .from('po_items').select('uploaded_at, file_name')
      .eq('project', selectedProject).eq('supplier', supplier)
      .order('uploaded_at', { ascending: false })
    if (data) {
      const batchMap = new Map<string, UploadBatch>()
      for (const row of data as { uploaded_at: string; file_name: string | null }[]) {
        if (!batchMap.has(row.uploaded_at)) {
          batchMap.set(row.uploaded_at, { uploaded_at: row.uploaded_at, file_name: row.file_name, count: 0 })
        }
        batchMap.get(row.uploaded_at)!.count++
      }
      setBatches([...batchMap.values()])
    }
    setLoadingBatches(false)
  }

  async function deleteBatch(supplier: string, uploaded_at: string) {
    setDeletingBatch(uploaded_at)
    const { error } = await supabase.from('po_items').delete()
      .eq('project', selectedProject).eq('supplier', supplier).eq('uploaded_at', uploaded_at)
    if (error) { alert('ลบไม่สำเร็จ: ' + error.message) }
    else {
      await openManageModal(supplier)
      await loadProject(selectedProject)
      await loadProjects()
    }
    setDeletingBatch(null)
  }

  async function handleFileUpload(file: File) {
    setUploading(true)
    const proj = uploadProject.trim()
    const supp = uploadSupplier.trim()
    try {
      const buf = await file.arrayBuffer()
      const { items } = parsePO(buf)
      if (items.length === 0) { alert('ไม่พบข้อมูล item ในไฟล์'); return }
      if (replaceMode) {
        const { error } = await supabase.from('po_items').delete().eq('project', proj).eq('supplier', supp)
        if (error) { alert('ลบข้อมูลเดิมไม่สำเร็จ: ' + error.message); return }
      }
      const rows = items.map(item => ({
        project: proj,
        supplier: supp,
        item_code: item.item_code,
        description: item.description || null,
        fob_price: item.fob_price,
        currency: item.currency,
        document_no: item.document_no || null,
        document_date: item.document_date || null,
        file_name: file.name,
      }))
      const { error } = await supabase.from('po_items').insert(rows)
      if (error) { alert('บันทึกไม่สำเร็จ: ' + error.message); return }
      alert(`บันทึก ${items.length} รายการสำเร็จ${replaceMode ? ' (แทนที่ข้อมูลเดิม)' : ''}`)
      await loadProjects()
      if (selectedProject === proj) await loadProject(proj)
      setUploadProject('')
      setUploadSupplier('')
    } catch (e) {
      alert('เกิดข้อผิดพลาด: ' + (e as Error).message)
    } finally {
      setUploading(false)
    }
  }

  async function showHistory(item_code: string, supplier: string) {
    const { data } = await supabase
      .from('po_items').select('*')
      .eq('project', selectedProject).eq('supplier', supplier).eq('item_code', item_code)
      .order('uploaded_at', { ascending: false })
    if (data) setHistory({ item_code, supplier, entries: data as POItemDB[] })
  }

  function exportExcel(rows: TableRow[]) {
    const header = ['Item Code', 'Description',
      ...suppliers.flatMap(s => {
        const currency = rows.find(r => r.prices[s])?.prices[s]?.currency ?? 'CNY'
        return [`${s} FOB ${currency}`, `${s} FOB THB`, `${s} DDP THB`]
      })
    ]
    const data = rows.map(row => [
      row.item_code, row.description,
      ...suppliers.flatMap(s => {
        const p = row.prices[s]
        if (!p) return ['', '', ''] as (string | number)[]
        const fobThb = p.fob_price * getRate(p.currency)
        return [p.fob_price, fobThb, fobThb * settings.ddp_multiplier] as (string | number)[]
      })
    ])
    const ws = XLSX.utils.aoa_to_sheet([header, ...data])
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Cost Compare')
    XLSX.writeFile(wb, `CostCompare_${selectedProject}_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

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
        .from('po_items')
        .select('fob_price, currency, description')
        .ilike('item_code', item.item_code)
        .eq('supplier', newSupplier)
        .order('uploaded_at', { ascending: false })
        .limit(1)
      const found = data?.[0] as { fob_price: number; currency: string; description: string | null } | undefined
      return { ...item, fob_price: found?.fob_price ?? null, currency: found?.currency ?? item.currency }
    }))
    setPoItems(updated)
  }

  function removePoItem(index: number) {
    setPoItems(prev => prev.filter((_, i) => i !== index))
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
    ws['!cols'] = [{ wch: 6 }, { wch: 28 }, { wch: 48 }, { wch: 10 }, { wch: 18 }, { wch: 14 }]

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'PO')
    XLSX.writeFile(wb, `PO_${poSupplier}_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  function getRate(currency: string) {
    return currency === 'USD' ? settings.usd_rate : settings.cny_rate
  }

  function fmtDate(iso: string) {
    return new Date(iso).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
  }

  function fmtN(n: number) {
    return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4">
        <span className="font-bold text-gray-800 text-sm">Import PO</span>
        <span className="text-gray-300">|</span>
        <Link href="/" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/report" className="text-sm text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/compare" className="text-sm text-blue-600 font-semibold">Cost Compare</Link>
        <div className="ml-auto">
          <LockButton onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />
        </div>
      </nav>

      {showPasswordModal && (
        <PasswordModal
          onSuccess={() => { setShowPasswordModal(false); pendingAction?.(); setPendingAction(null) }}
          onCancel={() => { setShowPasswordModal(false); setPendingAction(null) }}
        />
      )}

      <div className="max-w-screen-2xl mx-auto px-6 py-8">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">Cost Compare</h1>
          <p className="text-sm text-gray-500 mt-1">เปรียบเทียบราคา FOB และ DDP จากทุก Supplier ต่อ Project</p>
        </div>

        {/* Settings bar */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 mb-6">
          <div className="flex items-center gap-6 flex-wrap">
            <span className="text-xs font-semibold text-gray-400 uppercase tracking-wide">Estimate Rates</span>
            {editSettings ? (
              <>
                <label className="flex items-center gap-1.5 text-xs text-gray-600">
                  CNY/THB:
                  <input type="text" value={settingsInput.cny_rate}
                    onChange={e => setSettingsInput(s => ({ ...s, cny_rate: e.target.value }))}
                    className="border border-gray-300 rounded px-2 py-1 w-16 outline-none focus:border-blue-400" />
                </label>
                <label className="flex items-center gap-1.5 text-xs text-gray-600">
                  USD/THB:
                  <input type="text" value={settingsInput.usd_rate}
                    onChange={e => setSettingsInput(s => ({ ...s, usd_rate: e.target.value }))}
                    className="border border-gray-300 rounded px-2 py-1 w-16 outline-none focus:border-blue-400" />
                </label>
                <label className="flex items-center gap-1.5 text-xs text-gray-600">
                  DDP ×:
                  <input type="text" value={settingsInput.ddp_multiplier}
                    onChange={e => setSettingsInput(s => ({ ...s, ddp_multiplier: e.target.value }))}
                    className="border border-gray-300 rounded px-2 py-1 w-16 outline-none focus:border-blue-400" />
                </label>
                <button onClick={saveSettings} disabled={savingSettings}
                  className="text-xs px-3 py-1.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50">
                  {savingSettings ? 'บันทึก...' : 'บันทึก'}
                </button>
                <button onClick={() => setEditSettings(false)} className="text-xs text-gray-400 hover:text-gray-600">ยกเลิก</button>
              </>
            ) : (
              <>
                <span className="text-sm text-gray-700">CNY/THB: <strong>{settings.cny_rate}</strong></span>
                <span className="text-sm text-gray-700">USD/THB: <strong>{settings.usd_rate}</strong></span>
                <span className="text-sm text-gray-700">DDP ×: <strong>{settings.ddp_multiplier}</strong></span>
                <button
                  onClick={() => requireUnlock(() => {
                    setSettingsInput({ cny_rate: String(settings.cny_rate), usd_rate: String(settings.usd_rate), ddp_multiplier: String(settings.ddp_multiplier) })
                    setEditSettings(true)
                  })}
                  className="text-xs text-blue-500 hover:text-blue-700 ml-2">
                  แก้ไข
                </button>
              </>
            )}
          </div>
        </div>

        <div className="grid grid-cols-2 gap-6 mb-6">
          {/* Upload PO */}
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4">
            <p className="text-sm font-semibold text-gray-700 mb-3">อัปโหลด PO</p>
            <div className="space-y-2">
              <input type="text" value={uploadProject} onChange={e => setUploadProject(e.target.value)}
                placeholder="ชื่อ Project เช่น 7-11, PTT"
                className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm outline-none focus:border-blue-400" />
              <input type="text" value={uploadSupplier} onChange={e => setUploadSupplier(e.target.value)}
                placeholder="Supplier เช่น YONGGUAN, LITELON, YPN"
                className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm outline-none focus:border-blue-400" />
              <button
                onClick={() => fileRef.current?.click()}
                disabled={uploading || !uploadProject.trim() || !uploadSupplier.trim()}
                className="w-full flex items-center justify-center gap-2 border-2 border-dashed border-gray-200 rounded-lg py-3 text-sm text-gray-400 hover:border-blue-300 hover:text-blue-500 disabled:opacity-40 disabled:cursor-not-allowed transition-colors">
                {uploading ? 'กำลังบันทึก...' : '+ อัปโหลดไฟล์ PO (Excel)'}
              </button>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden"
                onChange={e => { const f = e.target.files?.[0]; if (f) handleFileUpload(f); e.target.value = '' }} />
              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input type="checkbox" checked={replaceMode} onChange={e => setReplaceMode(e.target.checked)}
                  className="w-3.5 h-3.5 accent-blue-600" />
                <span className="text-xs text-gray-500">แทนที่ข้อมูลเดิมของ Supplier นี้</span>
              </label>
              <p className="text-xs text-gray-400">รองรับเฉพาะ Excel (.xlsx, .xls)</p>
            </div>
          </div>

          {/* Project selector */}
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4">
            <p className="text-sm font-semibold text-gray-700 mb-3">เลือก Project</p>
            {projects.length === 0 ? (
              <p className="text-sm text-gray-400">ยังไม่มี Project — อัปโหลด PO ก่อน</p>
            ) : (
              <div className="flex flex-wrap gap-2">
                {projects.map(p => (
                  <button key={p} onClick={() => loadProject(p)}
                    className={`px-3 py-1.5 rounded-full text-sm font-medium border transition-colors ${
                      selectedProject === p
                        ? 'bg-blue-600 text-white border-blue-600'
                        : 'bg-white text-gray-600 border-gray-300 hover:border-blue-400 hover:text-blue-600'
                    }`}>
                    {p}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* PO Builder */}
        <div className="mb-8">
          <div className="mb-3">
            <h2 className="text-base font-semibold text-gray-800">PO Builder</h2>
            <p className="text-xs text-gray-400 mt-0.5">เลือก Supplier และเพิ่ม Item Code จากทุก Project เพื่อสร้าง PO</p>
          </div>
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-5">
              <div className="flex items-center gap-3 mb-4 flex-wrap">
                <label className="text-sm text-gray-600 font-medium whitespace-nowrap">Supplier:</label>
                <select
                  value={poSupplier}
                  onChange={e => changePoSupplier(e.target.value)}
                  className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 bg-white">
                  <option value="">— เลือก Supplier —</option>
                  {allSuppliers.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>

              {poSupplier && (
                <>
                  <div className="flex items-center gap-2 mb-4">
                    <input
                      type="text"
                      value={poItemInput}
                      onChange={e => setPoItemInput(e.target.value)}
                      onKeyDown={e => e.key === 'Enter' && addPoItem()}
                      placeholder="กรอก Item Code แล้วกด Enter..."
                      className="border border-gray-300 rounded-lg px-3 py-2 text-sm outline-none focus:border-blue-400 flex-1 max-w-sm font-mono" />
                    <button
                      onClick={addPoItem}
                      disabled={!poItemInput.trim() || poAdding}
                      className="px-4 py-2 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700 disabled:opacity-40 transition-colors">
                      {poAdding ? '...' : '+ เพิ่ม'}
                    </button>
                    {poItems.length > 0 && (
                      <button
                        onClick={exportPoExcel}
                        className="ml-auto flex items-center gap-1.5 px-4 py-2 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 transition-colors">
                        ↓ Export PO Excel
                      </button>
                    )}
                  </div>

                  {poItems.length === 0 ? (
                    <p className="text-sm text-gray-400 text-center py-6">เพิ่ม Item Code เพื่อสร้าง PO</p>
                  ) : (
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
                              <td className="px-3 py-2 font-mono text-gray-800 text-xs whitespace-nowrap">{item.item_code}</td>
                              <td className="px-3 py-2 text-gray-600 text-xs">{item.description || <span className="text-gray-300">—</span>}</td>
                              <td className="px-3 py-2 text-right text-gray-700 text-xs whitespace-nowrap">
                                {item.fob_price !== null ? `${fmtN(item.fob_price)} ${item.currency}` : <span className="text-gray-300">ไม่พบราคา</span>}
                              </td>
                              <td className="px-3 py-2 text-right">
                                <input
                                  type="number"
                                  min="0"
                                  value={item.qty}
                                  onChange={e => setPoItems(prev => prev.map((p, j) => j === i ? { ...p, qty: e.target.value } : p))}
                                  className="border border-gray-200 rounded px-2 py-1 w-24 text-right text-xs outline-none focus:border-blue-400"
                                  placeholder="0" />
                              </td>
                              <td className="px-3 py-2 text-right font-medium text-xs">
                                {qty > 0 && item.fob_price !== null
                                  ? <span className="text-gray-700">{fmtN(total)} {item.currency}</span>
                                  : <span className="text-gray-300">—</span>}
                              </td>
                              <td className="px-3 py-2 text-center">
                                <button onClick={() => removePoItem(i)} className="text-gray-300 hover:text-red-400 transition-colors text-xs">✕</button>
                              </td>
                            </tr>
                          )
                        })}
                      </tbody>
                      <tfoot>
                        <tr className="bg-gray-50 font-semibold text-xs border-t border-gray-200">
                          <td colSpan={4} className="px-3 py-2.5 text-right text-gray-500">TOTAL</td>
                          <td className="px-3 py-2.5 text-right text-gray-700">
                            {fmtN(poItems.reduce((s, i) => s + (parseFloat(i.qty) || 0), 0))}
                          </td>
                          <td className="px-3 py-2.5 text-right text-gray-700">
                            {fmtN(poItems.reduce((s, i) => s + (parseFloat(i.qty) || 0) * (i.fob_price ?? 0), 0))} {poItems.find(i => i.fob_price !== null)?.currency ?? ''}
                          </td>
                          <td></td>
                        </tr>
                      </tfoot>
                    </table>
                  )}
                </>
              )}
          </div>
        </div>

        {/* Compare table */}
        {selectedProject && (
          <>
            <div className="mb-3 flex items-center gap-3 flex-wrap">
              <h2 className="text-base font-semibold text-gray-800">Project: {selectedProject}</h2>
              <span className="text-xs text-gray-400">
                {search
                  ? `${tableRows.filter(r => (r.item_code + ' ' + r.description).toLowerCase().includes(search.toLowerCase())).length} / ${tableRows.length} รายการ`
                  : `${tableRows.length} รายการ`}
              </span>
              <input
                type="text" value={search} onChange={e => setSearch(e.target.value)}
                placeholder="ค้นหา เช่น Pole, Connector..."
                className="ml-auto border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 w-56" />
              {tableRows.length > 0 && (
                <button
                  onClick={() => {
                    const filtered = search
                      ? tableRows.filter(r => (r.item_code + ' ' + r.description).toLowerCase().includes(search.toLowerCase()))
                      : tableRows
                    exportExcel(filtered)
                  }}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 transition-colors">
                  ↓ Export Excel
                </button>
              )}
            </div>
            <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-auto">
              {loadingTable ? (
                <div className="text-center py-12 text-gray-400">กำลังโหลด...</div>
              ) : tableRows.length === 0 ? (
                <div className="text-center py-12 text-gray-400">ไม่มีข้อมูลสำหรับ Project นี้</div>
              ) : (() => {
                const filteredRows = search
                  ? tableRows.filter(r => (r.item_code + ' ' + r.description).toLowerCase().includes(search.toLowerCase()))
                  : tableRows
                return filteredRows.length === 0 ? (
                  <div className="text-center py-12 text-gray-400">ไม่พบรายการที่ตรงกับ "{search}"</div>
                ) : (
                <table className="text-xs border-collapse" style={{ minWidth: 420 + suppliers.length * 270 }}>
                  <thead>
                    <tr className="bg-gray-800 text-white">
                      <th className="px-3 py-2.5 text-left sticky left-0 z-30 bg-gray-800 border-r border-gray-600 whitespace-nowrap"
                          rowSpan={2} style={{ minWidth: 160 }}>
                        Item Code
                      </th>
                      <th className="px-3 py-2.5 text-left sticky z-30 bg-gray-800 border-r border-gray-600"
                          rowSpan={2} style={{ minWidth: 240, left: 160 }}>
                        Description
                      </th>
                      {suppliers.map(s => (
                        <th key={s} colSpan={3} className="px-3 py-2 text-center border-l border-gray-600">
                          <div className="flex items-center justify-center gap-2">
                            <span className="font-semibold">{s}</span>
                            <button
                              onClick={() => requireUnlock(() => openManageModal(s))}
                              title="จัดการไฟล์ที่อัปโหลด"
                              className="text-gray-400 hover:text-white transition-colors text-xs leading-none">
                              ⋯
                            </button>
                          </div>
                          <div className="font-normal text-gray-400 text-xs">
                            {supplierDates[s] ? `as of ${fmtDate(supplierDates[s])}` : '—'}
                          </div>
                        </th>
                      ))}
                    </tr>
                    <tr className="bg-gray-700 text-gray-300">
                      {suppliers.map(s => {
                        const currency = tableRows.find(r => r.prices[s])?.prices[s]?.currency ?? '?'
                        return (
                          <Fragment key={s}>
                            <th className="px-3 py-1.5 text-right border-l border-gray-600 whitespace-nowrap" style={{ minWidth: 80 }}>
                              FOB {currency}
                            </th>
                            <th className="px-3 py-1.5 text-right whitespace-nowrap" style={{ minWidth: 90 }}>
                              FOB THB
                            </th>
                            <th className="px-3 py-1.5 text-right border-r border-gray-600 whitespace-nowrap" style={{ minWidth: 90 }}>
                              DDP THB
                            </th>
                          </Fragment>
                        )
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.map(row => {
                      const ddpValues = suppliers
                        .filter(s => row.prices[s])
                        .map(s => ({ s, ddp: row.prices[s].fob_price * getRate(row.prices[s].currency) * settings.ddp_multiplier }))
                      const minDdp = ddpValues.length > 1 ? Math.min(...ddpValues.map(v => v.ddp)) : null

                      return (
                        <tr key={row.item_code} className="border-b border-gray-100 hover:bg-blue-50/20 group">
                          <td className="px-3 py-2 font-mono text-gray-800 sticky left-0 bg-white group-hover:bg-blue-50 z-10 border-r border-gray-100 whitespace-nowrap">
                            {row.item_code}
                          </td>
                          <td className="px-3 py-2 text-gray-600 sticky bg-white group-hover:bg-blue-50 z-10 border-r border-gray-200" style={{ left: 160 }}>
                            {row.description}
                          </td>
                          {suppliers.map(s => {
                            const p = row.prices[s]
                            if (!p) return (
                              <Fragment key={s}>
                                <td className="px-3 py-2 text-center text-gray-200 border-l border-gray-100">—</td>
                                <td className="px-3 py-2 text-center text-gray-200">—</td>
                                <td className="px-3 py-2 text-center text-gray-200 border-r border-gray-100">—</td>
                              </Fragment>
                            )
                            const rate = getRate(p.currency)
                            const fobThb = p.fob_price * rate
                            const ddpThb = fobThb * settings.ddp_multiplier
                            const isBest = minDdp !== null && Math.abs(ddpThb - minDdp) < 0.005
                            return (
                              <Fragment key={s}>
                                <td
                                  className="px-3 py-2 text-right font-medium text-gray-700 border-l border-gray-100 cursor-pointer hover:text-blue-600 hover:underline"
                                  onClick={() => showHistory(row.item_code, s)}>
                                  {fmtN(p.fob_price)}
                                </td>
                                <td className="px-3 py-2 text-right text-gray-500">
                                  {fmtN(fobThb)}
                                </td>
                                <td className={`px-3 py-2 text-right font-semibold border-r border-gray-100 ${isBest ? 'text-green-600 bg-green-50' : 'text-gray-700'}`}>
                                  {fmtN(ddpThb)}
                                </td>
                              </Fragment>
                            )
                          })}
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
                )
              })()}
            </div>
          </>
        )}
      </div>

      {/* Manage Uploads Modal */}
      {managingSupplier && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg">
            <div className="flex items-center justify-between p-5 border-b border-gray-100">
              <div>
                <h3 className="font-semibold text-gray-800">จัดการไฟล์ที่อัปโหลด</h3>
                <p className="text-xs text-gray-500 mt-0.5">{selectedProject} — {managingSupplier}</p>
              </div>
              <button onClick={() => setManagingSupplier(null)} className="text-gray-300 hover:text-gray-500 text-xl">✕</button>
            </div>
            <div className="p-5 overflow-auto max-h-[60vh]">
              {loadingBatches ? (
                <p className="text-sm text-gray-400 text-center py-6">กำลังโหลด...</p>
              ) : batches.length === 0 ? (
                <p className="text-sm text-gray-400 text-center py-6">ไม่มีข้อมูล</p>
              ) : (
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-gray-400 text-xs border-b border-gray-100">
                      <th className="pb-2 text-left font-medium">อัปโหลดเมื่อ</th>
                      <th className="pb-2 text-left font-medium">Document</th>
                      <th className="pb-2 text-right font-medium">จำนวน item</th>
                      <th className="pb-2"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {batches.map((b, i) => (
                      <tr key={b.uploaded_at} className={`border-b border-gray-50 ${i === 0 ? 'font-medium' : 'text-gray-500'}`}>
                        <td className="py-2.5">{fmtDate(b.uploaded_at)}</td>
                        <td className="py-2.5 text-xs">{b.file_name || '—'}</td>
                        <td className="py-2.5 text-right">{b.count}</td>
                        <td className="py-2.5 text-right">
                          <button
                            onClick={() => {
                              if (confirm(`ลบไฟล์ที่อัปโหลดเมื่อ ${fmtDate(b.uploaded_at)} (${b.count} รายการ)?`))
                                deleteBatch(managingSupplier, b.uploaded_at)
                            }}
                            disabled={deletingBatch === b.uploaded_at}
                            className="text-xs text-red-400 hover:text-red-600 disabled:opacity-40">
                            {deletingBatch === b.uploaded_at ? 'กำลังลบ...' : 'ลบ'}
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}

      {/* History Modal */}
      {history && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg">
            <div className="flex items-center justify-between p-5 border-b border-gray-100">
              <div>
                <h3 className="font-semibold text-gray-800">ประวัติราคา</h3>
                <p className="text-xs text-gray-500 mt-0.5">{history.supplier} — {history.item_code}</p>
              </div>
              <button onClick={() => setHistory(null)} className="text-gray-300 hover:text-gray-500 text-xl">✕</button>
            </div>
            <div className="p-5 overflow-auto max-h-[60vh]">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-gray-400 text-xs border-b border-gray-100">
                    <th className="pb-2 text-left font-medium">บันทึกเมื่อ</th>
                    <th className="pb-2 text-left font-medium">Document</th>
                    <th className="pb-2 text-right font-medium">FOB Price</th>
                    <th className="pb-2 text-right font-medium">DDP THB</th>
                  </tr>
                </thead>
                <tbody>
                  {history.entries.map((e, i) => (
                    <tr key={e.id} className={`border-b border-gray-50 ${i === 0 ? 'font-medium' : 'text-gray-500'}`}>
                      <td className="py-2">{fmtDate(e.uploaded_at)}</td>
                      <td className="py-2 text-xs">{e.file_name || e.document_no || '—'}</td>
                      <td className="py-2 text-right">{fmtN(e.fob_price)} {e.currency}</td>
                      <td className="py-2 text-right font-semibold">{fmtN(e.fob_price * getRate(e.currency) * settings.ddp_multiplier)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

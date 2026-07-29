'use client'

import { useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'

interface ParsedRow {
  item_code: string
  description: string
  stock_thai: number
  po_thai: number
  lonsua: number
  week1: number
  week2: number
  week3_4: number
  next_month: number
}

interface DdpPrice {
  supplier: string
  ddp_thb: number
}

interface PlanRow extends ParsedRow {
  ddp_prices: DdpPrice[]
  L: number
  S: number
  T: number
  U: number
}

interface Settings {
  cny_rate: number
  usd_rate: number
  ddp_multiplier: number
}

interface PoDetail {
  poNo: string
  total: number
  notProduced: number
  inProduction: number
  finished: number
  onBoard: number
}

interface StockItem {
  total: number
  notProduced: number
  inProduction: number
  finished: number
  onBoard: number
  pos: PoDetail[]
}

interface SupplierStock {
  supplierName: string
  fileName: string
  items: Record<string, StockItem>
}

type LoadingPlan = Record<string, Record<string, [string, string]>>

interface StatusPopup {
  itemCode: string
  description: string
  supplierName: string
  stockItem: StockItem
}

interface HistorySession {
  id: string
  savedAt: string
  fileName: string
  itemCount: number
  parsedCache: ParsedRow[]
  loadingPlan: LoadingPlan
  supplierStocks: SupplierStock[]
  selectedProject: string
  supplierY: Record<string, string>
}

const HISTORY_KEY = 'order-plan-history'
const MAX_HISTORY = 20

function readHistory(): HistorySession[] {
  try { return JSON.parse(localStorage.getItem(HISTORY_KEY) ?? '[]') } catch { return [] }
}

export default function OrderPlanPage() {
  const [projects, setProjects] = useState<string[]>([])
  const [allDbSuppliers, setAllDbSuppliers] = useState<string[]>([])
  const [selectedProject, setSelectedProject] = useState('')
  const [settings, setSettings] = useState<Settings>({ cny_rate: 4.85, usd_rate: 33.00, ddp_multiplier: 1.11 })
  const [rows, setRows] = useState<PlanRow[]>([])
  const [parsedCache, setParsedCache] = useState<ParsedRow[]>([])
  const [ddpSuppliers, setDdpSuppliers] = useState<string[]>([])
  const [loading, setLoading] = useState(false)
  const [fileName, setFileName] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)

  const [supplierStocks, setSupplierStocks] = useState<SupplierStock[]>([])
  const [addingSupplier, setAddingSupplier] = useState(false)
  const [newSupplierName, setNewSupplierName] = useState('')
  const pendingSupplierRef = useRef<string>('')
  const addSupplierFileRef = useRef<HTMLInputElement>(null)
  const updateFileRefs = useRef<Record<string, HTMLInputElement | null>>({})

  const [loadingPlan, setLoadingPlan] = useState<LoadingPlan>({})
  const [supplierY, setSupplierY] = useState<Record<string, string>>({})
  const [statusPopup, setStatusPopup] = useState<StatusPopup | null>(null)

  const [history, setHistory] = useState<HistorySession[]>([])
  const [showHistory, setShowHistory] = useState(false)

  useEffect(() => {
    loadProjects()
    loadSettings()
    setHistory(readHistory())
  }, [])

  async function loadSettings() {
    const { data } = await supabase.from('cost_settings').select('key, value')
    if (data) {
      const m = Object.fromEntries((data as { key: string; value: string }[]).map(r => [r.key, r.value]))
      setSettings({ cny_rate: parseFloat(m.cny_rate ?? '4.85'), usd_rate: parseFloat(m.usd_rate ?? '33.00'), ddp_multiplier: parseFloat(m.ddp_multiplier ?? '1.11') })
    }
  }

  async function loadProjects() {
    const { data } = await supabase.from('po_items').select('project, supplier')
    if (data) {
      const rows = data as { project: string; supplier: string }[]
      setProjects([...new Set(rows.map(r => r.project))].sort())
      setAllDbSuppliers([...new Set(rows.map(r => r.supplier))].sort())
    }
  }

  async function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    setFileName(file.name)
    setLoading(true)
    const buffer = await file.arrayBuffer()
    const wb = XLSX.read(buffer, { type: 'array' })
    const ws = wb.Sheets[wb.SheetNames[0]]
    const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' })

    let dataStart = 1
    for (let i = 0; i < Math.min(raw.length, 10); i++) {
      const row = raw[i] as unknown[]
      const colA = String(row[0] ?? '').trim()
      const colB = String(row[1] ?? '').toLowerCase()
      if (colA === '#' || colB.includes('item') || colB.includes('code')) { dataStart = i + 1; break }
    }

    const parsed: ParsedRow[] = []
    for (let i = dataStart; i < raw.length; i++) {
      const row = raw[i] as unknown[]
      const itemCode = String(row[1] ?? '').trim()
      if (!itemCode || itemCode.toLowerCase().includes('total') || itemCode.toLowerCase().includes('รวม')) continue
      parsed.push({
        item_code: itemCode,
        description: String(row[2] ?? '').trim(),
        stock_thai: Number(row[3]) || 0,
        po_thai: Number(row[6]) || 0,
        lonsua: Number(row[12]) || 0,
        week1: Number(row[13]) || 0,
        week2: Number(row[14]) || 0,
        week3_4: Number(row[15]) || 0,
        next_month: Number(row[16]) || 0,
      })
    }

    setParsedCache(parsed)
    await buildPlanRows(parsed, selectedProject)
    setLoading(false)
  }

  async function buildPlanRows(parsed: ParsedRow[], project: string) {
    const priceMap = new Map<string, Map<string, { fob_price: number; currency: string }>>()
    const supplierSet = new Set<string>()

    if (project) {
      const { data } = await supabase.from('po_items').select('item_code, supplier, fob_price, currency').eq('project', project).order('uploaded_at', { ascending: false })
      if (data) {
        for (const item of data as { item_code: string; supplier: string; fob_price: number; currency: string }[]) {
          supplierSet.add(item.supplier)
          if (!priceMap.has(item.item_code)) priceMap.set(item.item_code, new Map())
          const sup = priceMap.get(item.item_code)!
          if (!sup.has(item.supplier)) sup.set(item.supplier, { fob_price: item.fob_price, currency: item.currency })
        }
      }
    }
    setDdpSuppliers([...supplierSet].sort())

    const { cny_rate, usd_rate, ddp_multiplier } = settings
    function toDdp(fob: number, currency: string) { return fob * (currency === 'USD' ? usd_rate : cny_rate) * ddp_multiplier }

    const planRows: PlanRow[] = parsed.map(r => {
      const supplierPrices = priceMap.get(r.item_code)
      const ddp_prices: DdpPrice[] = []
      if (supplierPrices) {
        for (const [supplier, { fob_price, currency }] of supplierPrices.entries())
          ddp_prices.push({ supplier, ddp_thb: toDdp(fob_price, currency) })
        ddp_prices.sort((a, b) => a.ddp_thb - b.ddp_thb)
      }
      const L = r.po_thai / 2
      const S = r.stock_thai + L - r.week1 - r.week2
      const T = S + r.lonsua - r.week3_4
      const U = T - r.next_month
      return { ...r, ddp_prices, L, S, T, U }
    })
    setRows(planRows)
  }

  async function handleProjectChange(project: string) {
    setSelectedProject(project)
    if (parsedCache.length > 0) { setLoading(true); await buildPlanRows(parsedCache, project); setLoading(false) }
  }

  async function parseSupplierStockFile(supplierName: string, file: File) {
    const buffer = await file.arrayBuffer()
    const wb = XLSX.read(buffer, { type: 'array' })
    const ws = wb.Sheets[wb.SheetNames[0]]
    const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' })

    let dataStart = 1
    for (let i = 0; i < Math.min(raw.length, 20); i++) {
      const row = raw[i] as unknown[]
      const colB = String(row[1] ?? '').toLowerCase()
      if (colB.includes('item') && colB.includes('code')) { dataStart = i + 1; break }
    }

    const itemMap = new Map<string, StockItem>()
    for (let i = dataStart; i < raw.length; i++) {
      const row = raw[i] as unknown[]
      const itemCode = String(row[1] ?? '').trim()
      if (!itemCode) continue
      const poNo = String(row[0] ?? '').trim()
      const total = Number(row[3]) || 0
      const notProduced = Number(row[5]) || 0
      const inProduction = Number(row[6]) || 0
      const finished = Number(row[7]) || 0
      const onBoard = Number(row[8]) || 0
      if (!itemMap.has(itemCode)) itemMap.set(itemCode, { total: 0, notProduced: 0, inProduction: 0, finished: 0, onBoard: 0, pos: [] })
      const item = itemMap.get(itemCode)!
      item.total += total; item.notProduced += notProduced; item.inProduction += inProduction
      item.finished += finished; item.onBoard += onBoard
      if (poNo && total > 0) item.pos.push({ poNo, total, notProduced, inProduction, finished, onBoard })
    }

    const items: Record<string, StockItem> = {}
    for (const [code, data] of itemMap.entries()) items[code] = data

    setSupplierStocks(prev => {
      const idx = prev.findIndex(s => s.supplierName === supplierName)
      const entry: SupplierStock = { supplierName, fileName: file.name, items }
      if (idx >= 0) { const u = [...prev]; u[idx] = entry; return u }
      return [...prev, entry]
    })
  }

  function setLoadQty(supplier: string, itemCode: string, batch: 0 | 1, value: string) {
    setLoadingPlan(prev => {
      const sup = { ...(prev[supplier] ?? {}) }
      const cur: [string, string] = [...(sup[itemCode] ?? ['', ''])] as [string, string]
      cur[batch] = value
      sup[itemCode] = cur
      return { ...prev, [supplier]: sup }
    })
  }

  function getStockItem(supplierName: string, itemCode: string): StockItem | null {
    return supplierStocks.find(s => s.supplierName === supplierName)?.items[itemCode] ?? null
  }

  function getLoadTotal(supplierName: string, itemCode: string): number {
    const plan = loadingPlan[supplierName]?.[itemCode] ?? ['', '']
    return (Number(plan[0]) || 0) + (Number(plan[1]) || 0)
  }

  function getRemaining(supplierName: string, itemCode: string): number {
    const stock = getStockItem(supplierName, itemCode)?.total ?? 0
    return stock - getLoadTotal(supplierName, itemCode)
  }

  function computeV(itemCode: string): number {
    return supplierStocks.reduce((s, ss) => s + getLoadTotal(ss.supplierName, itemCode), 0)
  }

  function computeW(row: PlanRow): number {
    return row.T + computeV(row.item_code)
  }

  function computeZ(row: PlanRow): number | null {
    const sel = supplierY[row.item_code]
    if (!sel) return null
    return computeW(row) - row.next_month - getRemaining(sel, row.item_code)
  }

  // ── History ──────────────────────────────────────────────────────────
  function saveSession() {
    if (parsedCache.length === 0) return
    const session: HistorySession = {
      id: Date.now().toString(),
      savedAt: new Date().toISOString(),
      fileName,
      itemCount: parsedCache.length,
      parsedCache,
      loadingPlan,
      supplierStocks,
      selectedProject,
      supplierY,
    }
    try {
      const existing = readHistory()
      const updated = [session, ...existing].slice(0, MAX_HISTORY)
      localStorage.setItem(HISTORY_KEY, JSON.stringify(updated))
      setHistory(updated)
    } catch { /* ignore storage errors */ }
  }

  async function restoreSession(session: HistorySession) {
    setParsedCache(session.parsedCache)
    setFileName(session.fileName)
    setLoadingPlan(session.loadingPlan)
    setSupplierStocks(session.supplierStocks)
    setSelectedProject(session.selectedProject)
    setSupplierY(session.supplierY)
    setShowHistory(false)
    setLoading(true)
    await buildPlanRows(session.parsedCache, session.selectedProject)
    setLoading(false)
  }

  function deleteSession(id: string) {
    const updated = history.filter(s => s.id !== id)
    localStorage.setItem(HISTORY_KEY, JSON.stringify(updated))
    setHistory(updated)
  }

  // ── Export Excel ──────────────────────────────────────────────────────
  function exportExcel() {
    const ddpCols = Math.min(ddpSuppliers.length, 5)
    const headers: string[] = [
      'Item Code', 'Description',
      ...Array.from({ length: ddpCols }, (_, i) => `DDP ${ddpSuppliers[i] ?? i + 1} (THB)`),
      'PO ไทย', 'Stock ไทย', 'PO ไทย/2', 'ลงเรือ', 'W1', 'W2', 'W3+4', 'Next Month',
      'เหลือให้ W3W4', 'เหลือให้ Next Month', 'ต้องสั่ง',
      ...supplierStocks.flatMap(ss => [`Stock ${ss.supplierName}`, 'โหลด 1', 'โหลด 2', 'คงเหลือ`${ss.supplierName}`']),
      ...(supplierStocks.length > 0 ? ['รวมโหลด', 'Stock หลังโหลด', 'เลือก Sup', 'แนะนำสั่ง PO'] : []),
    ]

    const dataRows = rows.map(row => {
      const V = computeV(row.item_code)
      const W = computeW(row)
      const Z = computeZ(row)
      return [
        row.item_code,
        row.description,
        ...Array.from({ length: ddpCols }, (_, i) => row.ddp_prices[i] ? parseFloat(row.ddp_prices[i].ddp_thb.toFixed(2)) : ''),
        row.po_thai, row.stock_thai,
        parseFloat(row.L.toFixed(1)),
        row.lonsua, row.week1, row.week2, row.week3_4, row.next_month,
        parseFloat(row.S.toFixed(1)), parseFloat(row.T.toFixed(1)), parseFloat(row.U.toFixed(1)),
        ...supplierStocks.flatMap(ss => {
          const stockQty = getStockItem(ss.supplierName, row.item_code)?.total ?? 0
          const plan = loadingPlan[ss.supplierName]?.[row.item_code] ?? ['', '']
          const remaining = getRemaining(ss.supplierName, row.item_code)
          return [stockQty || '', Number(plan[0]) || '', Number(plan[1]) || '', stockQty > 0 ? parseFloat(remaining.toFixed(1)) : '']
        }),
        ...(supplierStocks.length > 0 ? [V || '', parseFloat(W.toFixed(1)), supplierY[row.item_code] ?? '', Z !== null ? parseFloat(Z.toFixed(1)) : ''] : []),
      ]
    })

    const ws = XLSX.utils.aoa_to_sheet([headers, ...dataRows])
    ws['!cols'] = [{ wch: 18 }, { wch: 40 }, ...headers.slice(2).map(() => ({ wch: 14 }))]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Order Plan')
    XLSX.writeFile(wb, `Order_Plan_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  // ── Formatting ────────────────────────────────────────────────────────
  function fmtN(n: number) { return n > 0 ? n.toLocaleString(undefined, { maximumFractionDigits: 0 }) : '—' }
  function fmtF(n: number) { return n.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) }
  function fmtDdp(n: number) { return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) }

  const ddpCols = Math.min(ddpSuppliers.length, 5)
  const hasStock = supplierStocks.length > 0

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-6 text-sm sticky top-0 z-20 shadow-sm flex-wrap">
        <span className="font-bold text-gray-900">Import PO</span>
        <Link href="/" className="text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/calendar" className="text-gray-500 hover:text-gray-800 transition-colors">Calendar</Link>
        <Link href="/report" className="text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/compare" className="text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-gray-500 hover:text-gray-800 transition-colors">PO Builder</Link>
        <Link href="/order-plan" className="text-blue-600">Order Plan</Link>
        <div className="relative group">
          <span className="text-gray-500 cursor-default hover:text-gray-800">Summary ▾</span>
          <div className="absolute left-0 top-full pt-1 hidden group-hover:block z-50">
            <div className="bg-white border border-gray-200 rounded-lg shadow-lg py-1 min-w-[150px]">
              <Link href="/summary" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">Item Summary</Link>
              <Link href="/qc/summary" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">QC Summary</Link>
            </div>
          </div>
        </div>
        <Link href="/qc" className="text-gray-500 hover:text-gray-800 transition-colors">QC Report</Link>
        <Link href="/guide" className="text-gray-500 hover:text-gray-800 transition-colors">Guide</Link>
      </nav>

      <div className="max-w-full mx-auto px-6 py-8">
        <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">Order Plan (สั่งโหลด)</h1>
            <p className="text-sm text-gray-500 mt-1">อัพโหลด stock_dashboard Excel และเลือก Project เพื่อดูแผนการสั่งโหลด</p>
          </div>
          <div className="flex items-center gap-2">
            {rows.length > 0 && (
              <>
                <button onClick={saveSession}
                  className="px-4 py-2 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 transition-colors">
                  บันทึก
                </button>
                <button onClick={exportExcel}
                  className="px-4 py-2 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 transition-colors">
                  ↓ Export Excel
                </button>
              </>
            )}
            <button onClick={() => setShowHistory(true)}
              className="px-4 py-2 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 transition-colors flex items-center gap-1.5">
              ประวัติ {history.length > 0 && <span className="bg-gray-200 text-gray-600 text-xs rounded-full px-1.5">{history.length}</span>}
            </button>
          </div>
        </div>

        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mb-6">
          <div className="flex items-start gap-8 flex-wrap">
            <div className="flex flex-col gap-4">
              <div>
                <label className="text-sm text-gray-600 font-medium block mb-1.5">Project (ราคา DDP)</label>
                <select value={selectedProject} onChange={e => handleProjectChange(e.target.value)}
                  className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 bg-white min-w-[160px]">
                  <option value="">— เลือก Project —</option>
                  {projects.map(p => <option key={p} value={p}>{p}</option>)}
                </select>
              </div>
              <div>
                <label className="text-sm text-gray-600 font-medium block mb-1.5">stock_dashboard Excel</label>
                <div className="flex items-center gap-2">
                  <button onClick={() => fileRef.current?.click()}
                    className="px-4 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 transition-colors">
                    เลือกไฟล์
                  </button>
                  {fileName && <span className="text-sm text-gray-500 truncate max-w-[200px]">{fileName}</span>}
                  <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile} />
                </div>
              </div>
            </div>

            <div className="flex-1 min-w-[280px]">
              <label className="text-sm text-gray-600 font-medium block mb-1.5">Stock ที่ Supplier</label>
              <div className="space-y-2">
                {supplierStocks.map(ss => (
                  <div key={ss.supplierName} className="flex items-center gap-2 bg-gray-50 rounded-lg px-3 py-2">
                    <span className="text-xs font-semibold text-gray-700 min-w-[100px]">{ss.supplierName}</span>
                    <span className="text-xs text-gray-400 truncate flex-1">{ss.fileName}</span>
                    <button onClick={() => updateFileRefs.current[ss.supplierName]?.click()}
                      className="text-xs text-blue-500 hover:text-blue-700 whitespace-nowrap">อัพเดต</button>
                    <input type="file" accept=".xlsx,.xls" className="hidden"
                      ref={el => { updateFileRefs.current[ss.supplierName] = el }}
                      onChange={e => { const f = e.target.files?.[0]; if (f) parseSupplierStockFile(ss.supplierName, f) }} />
                    <button onClick={() => setSupplierStocks(prev => prev.filter(s => s.supplierName !== ss.supplierName))}
                      className="text-gray-300 hover:text-red-400 text-xs">✕</button>
                  </div>
                ))}

                {addingSupplier ? (
                  <div className="flex items-center gap-2 bg-blue-50 rounded-lg px-3 py-2">
                    <select value={newSupplierName} onChange={e => setNewSupplierName(e.target.value)}
                      className="border border-gray-300 rounded px-2 py-1 text-xs outline-none focus:border-blue-400 bg-white flex-1">
                      <option value="">— เลือก Supplier —</option>
                      {allDbSuppliers.filter(s => !supplierStocks.some(ss => ss.supplierName === s)).map(s =>
                        <option key={s} value={s}>{s}</option>)}
                    </select>
                    {newSupplierName && (
                      <button onClick={() => { pendingSupplierRef.current = newSupplierName; addSupplierFileRef.current?.click() }}
                        className="text-xs px-3 py-1 bg-blue-600 text-white rounded-lg hover:bg-blue-700 whitespace-nowrap">
                        Upload Stock
                      </button>
                    )}
                    <button onClick={() => { setAddingSupplier(false); setNewSupplierName('') }}
                      className="text-xs text-gray-400 hover:text-gray-600">ยกเลิก</button>
                  </div>
                ) : (
                  <button onClick={() => setAddingSupplier(true)}
                    className="text-xs text-blue-500 hover:text-blue-700 px-3 py-2 border border-dashed border-blue-200 rounded-lg w-full text-left hover:border-blue-400 transition-colors">
                    + เพิ่ม Supplier Stock
                  </button>
                )}

                <input type="file" accept=".xlsx,.xls" className="hidden" ref={addSupplierFileRef}
                  onChange={async e => {
                    const f = e.target.files?.[0]
                    if (f && pendingSupplierRef.current) {
                      await parseSupplierStockFile(pendingSupplierRef.current, f)
                      setAddingSupplier(false); setNewSupplierName(''); pendingSupplierRef.current = ''; e.target.value = ''
                    }
                  }} />
              </div>
            </div>
          </div>

          {rows.length > 0 && (
            <p className="mt-4 text-xs text-gray-400 border-t border-gray-100 pt-3">
              {rows.length} items{selectedProject ? ` · DDP จาก ${selectedProject}` : ''}
              {hasStock ? ` · Stock: ${supplierStocks.map(s => s.supplierName).join(', ')}` : ''}
            </p>
          )}
        </div>

        {loading && <div className="text-center py-16 text-gray-400 text-sm">กำลังโหลด...</div>}
        {!loading && rows.length === 0 && <div className="text-center py-16 text-gray-400 text-sm">อัพโหลด stock_dashboard เพื่อดูแผน</div>}

        {!loading && rows.length > 0 && (
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="text-xs border-collapse">
                <thead>
                  <tr className="bg-gray-50 text-gray-500 border-b border-gray-200">
                    <th className="px-3 py-2.5 text-left whitespace-nowrap font-semibold sticky left-0 bg-gray-50 z-10 border-r border-gray-200">Item Code</th>
                    <th className="px-3 py-2.5 text-left whitespace-nowrap font-semibold">Description</th>
                    {Array.from({ length: ddpCols }, (_, i) => (
                      <th key={i} className="px-3 py-2.5 text-right whitespace-nowrap font-semibold text-blue-600">DDP {i + 1}</th>
                    ))}
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">PO ไทย</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">Stock ไทย</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">PO ไทย/2</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">ลงเรือ</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">W1</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">W2</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">W3+4</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">Next Month</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-green-50 text-green-700">เหลือให้ W3W4</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-green-50 text-green-700">เหลือให้ Next</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-red-50 text-red-700">ต้องสั่ง</th>

                    {supplierStocks.map(ss => (
                      <>
                        <th key={`h-${ss.supplierName}-stk`} className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-indigo-50 text-indigo-700 border-l border-indigo-100">
                          Stock {ss.supplierName}
                        </th>
                        <th key={`h-${ss.supplierName}-l1`} className="px-3 py-2.5 text-center whitespace-nowrap font-semibold bg-indigo-50 text-indigo-600">โหลด 1</th>
                        <th key={`h-${ss.supplierName}-l2`} className="px-3 py-2.5 text-center whitespace-nowrap font-semibold bg-indigo-50 text-indigo-600">โหลด 2</th>
                        <th key={`h-${ss.supplierName}-rem`} className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-indigo-50 text-indigo-700">คงเหลือ</th>
                      </>
                    ))}

                    {hasStock && <>
                      <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-violet-50 text-violet-700 border-l border-violet-100">รวมโหลด</th>
                      <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-violet-50 text-violet-700">Stock หลังโหลด</th>
                      <th className="px-3 py-2.5 text-center whitespace-nowrap font-semibold">เลือก Sup</th>
                      <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-orange-50 text-orange-700">แนะนำสั่ง PO</th>
                    </>}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row, i) => {
                    const V = computeV(row.item_code)
                    const W = computeW(row)
                    const Z = computeZ(row)
                    const wNeg = W < 0
                    const zNeg = Z !== null && Z < 0

                    return (
                      <tr key={i} className="border-b border-gray-100 hover:bg-blue-50/20">
                        <td className="px-3 py-2 font-mono text-gray-800 whitespace-nowrap sticky left-0 bg-white border-r border-gray-100 z-10">{row.item_code}</td>
                        <td className="px-3 py-2 text-gray-600 whitespace-nowrap max-w-[220px] overflow-hidden text-ellipsis">{row.description || <span className="text-gray-300">—</span>}</td>

                        {Array.from({ length: ddpCols }, (_, j) => {
                          const p = row.ddp_prices[j]
                          return (
                            <td key={j} className="px-3 py-2 text-right whitespace-nowrap">
                              {p ? (
                                <span>
                                  <span className="text-gray-400 mr-1">{p.supplier}</span>
                                  <span className="text-gray-700 font-medium">{fmtDdp(p.ddp_thb)}</span>
                                </span>
                              ) : <span className="text-gray-300">—</span>}
                            </td>
                          )
                        })}

                        <td className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap">{fmtN(row.po_thai)}</td>
                        <td className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap">{fmtN(row.stock_thai)}</td>
                        <td className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap">{fmtF(row.L)}</td>
                        <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.lonsua)}</td>
                        <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.week1)}</td>
                        <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.week2)}</td>
                        <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.week3_4)}</td>
                        <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.next_month)}</td>
                        <td className="px-3 py-2 text-right font-medium bg-green-50/20 whitespace-nowrap text-gray-700">{fmtF(row.S)}</td>
                        <td className="px-3 py-2 text-right font-medium bg-green-50/20 whitespace-nowrap text-gray-700">{fmtF(row.T)}</td>
                        <td className={`px-3 py-2 text-right font-semibold bg-red-50/20 whitespace-nowrap ${row.U < 0 ? 'text-red-600' : 'text-gray-700'}`}>
                          {fmtF(row.U)}
                        </td>

                        {supplierStocks.map(ss => {
                          const stockItem = getStockItem(ss.supplierName, row.item_code)
                          const stockQty = stockItem?.total ?? 0
                          const plan = loadingPlan[ss.supplierName]?.[row.item_code] ?? ['', '']
                          const remaining = getRemaining(ss.supplierName, row.item_code)

                          return (
                            <>
                              <td key={`${ss.supplierName}-stk`} className="px-3 py-2 text-right bg-indigo-50/20 whitespace-nowrap border-l border-indigo-50">
                                {stockQty > 0 ? (
                                  <button onClick={() => stockItem && setStatusPopup({ itemCode: row.item_code, description: row.description, supplierName: ss.supplierName, stockItem })}
                                    className="text-indigo-600 hover:underline font-medium">
                                    {stockQty.toLocaleString()}
                                  </button>
                                ) : <span className="text-gray-300">—</span>}
                              </td>
                              <td key={`${ss.supplierName}-l1`} className="px-1.5 py-1 bg-indigo-50/10">
                                <input type="number" min="0" value={plan[0]}
                                  onChange={e => setLoadQty(ss.supplierName, row.item_code, 0, e.target.value)}
                                  placeholder="0"
                                  className="border border-gray-200 rounded px-1.5 py-1 w-20 text-right text-xs outline-none focus:border-indigo-400" />
                              </td>
                              <td key={`${ss.supplierName}-l2`} className="px-1.5 py-1 bg-indigo-50/10">
                                <input type="number" min="0" value={plan[1]}
                                  onChange={e => setLoadQty(ss.supplierName, row.item_code, 1, e.target.value)}
                                  placeholder="0"
                                  className="border border-gray-200 rounded px-1.5 py-1 w-20 text-right text-xs outline-none focus:border-indigo-400" />
                              </td>
                              <td key={`${ss.supplierName}-rem`} className={`px-3 py-2 text-right whitespace-nowrap bg-indigo-50/20 font-medium ${stockQty > 0 && remaining < 0 ? 'text-red-500' : 'text-gray-700'}`}>
                                {stockQty > 0 ? fmtF(remaining) : <span className="text-gray-300">—</span>}
                              </td>
                            </>
                          )
                        })}

                        {hasStock && <>
                          <td className="px-3 py-2 text-right whitespace-nowrap bg-violet-50/20 font-medium text-gray-700 border-l border-violet-50">
                            {V > 0 ? V.toLocaleString() : <span className="text-gray-300">—</span>}
                          </td>
                          <td className={`px-3 py-2 text-right whitespace-nowrap bg-violet-50/20 font-semibold ${wNeg ? 'text-red-600' : 'text-gray-700'}`}>
                            {wNeg && <span className="mr-1">⚠</span>}{fmtF(W)}
                          </td>
                          <td className="px-1.5 py-1">
                            <select value={supplierY[row.item_code] ?? ''}
                              onChange={e => setSupplierY(prev => ({ ...prev, [row.item_code]: e.target.value }))}
                              className="border border-gray-200 rounded px-1.5 py-1 text-xs outline-none focus:border-blue-400 bg-white min-w-[90px]">
                              <option value="">—</option>
                              {supplierStocks.map(ss => <option key={ss.supplierName} value={ss.supplierName}>{ss.supplierName}</option>)}
                            </select>
                          </td>
                          <td className={`px-3 py-2 text-right whitespace-nowrap bg-orange-50/20 font-semibold`}>
                            {Z === null ? (
                              <span className="text-gray-300 font-normal text-xs">เลือก Sup</span>
                            ) : zNeg ? (
                              <span className="text-orange-600">ควรสั่ง {Math.ceil(Math.abs(Z)).toLocaleString()}</span>
                            ) : (
                              <span className="text-green-600 font-normal">พอ +{fmtF(Z)}</span>
                            )}
                          </td>
                        </>}
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* ── Status popup ───────────────────────────────────────────────── */}
      {statusPopup && (
        <div className="fixed inset-0 bg-black/30 z-50 flex items-center justify-center p-4" onClick={() => setStatusPopup(null)}>
          <div className="bg-white rounded-xl shadow-xl p-6 max-w-md w-full" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-start mb-4">
              <div>
                <h3 className="font-bold text-gray-900 font-mono">{statusPopup.itemCode}</h3>
                <p className="text-sm text-gray-500 mt-0.5">{statusPopup.description}</p>
                <span className="text-xs text-indigo-600 font-medium">{statusPopup.supplierName}</span>
              </div>
              <button onClick={() => setStatusPopup(null)} className="text-gray-300 hover:text-gray-600 text-lg leading-none">✕</button>
            </div>
            <div className="grid grid-cols-2 gap-2 mb-4">
              {[
                { label: 'ยังไม่ผลิต', value: statusPopup.stockItem.notProduced, cls: 'bg-gray-50 text-gray-700' },
                { label: 'กำลังผลิต', value: statusPopup.stockItem.inProduction, cls: 'bg-yellow-50 text-yellow-800' },
                { label: 'ผลิตเสร็จแล้ว', value: statusPopup.stockItem.finished, cls: 'bg-green-50 text-green-800' },
                { label: 'On Board', value: statusPopup.stockItem.onBoard, cls: 'bg-blue-50 text-blue-800' },
              ].map(s => (
                <div key={s.label} className={`rounded-lg p-3 ${s.cls}`}>
                  <p className="text-xs opacity-70">{s.label}</p>
                  <p className="text-xl font-bold mt-0.5">{s.value > 0 ? s.value.toLocaleString() : '—'}</p>
                </div>
              ))}
            </div>
            <div className="flex justify-between items-center border-t border-gray-100 pt-3 mb-3">
              <span className="text-xs font-semibold text-gray-500">Total</span>
              <span className="text-sm font-bold text-gray-900">{statusPopup.stockItem.total.toLocaleString()}</span>
            </div>
            {statusPopup.stockItem.pos.length > 0 && (
              <div>
                <p className="text-xs font-semibold text-gray-400 mb-2">PO Detail</p>
                <div className="space-y-1 max-h-40 overflow-y-auto">
                  {statusPopup.stockItem.pos.map((po, i) => (
                    <div key={i} className="flex justify-between items-center text-xs py-1 border-b border-gray-50">
                      <span className="font-mono text-gray-700">{po.poNo}</span>
                      <span className="text-gray-500">{po.total.toLocaleString()}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── History modal ──────────────────────────────────────────────── */}
      {showHistory && (
        <div className="fixed inset-0 bg-black/30 z-50 flex items-center justify-center p-4" onClick={() => setShowHistory(false)}>
          <div className="bg-white rounded-xl shadow-xl p-6 max-w-lg w-full max-h-[80vh] flex flex-col" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-4">
              <h3 className="font-bold text-gray-900">ประวัติ Order Plan</h3>
              <button onClick={() => setShowHistory(false)} className="text-gray-300 hover:text-gray-600 text-lg leading-none">✕</button>
            </div>

            {history.length === 0 ? (
              <p className="text-sm text-gray-400 text-center py-8">ยังไม่มีประวัติ กด "บันทึก" เพื่อเก็บแผนปัจจุบัน</p>
            ) : (
              <div className="overflow-y-auto flex-1 space-y-2">
                {history.map(s => (
                  <div key={s.id} className="flex items-center justify-between bg-gray-50 rounded-lg px-4 py-3 gap-3">
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-gray-800 truncate">{s.fileName}</p>
                      <p className="text-xs text-gray-500 mt-0.5">
                        {new Date(s.savedAt).toLocaleString('th-TH', { dateStyle: 'medium', timeStyle: 'short' })}
                        {' · '}{s.itemCount} items
                        {s.selectedProject ? ` · ${s.selectedProject}` : ''}
                      </p>
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <button onClick={() => restoreSession(s)}
                        className="text-xs px-3 py-1.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors">
                        โหลด
                      </button>
                      <button onClick={() => deleteSession(s.id)}
                        className="text-xs text-gray-300 hover:text-red-400 transition-colors">✕</button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

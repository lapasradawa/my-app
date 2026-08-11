'use client'

import { useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { Workbook as ExcelJSWorkbook } from 'exceljs'
import { supabase } from '@/lib/supabase'
import { tryUnlock } from '@/lib/auth'

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
  fob_price: number
  currency: string
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

type LoadingPlan = Record<string, Record<string, string[]>>

interface StatusPopup {
  itemCode: string
  description: string
  supplierName: string
  stockItem: StockItem
}

// ── Formula popup ──────────────────────────────────────────────────────────
interface FLine {
  op: '' | '+' | '−' | '×' | '÷' | '='
  label: string
  val: number | string
  isResult?: boolean
  note?: string
}

interface FormulaPopup {
  itemCode: string
  description: string
  colName: string
  source?: string
  formulaStr?: string
  lines: FLine[]
}

interface UsageData {
  fileName: string
  items: Record<string, number[]>   // item_code → [m1, m2, m3] values
  labels: string[]                  // e.g. ['Apr', 'May', 'Jun']
}

// ── History ────────────────────────────────────────────────────────────────
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
  usageData?: UsageData | null
  supplierSlotDates?: Record<string, string[]>
  ddpExcluded?: string[]
  supplierCurrencyPref?: Record<string, string>
  usageLabel?: string
}

const THAI_COST = 'ทุนไทย'

const SUPPLIER_COLORS = [
  { text: 'text-blue-700',    bg: 'bg-blue-100',    dot: 'bg-blue-500',    hover: 'hover:bg-blue-50' },
  { text: 'text-emerald-700', bg: 'bg-emerald-100', dot: 'bg-emerald-500', hover: 'hover:bg-emerald-50' },
  { text: 'text-violet-700',  bg: 'bg-violet-100',  dot: 'bg-violet-500',  hover: 'hover:bg-violet-50' },
  { text: 'text-orange-700',  bg: 'bg-orange-100',  dot: 'bg-orange-500',  hover: 'hover:bg-orange-50' },
  { text: 'text-rose-700',    bg: 'bg-rose-100',    dot: 'bg-rose-500',    hover: 'hover:bg-rose-50' },
  { text: 'text-teal-700',    bg: 'bg-teal-100',    dot: 'bg-teal-500',    hover: 'hover:bg-teal-50' },
  { text: 'text-amber-700',   bg: 'bg-amber-100',   dot: 'bg-amber-500',   hover: 'hover:bg-amber-50' },
  { text: 'text-pink-700',    bg: 'bg-pink-100',    dot: 'bg-pink-500',    hover: 'hover:bg-pink-50' },
]

const HISTORY_KEY = 'order-plan-history'
const MAX_HISTORY = 20
const TEMPLATE_KEY = 'order-plan-template'

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
  const [formulaPopup, setFormulaPopup] = useState<FormulaPopup | null>(null)

  const [history, setHistory] = useState<HistorySession[]>([])
  const [showHistory, setShowHistory] = useState(false)
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(null)
  const [pwDialog, setPwDialog] = useState<{ action: 'save' } | { action: 'delete'; id: string } | null>(null)
  const [pwInput, setPwInput] = useState('')
  const [pwError, setPwError] = useState(false)

  const [allParsedCache, setAllParsedCache] = useState<ParsedRow[]>([])
  const [templateCodes, setTemplateCodes] = useState<Set<string> | null>(null)
  const [templateInfo, setTemplateInfo] = useState<{ fileName: string; count: number } | null>(null)
  const templateFileRef = useRef<HTMLInputElement>(null)

  const [usageData, setUsageData] = useState<UsageData | null>(null)
  const [usageLabel, setUsageLabel] = useState<string>('Jun 2026')
  const usageFileRef = useRef<HTMLInputElement>(null)

  const [bufferPct, setBufferPct] = useState(30)
  const [supplierCurrencyPref, setSupplierCurrencyPref] = useState<Record<string, string>>({})
  const [supplierAvailableCurrencies, setSupplierAvailableCurrencies] = useState<Record<string, string[]>>({})
  const [thaiCostMap, setThaiCostMap] = useState<Record<string, number>>({})
  const [ddpExcluded, setDdpExcluded] = useState<Set<string>>(new Set())
  const rawPriceCacheRef = useRef<{
    byItem: Map<string, Map<string, Map<string, number>>>
    latestCurrency: Map<string, string>
  } | null>(null)

  // no hard limit on loading slots
  const [supplierSlotDates, setSupplierSlotDates] = useState<Record<string, string[]>>({})
  const [slotDateDialog, setSlotDateDialog] = useState<{
    supplierName: string
    action: 'add' | 'edit' | 'delete'
    slotIdx: number
    step: 'pw' | 'date'
    pw: string
    dateInput: string
    pwError: boolean
  } | null>(null)

  useEffect(() => {
    loadProjects(); loadSettings(); setHistory(readHistory())
    try {
      const saved = localStorage.getItem(TEMPLATE_KEY)
      if (saved) {
        const data = JSON.parse(saved) as { codes: string[]; fileName: string }
        setTemplateCodes(new Set(data.codes))
        setTemplateInfo({ fileName: data.fileName, count: data.codes.length })
      }
    } catch { /* corrupt storage */ }
  }, [])

  // Reapply rows when currency preference or supplier exclusion changes
  useEffect(() => {
    if (parsedCache.length === 0 || !rawPriceCacheRef.current) return
    applyRows(parsedCache, supplierCurrencyPref, ddpExcluded)
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [supplierCurrencyPref])

  useEffect(() => {
    if (parsedCache.length === 0 || !rawPriceCacheRef.current) return
    applyRows(parsedCache, supplierCurrencyPref, ddpExcluded)
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ddpExcluded])

  // Auto-suggest slot 0: min(เหลือให้ Next Month, supplier stock) — only when slot is untouched
  useEffect(() => {
    if (rows.length === 0 || supplierStocks.length === 0) return
    setLoadingPlan(prev => {
      let changed = false
      const next = { ...prev }
      supplierStocks.forEach(ss => {
        const slots = supplierSlotDates[ss.supplierName] ?? []
        if (slots.length === 0) return
        const supPlan = { ...(prev[ss.supplierName] ?? {}) }
        rows.forEach(row => {
          const existing = supPlan[row.item_code] ?? []
          if (existing.some(v => v !== undefined && v !== '')) return
          const stockQty = getStockItem(ss.supplierName, row.item_code)?.total ?? 0
          if (stockQty <= 0) return
          const suggested = Math.min(Math.max(0, Math.ceil(row.T)), stockQty)
          if (suggested <= 0) return
          const cur = [...(supPlan[row.item_code] ?? [])]
          cur[0] = String(suggested)
          supPlan[row.item_code] = cur
          changed = true
        })
        next[ss.supplierName] = supPlan
      })
      return changed ? next : prev
    })
  }, [rows, supplierStocks, supplierSlotDates])

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
    const newId = Date.now().toString()
    setCurrentSessionId(newId)
    setFileName(file.name)
    setLoadingPlan({})
    setSupplierY({})
    setSupplierSlotDates({})
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

    const allParsed: ParsedRow[] = []
    for (let i = dataStart; i < raw.length; i++) {
      const row = raw[i] as unknown[]
      const itemCode = String(row[1] ?? '').trim()
      if (!itemCode || itemCode.toLowerCase().includes('total') || itemCode.toLowerCase().includes('รวม')) continue
      allParsed.push({
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

    setAllParsedCache(allParsed)
    const filtered = templateCodes ? allParsed.filter(r => templateCodes.has(r.item_code)) : allParsed
    setParsedCache(filtered)
    await buildPlanRows(filtered, selectedProject)

    // Auto-create new history entry for this upload
    const initSession: HistorySession = {
      id: newId, savedAt: new Date().toISOString(),
      fileName: file.name, itemCount: filtered.length,
      parsedCache: filtered, loadingPlan: {}, supplierStocks: [], selectedProject, supplierY: {},
    }
    setHistory(prev => {
      const updated = [initSession, ...prev].slice(0, MAX_HISTORY)
      localStorage.setItem(HISTORY_KEY, JSON.stringify(updated))
      return updated
    })
    setLoading(false)
  }

  async function handleTemplateFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    const buffer = await file.arrayBuffer()
    const wb = XLSX.read(buffer, { type: 'array' })
    const ws = wb.Sheets[wb.SheetNames[0]]
    const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' })

    // Find which column has item codes (header row), default col 0
    let codeCol = 0
    let dataStart = 0
    for (let i = 0; i < Math.min(raw.length, 10); i++) {
      const row = raw[i] as unknown[]
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] ?? '').toLowerCase()
        if (cell.includes('item') || cell === 'item code' || cell === 'itemcode') {
          codeCol = c; dataStart = i + 1; break
        }
      }
      if (dataStart > 0) break
    }

    const codes: string[] = []
    for (let i = dataStart; i < raw.length; i++) {
      const row = raw[i] as unknown[]
      const code = String(row[codeCol] ?? '').trim()
      if (code && !code.toLowerCase().includes('total') && !code.toLowerCase().includes('รวม')) {
        codes.push(code)
      }
    }

    const newSet = new Set(codes)
    setTemplateCodes(newSet)
    setTemplateInfo({ fileName: file.name, count: codes.length })
    localStorage.setItem(TEMPLATE_KEY, JSON.stringify({ codes, fileName: file.name }))

    // Re-filter if stock_dashboard already loaded
    if (allParsedCache.length > 0) {
      const filtered = allParsedCache.filter(r => newSet.has(r.item_code))
      setParsedCache(filtered)
      setLoading(true)
      await buildPlanRows(filtered, selectedProject)
      setLoading(false)
    }
    e.target.value = ''
  }

  function clearTemplate() {
    setTemplateCodes(null)
    setTemplateInfo(null)
    localStorage.removeItem(TEMPLATE_KEY)
    if (allParsedCache.length > 0) {
      setParsedCache(allParsedCache)
      buildPlanRows(allParsedCache, selectedProject)
    }
  }

  async function handleUsageFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    const buffer = await file.arrayBuffer()
    const wb = XLSX.read(buffer, { type: 'array' })
    const yearFromFile = (file.name.match(/20\d\d/) ?? [])[0] ?? String(new Date().getFullYear())

    // Pick the sheet with the most actual cell data (handles missing !ref)
    let ws = wb.Sheets[wb.SheetNames[0]]
    let maxCells = 0
    for (const name of wb.SheetNames) {
      const s = wb.Sheets[name]
      const n = Object.keys(s).filter(k => !k.startsWith('!')).length
      if (n > maxCells) { maxCells = n; ws = s }
    }

    // If !ref is missing or wrong, compute it from actual cell keys
    if (!ws['!ref'] || XLSX.utils.decode_range(ws['!ref']).e.r < 3) {
      const cellKeys = Object.keys(ws).filter(k => !k.startsWith('!') && /^[A-Z]+\d+$/.test(k))
      if (cellKeys.length > 0) {
        const decoded = cellKeys.map(k => XLSX.utils.decode_cell(k))
        ws['!ref'] = XLSX.utils.encode_range({
          s: { r: Math.min(...decoded.map(c => c.r)), c: Math.min(...decoded.map(c => c.c)) },
          e: { r: Math.max(...decoded.map(c => c.r)), c: Math.max(...decoded.map(c => c.c)) },
        })
      }
    }

    const wsRange = XLSX.utils.decode_range(ws['!ref'] ?? 'A1')

    // Robust cell reader: handles string, number, formula (cached .v), and formatted .w
    const readCell = (r: number, c: number): string => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })]
      if (!cell) return ''
      if (typeof cell.v === 'string' && cell.v.trim()) return cell.v.trim()
      if (typeof cell.v === 'number') return String(cell.v)
      // formula with no cached value → try formatted display
      const fmt = XLSX.utils.format_cell(cell).trim()
      return fmt || String(cell.w ?? '').trim()
    }
    const readNum = (r: number, c: number): number => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })]
      if (!cell) return 0
      if (typeof cell.v === 'number') return cell.v
      return Number(String(cell.v ?? cell.w ?? '').replace(/,/g, '')) || 0
    }

    // Find header row AND item-code column — search rows 0-15, cols 0-3
    let headerRowR = 0
    let itemCol = 0
    outer: for (let r = 0; r <= Math.min(15, wsRange.e.r); r++) {
      for (let c = 0; c <= Math.min(3, wsRange.e.c); c++) {
        if (/item/i.test(readCell(r, c))) { headerRowR = r; itemCol = c; break outer }
      }
    }

    // Find month columns from header row — search right of itemCol
    const MONTHS = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    const monthCols: number[] = []
    const monthLabels: string[] = []
    for (let c = itemCol + 1; c <= wsRange.e.c && monthCols.length < 3; c++) {
      const txt = readCell(headerRowR, c).toLowerCase().replace(/[^a-z]/g, '')
      const m = MONTHS.find(x => txt.startsWith(x))
      if (m) {
        monthCols.push(c)
        monthLabels.push(`${m.charAt(0).toUpperCase()}${m.slice(1)} ${yearFromFile}`)
      }
    }
    // Fallback: if < 3 months found, use offset from itemCol (C/D/E pattern)
    while (monthCols.length < 3) {
      const fallbackC = itemCol + 2 + monthCols.length
      monthCols.push(fallbackC)
      monthLabels.push(`เดือน ${monthCols.length}`)
    }

    const SKIP_CELL = /^(item|description|desc|total|รวม|usage|ลำดับ|no\.|#)$/i
    const items: Record<string, number[]> = {}
    for (let r = headerRowR + 1; r <= wsRange.e.r; r++) {
      const code = readCell(r, itemCol)
      if (!code || code.length > 60) continue
      if (SKIP_CELL.test(code.trim())) continue
      const vals = [readNum(r, monthCols[0]), readNum(r, monthCols[1]), readNum(r, monthCols[2])]
      if (vals.every(v => v === 0) && !code.match(/^[A-Z0-9]+-?[A-Z0-9]+/i)) continue
      if (code in items) items[code] = items[code].map((v, j) => v + vals[j])
      else items[code] = vals
    }

    const lastLabel = monthLabels[2] && !monthLabels[2].startsWith('เดือน') ? `Usage ${monthLabels[2]}` : `Usage Jun ${yearFromFile}`
    setUsageLabel(lastLabel)
    setUsageData({ fileName: file.name, items, labels: monthLabels })
    e.target.value = ''
  }

  function getUsageLast(itemCode: string): number | null {
    if (!usageData) return null
    const vals = usageData.items[itemCode]
    // Always use col E (index 2 = last/Jun) — data is locked to the latest month column
    return vals ? (vals[2] ?? null) : null
  }

  function getUsageAvg(itemCode: string): number | null {
    if (!usageData) return null
    const vals = usageData.items[itemCode]
    if (!vals || vals.length === 0) return null
    return vals.reduce((a, b) => a + b, 0) / vals.length
  }

  async function buildPlanRows(parsed: ParsedRow[], project: string, excludedOverride?: Set<string>, currencyPrefsOverride?: Record<string, string>) {
    // item_code → supplier → currency → fob (latest per supplier+currency, ORDER BY uploaded_at DESC)
    const byItem = new Map<string, Map<string, Map<string, number>>>()
    const latestCurrency = new Map<string, string>()   // supplier → most-recently-uploaded currency
    const supplierSet = new Set<string>()
    const supCurrenciesMap = new Map<string, Set<string>>()
    const thaiCostByItem = new Map<string, number>()

    if (project) {
      const { data } = await supabase.from('po_items').select('item_code, supplier, fob_price, currency').eq('project', project).order('uploaded_at', { ascending: false })
      if (data) {
        for (const item of data as { item_code: string; supplier: string; fob_price: number; currency: string }[]) {
          if (item.supplier === THAI_COST) {
            // ทุนไทย: store separately (latest wins via ORDER BY uploaded_at DESC)
            if (!thaiCostByItem.has(item.item_code)) thaiCostByItem.set(item.item_code, item.fob_price)
            continue
          }
          supplierSet.add(item.supplier)
          if (!latestCurrency.has(item.supplier)) latestCurrency.set(item.supplier, item.currency)
          if (!supCurrenciesMap.has(item.supplier)) supCurrenciesMap.set(item.supplier, new Set())
          supCurrenciesMap.get(item.supplier)!.add(item.currency)

          if (!byItem.has(item.item_code)) byItem.set(item.item_code, new Map())
          const bySupplier = byItem.get(item.item_code)!
          if (!bySupplier.has(item.supplier)) bySupplier.set(item.supplier, new Map())
          const byCur = bySupplier.get(item.supplier)!
          if (!byCur.has(item.currency)) byCur.set(item.currency, item.fob_price)
        }
      }
    }
    rawPriceCacheRef.current = { byItem, latestCurrency }
    setDdpSuppliers([...supplierSet].sort())
    const newSupCur: Record<string, string[]> = {}
    supCurrenciesMap.forEach((cs, s) => { newSupCur[s] = [...cs].sort() })
    setSupplierAvailableCurrencies(newSupCur)
    const tcObj: Record<string, number> = {}
    thaiCostByItem.forEach((price, code) => { tcObj[code] = price })
    setThaiCostMap(tcObj)

    applyRows(parsed, currencyPrefsOverride ?? supplierCurrencyPref, excludedOverride ?? ddpExcluded)
  }

  function applyRows(parsed: ParsedRow[], currencyPrefs: Record<string, string>, excluded: Set<string> = new Set()) {
    const cache = rawPriceCacheRef.current
    if (!cache) return
    const { byItem, latestCurrency } = cache
    const { cny_rate, usd_rate, ddp_multiplier } = settings
    const toDdp = (fob: number, currency: string) => fob * (currency === 'USD' ? usd_rate : cny_rate) * ddp_multiplier

    const planRows: PlanRow[] = parsed.map(r => {
      const bySupplier = byItem.get(r.item_code)
      const ddp_prices: DdpPrice[] = []
      if (bySupplier) {
        for (const [supplier, byCur] of bySupplier.entries()) {
          if (excluded.has(supplier)) continue
          const pref = currencyPrefs[supplier]
          const useCur = pref && byCur.has(pref) ? pref : (latestCurrency.get(supplier) ?? '')
          const fob = byCur.get(useCur)
          if (fob === undefined) continue
          ddp_prices.push({ supplier, ddp_thb: toDdp(fob, useCur), fob_price: fob, currency: useCur })
        }
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

  function supplierColor(supplier: string) {
    const idx = ddpSuppliers.indexOf(supplier)
    return SUPPLIER_COLORS[Math.max(0, idx) % SUPPLIER_COLORS.length]
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
      const notProduced = Number(row[4]) || 0
      const inProduction = Number(row[5]) || 0
      const finished = Number(row[6]) || 0
      const onBoard = Number(row[7]) || 0
      if (!itemMap.has(itemCode)) itemMap.set(itemCode, { total: 0, notProduced: 0, inProduction: 0, finished: 0, onBoard: 0, pos: [] })
      const item = itemMap.get(itemCode)!
      item.total += total; item.notProduced += notProduced; item.inProduction += inProduction
      item.finished += finished; item.onBoard += onBoard
      if (poNo && (total + notProduced + inProduction + finished + onBoard > 0))
        item.pos.push({ poNo, total, notProduced, inProduction, finished, onBoard })
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

  function setLoadQty(supplier: string, itemCode: string, slotIdx: number, value: string) {
    setLoadingPlan(prev => {
      const sup = { ...(prev[supplier] ?? {}) }
      const cur = [...(sup[itemCode] ?? [])]
      cur[slotIdx] = value
      sup[itemCode] = cur
      return { ...prev, [supplier]: sup }
    })
  }

  function openSlotDateDialog(supplierName: string, action: 'add' | 'edit' | 'delete', slotIdx: number) {
    const currentDate = action !== 'add' ? (supplierSlotDates[supplierName]?.[slotIdx] ?? '') : ''
    setSlotDateDialog({ supplierName, action, slotIdx, step: 'pw', pw: '', dateInput: currentDate, pwError: false })
  }

  function confirmSlotDatePw() {
    if (!slotDateDialog) return
    if (!tryUnlock(slotDateDialog.pw)) { setSlotDateDialog(d => d ? { ...d, pwError: true } : null); return }
    if (slotDateDialog.action === 'delete') { confirmSlotDateAction(); return }
    setSlotDateDialog(d => d ? { ...d, step: 'date', pwError: false } : null)
  }

  function confirmSlotDateAction() {
    if (!slotDateDialog) return
    const { supplierName, action, slotIdx, dateInput } = slotDateDialog
    if (action === 'add' || action === 'edit') {
      if (!dateInput.trim()) return
      setSupplierSlotDates(prev => {
        const cur = [...(prev[supplierName] ?? [])]
        if (action === 'add') cur.push(dateInput.trim())
        else cur[slotIdx] = dateInput.trim()
        return { ...prev, [supplierName]: cur }
      })
    } else {
      setSupplierSlotDates(prev => {
        const cur = [...(prev[supplierName] ?? [])]
        cur.splice(slotIdx, 1)
        return { ...prev, [supplierName]: cur }
      })
      setLoadingPlan(prev => {
        const sup = { ...(prev[supplierName] ?? {}) }
        Object.keys(sup).forEach(code => { const p = [...(sup[code] ?? [])]; p.splice(slotIdx, 1); sup[code] = p })
        return { ...prev, [supplierName]: sup }
      })
    }
    setSlotDateDialog(null)
  }

  function getStockItem(supplierName: string, itemCode: string): StockItem | null {
    return supplierStocks.find(s => s.supplierName === supplierName)?.items[itemCode] ?? null
  }

  function getLoadTotal(supplierName: string, itemCode: string): number {
    return (loadingPlan[supplierName]?.[itemCode] ?? []).reduce((s, v) => s + (Number(v) || 0), 0)
  }

  function getRemaining(supplierName: string, itemCode: string): number {
    return (getStockItem(supplierName, itemCode)?.total ?? 0) - getLoadTotal(supplierName, itemCode)
  }

  function computeV(itemCode: string): number {
    return supplierStocks.reduce((s, ss) => s + getLoadTotal(ss.supplierName, itemCode), 0)
  }

  function computeW(row: PlanRow): number { return row.T + computeV(row.item_code) }

  function getTotalRemaining(itemCode: string): number {
    return supplierStocks.reduce((s, ss) => s + getRemaining(ss.supplierName, itemCode), 0)
  }

  function computeZ(row: PlanRow): number {
    return getTotalRemaining(row.item_code) + computeW(row) - row.next_month
  }

  // ── Formula builder ────────────────────────────────────────────────────
  function buildFormula(row: PlanRow, colType: string, supplierName?: string): FormulaPopup {
    const base = { itemCode: row.item_code, description: row.description }

    const src = (colName: string, col: string, val: number): FormulaPopup => ({
      ...base, colName,
      source: `ดึงจาก stock_dashboard คอลัมน์ ${col}`,
      lines: [{ op: '=', label: colName, val, isResult: true }],
    })

    switch (colType) {
      case 'po_thai':  return src('PO ไทย', 'G', row.po_thai)
      case 'stock_thai': return src('Stock ไทย', 'D', row.stock_thai)
      case 'lonsua':   return src('ลงเรือ', 'M', row.lonsua)
      case 'week1':    return src('Fc. W1', 'N', row.week1)
      case 'week2':    return src('Fc. W2', 'O', row.week2)
      case 'week3_4':  return src('Fc. W3+4', 'P', row.week3_4)
      case 'next_month': return src('Fc. Next Month', 'Q', row.next_month)

      case 'L':
        return { ...base, colName: 'PO ไทย/2', formulaStr: 'PO ไทย ÷ 2', lines: [
          { op: '', label: 'PO ไทย', val: row.po_thai },
          { op: '÷', label: '2', val: 2 },
          { op: '=', label: 'PO ไทย/2', val: row.L, isResult: true },
        ]}

      case 'S':
        return { ...base, colName: 'เหลือให้ W3W4', formulaStr: 'Stock ไทย + PO ไทย/2 − Fc.W1 − Fc.W2', lines: [
          { op: '', label: 'Stock ไทย', val: row.stock_thai },
          { op: '+', label: 'PO ไทย/2', val: row.L },
          { op: '−', label: 'Fc. W1', val: row.week1 },
          { op: '−', label: 'Fc. W2', val: row.week2 },
          { op: '=', label: 'เหลือให้ W3W4', val: row.S, isResult: true },
        ]}

      case 'T':
        return { ...base, colName: 'เหลือให้ Next Month', formulaStr: 'เหลือให้ W3W4 + ลงเรือ − Fc.W3+4', lines: [
          { op: '', label: 'เหลือให้ W3W4 (S)', val: row.S },
          { op: '+', label: 'ลงเรือ', val: row.lonsua },
          { op: '−', label: 'Fc. W3+4', val: row.week3_4 },
          { op: '=', label: 'เหลือให้ Next Month', val: row.T, isResult: true },
        ]}

      case 'U':
        return { ...base, colName: 'ต้องสั่งโหลด', formulaStr: 'เหลือให้ Next Month − Fc. Next Month', lines: [
          { op: '', label: 'เหลือให้ Next Month (T)', val: row.T },
          { op: '−', label: 'Fc. Next Month', val: row.next_month },
          { op: '=', label: 'ต้องสั่งโหลด', val: row.U, isResult: true },
        ]}

      case 'V': {
        const lines: FLine[] = []
        let first = true
        supplierStocks.forEach(ss => {
          const dates = supplierSlotDates[ss.supplierName] ?? []
          const plan = loadingPlan[ss.supplierName]?.[row.item_code] ?? []
          dates.forEach((date, i) => {
            const qty = Number(plan[i]) || 0
            if (qty > 0) {
              lines.push({ op: first ? '' : '+', label: `${ss.supplierName} — ${date}`, val: qty })
              first = false
            }
          })
        })
        const V = computeV(row.item_code)
        if (lines.length === 0) lines.push({ op: '', label: '(ยังไม่ได้กรอกแผนโหลด)', val: 0 })
        lines.push({ op: '=', label: 'รวมโหลด', val: V, isResult: true })
        return { ...base, colName: 'รวมโหลด', formulaStr: 'Σ ทุกวันโหลด ทุก Supplier', lines }
      }

      case 'W': {
        const V = computeV(row.item_code)
        return { ...base, colName: 'Stock หลังโหลด', formulaStr: 'เหลือให้ Next Month + รวมโหลด', lines: [
          { op: '', label: 'เหลือให้ Next Month (T)', val: row.T },
          { op: '+', label: 'รวมโหลด (V)', val: V },
          { op: '=', label: 'Stock หลังโหลด', val: row.T + V, isResult: true },
        ]}
      }

      case 'remaining': {
        const sn = supplierName ?? ''
        const stockQty = getStockItem(sn, row.item_code)?.total ?? 0
        const dates = supplierSlotDates[sn] ?? []
        const plan = loadingPlan[sn]?.[row.item_code] ?? []
        const lines: FLine[] = [{ op: '', label: `Stock ${sn}`, val: stockQty }]
        dates.forEach((date, i) => lines.push({ op: '−', label: date, val: Number(plan[i]) || 0 }))
        const remaining = getRemaining(sn, row.item_code)
        lines.push({ op: '=', label: `คงเหลือ ${sn}`, val: remaining, isResult: true })
        return { ...base, colName: `คงเหลือ (${sn})`, formulaStr: `Stock ${sn} − Σ ทุกวันโหลด`, lines }
      }

      case 'Z': {
        const W = computeW(row)
        const Z = computeZ(row)
        const stockLines: FLine[] = supplierStocks
          .map(ss => ({ name: ss.supplierName, qty: getRemaining(ss.supplierName, row.item_code) }))
          .filter(x => getStockItem(x.name, row.item_code) !== null)
          .map((x, idx) => ({ op: (idx === 0 ? '' : '+') as FLine['op'], label: `คงเหลือ ${x.name}`, val: x.qty }))
        if (stockLines.length === 0) stockLines.push({ op: '', label: 'คงเหลือ Supplier (ไม่มีข้อมูล)', val: 0 })
        return { ...base, colName: 'PO Coverage for Fc. Next Month', formulaStr: 'รวม คงเหลือ ทุก Supplier + Stock หลังโหลด − Fc. Next Month', lines: [
          ...stockLines,
          { op: '+', label: 'Stock หลังโหลด (W)', val: W },
          { op: '−', label: 'Fc. Next Month', val: row.next_month },
          { op: '=', label: 'PO Coverage for Fc. Next Month', val: Z, isResult: true, note: Z < 0 ? `ควรสั่ง ${Math.ceil(Math.abs(Z)).toLocaleString()} ชิ้น` : 'stock เพียงพอ' },
        ]}
      }

      case 'buf': {
        const buf = row.next_month * bufferPct / 100
        return { ...base, colName: `${bufferPct}% Buffer Stock`, formulaStr: `Fc. Next Month × ${bufferPct}%`, lines: [
          { op: '', label: 'Fc. Next Month', val: row.next_month },
          { op: '×', label: `${bufferPct}%`, val: bufferPct / 100 },
          { op: '=', label: `${bufferPct}% Buffer Stock`, val: buf, isResult: true },
        ]}
      }

      case 'bufNeeded': {
        const Z = computeZ(row)
        const buf = row.next_month * bufferPct / 100
        const bufNeeded = buf - Z
        return { ...base, colName: `รวมที่ต้องสั่ง`, formulaStr: `${bufferPct}% Buffer Stock − PO Coverage`, lines: [
          { op: '', label: `${bufferPct}% Buffer Stock`, val: buf },
          { op: '−', label: `PO Coverage (${Z >= 0 ? `+${fmtF(Z)} → หักออก` : `${fmtF(Z)} → ติดลบ รวมเพิ่ม`})`, val: Z },
          { op: '=', label: `รวมที่ต้องสั่ง`, val: bufNeeded, isResult: true, note: bufNeeded > 0 ? `ต้องสั่ง ${Math.ceil(bufNeeded).toLocaleString()} ชิ้น` : 'ไม่ต้องสั่งเพิ่ม' },
        ]}
      }

      default:
        return { ...base, colName: colType, lines: [] }
    }
  }

  function showDdpFormula(row: PlanRow, p: DdpPrice) {
    const rate = p.currency === 'USD' ? settings.usd_rate : settings.cny_rate
    const currencyLabel = p.currency === 'USD' ? `USD (${settings.usd_rate} ฿/USD)` : `CNY (${settings.cny_rate} ฿/CNY)`
    setFormulaPopup({
      itemCode: row.item_code, description: row.description,
      colName: `DDP — ${p.supplier}`,
      formulaStr: 'FOB Price × อัตราแลกเปลี่ยน × DDP Multiplier',
      lines: [
        { op: '', label: `FOB Price (${p.currency})`, val: p.fob_price },
        { op: '×', label: currencyLabel, val: rate },
        { op: '×', label: `DDP Multiplier`, val: settings.ddp_multiplier },
        { op: '=', label: 'DDP (THB)', val: p.ddp_thb, isResult: true },
      ],
    })
  }

  // ── History ────────────────────────────────────────────────────────────
  function doSaveSession() {
    if (parsedCache.length === 0) return
    const id = currentSessionId ?? Date.now().toString()
    if (!currentSessionId) setCurrentSessionId(id)
    const session: HistorySession = {
      id, savedAt: new Date().toISOString(),
      fileName, itemCount: parsedCache.length, parsedCache,
      loadingPlan, supplierStocks, selectedProject, supplierY, usageData, supplierSlotDates,
      ddpExcluded: [...ddpExcluded],
      supplierCurrencyPref,
      usageLabel,
    }
    setHistory(prev => {
      const idx = prev.findIndex(s => s.id === id)
      const updated = idx >= 0
        ? prev.map((s, i) => i === idx ? session : s)
        : [session, ...prev].slice(0, MAX_HISTORY)
      localStorage.setItem(HISTORY_KEY, JSON.stringify(updated))
      return updated
    })
  }

  function doDeleteSession(id: string) {
    setHistory(prev => {
      const updated = prev.filter(s => s.id !== id)
      localStorage.setItem(HISTORY_KEY, JSON.stringify(updated))
      return updated
    })
  }

  function confirmPasswordAction() {
    if (!tryUnlock(pwInput)) { setPwError(true); return }
    if (pwDialog?.action === 'save') doSaveSession()
    else if (pwDialog?.action === 'delete') doDeleteSession(pwDialog.id)
    setPwDialog(null); setPwInput(''); setPwError(false)
  }

  async function restoreSession(session: HistorySession) {
    setCurrentSessionId(session.id)
    setParsedCache(session.parsedCache); setFileName(session.fileName)
    setLoadingPlan(session.loadingPlan); setSupplierStocks(session.supplierStocks)
    setSelectedProject(session.selectedProject); setSupplierY(session.supplierY)
    if (session.usageData !== undefined) setUsageData(session.usageData ?? null)
    setSupplierSlotDates(session.supplierSlotDates ?? {})
    const restoredExcluded = new Set<string>(session.ddpExcluded ?? [])
    const restoredCurrencyPref = session.supplierCurrencyPref ?? {}
    setDdpExcluded(restoredExcluded)
    setSupplierCurrencyPref(restoredCurrencyPref)
    if (session.usageLabel) setUsageLabel(session.usageLabel)
    setShowHistory(false); setLoading(true)
    await buildPlanRows(session.parsedCache, session.selectedProject, restoredExcluded, restoredCurrencyPref)
    setLoading(false)
  }

  // ── Export Excel ───────────────────────────────────────────────────────
  async function exportExcel() {
    const workbook = new ExcelJSWorkbook()
    const ws = workbook.addWorksheet('Order Plan')

    const exportDdpSuppliers = activeDdpSuppliers.slice(0, 8)
    const exportDdpCols = exportDdpSuppliers.length
    // Fixed color per supplier name — stable regardless of order or exclusions
    const SUP_COLOR: Record<string, { bg: string; fg: string }> = {
      'KNCD':     { bg: 'C6EFC5', fg: '166534' },
      'LITELON':  { bg: 'FDDCB5', fg: '9A3412' },
      'MK':       { bg: 'D5B8FF', fg: '4C1D95' },
      'SGL':      { bg: 'FFF2CC', fg: '92400E' },
      'YONGGUAN': { bg: 'DCDCDC', fg: '374151' },
      'YG':       { bg: 'DCDCDC', fg: '374151' },
      'YPN':      { bg: 'BDD7EE', fg: '1E3A8A' },
    }
    // Fallback palette for unknown suppliers
    const FALLBACK_BG = ['FFD0D0','B3E5FC','FEF9C3','E0F2FE','F3E8FF','FCE7F3']
    const supBg = (s: string) => SUP_COLOR[s]?.bg ?? FALLBACK_BG[ddpSuppliers.indexOf(s) % FALLBACK_BG.length]
    const supFg = (s: string) => SUP_COLOR[s]?.fg ?? '374151'

    // Row 1: legend — "Supplier Color:" then one colored cell per supplier
    ws.getColumn(1).width = 20
    ws.getColumn(2).width = 42
    const legendEntries: string[] = [...(hasThaiCost ? [THAI_COST] : []), ...exportDdpSuppliers]
    const legendRow = ws.addRow(['Supplier Color:', ...legendEntries])
    legendRow.getCell(1).font = { bold: true }
    if (hasThaiCost) {
      const thaiCell = legendRow.getCell(2)
      thaiCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF99F6E4' } }
      thaiCell.font = { bold: true, color: { argb: 'FF0F766E' } }
      thaiCell.alignment = { horizontal: 'center' }
    }
    exportDdpSuppliers.forEach((s, i) => {
      const cell = legendRow.getCell(2 + (hasThaiCost ? 1 : 0) + i)
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + supBg(s) } }
      cell.font = { bold: true, color: { argb: 'FF' + supFg(s) } }
      cell.alignment = { horizontal: 'center' }
    })
    // Row 2: spacer
    ws.addRow([])

    const ddpHeaders = Array.from({ length: exportDdpCols }, (_, i) => {
      if (exportDdpCols === 1) return 'DDP Price (THB)'
      if (i === 0) return 'DDP Price Cheapest (THB)'
      if (i === exportDdpCols - 1) return 'DDP Price Most Expensive (THB)'
      return `DDP Price Rank ${i + 1} (THB)`
    })

    const headers = [
      'Item Code', 'Description',
      ...(hasThaiCost ? ['ทุนไทย (THB)'] : []),
      ...ddpHeaders,
      'PO ไทย', 'Stock ไทย', 'PO ไทย/2', 'ลงเรือ', 'Fc. W1', 'Fc. W2', 'Fc. W3+4', 'Fc. Next Month',
      ...(usageData ? [usageLabel, 'Avg. Usage 3M'] : []),
      'เหลือให้ W3W4', 'เหลือให้ Next Month', 'ต้องสั่งโหลด',
      ...supplierStocks.flatMap(ss => {
        const dates = supplierSlotDates[ss.supplierName] ?? []
        return [`Stock ${ss.supplierName}`, ...dates, `คงเหลือ ${ss.supplierName}`]
      }),
      ...(supplierStocks.length > 0 ? ['รวมโหลด', 'Stock หลังโหลด', 'PO Coverage for Fc. Next Month', `${bufferPct}% Buffer Stock (Based on Fc. Next Month)`, 'รวมที่ต้องสั่ง (Buffer + ขาด)'] : []),
    ]

    const headerRow = ws.addRow(headers)
    headerRow.font = { bold: true }
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }
    headers.slice(2).forEach((_, i) => { ws.getColumn(i + 3).width = 16 })

    rows.forEach(row => {
      const V = computeV(row.item_code)
      const W = computeW(row)
      const Z = computeZ(row)
      const rowData = [
        row.item_code, row.description,
        ...(hasThaiCost ? [thaiCostMap[row.item_code] ?? ''] : []),
        ...Array.from({ length: exportDdpCols }, (_, i) => row.ddp_prices[i] ? parseFloat(row.ddp_prices[i].ddp_thb.toFixed(2)) : ''),
        row.po_thai, row.stock_thai, parseFloat(row.L.toFixed(1)),
        row.lonsua, row.week1, row.week2, row.week3_4, row.next_month,
        ...(usageData ? [
          getUsageLast(row.item_code) ?? '',
          (() => { const a = getUsageAvg(row.item_code); return a !== null ? parseFloat(a.toFixed(1)) : '' })(),
        ] : []),
        parseFloat(row.S.toFixed(1)), parseFloat(row.T.toFixed(1)), parseFloat(row.U.toFixed(1)),
        ...supplierStocks.flatMap(ss => {
          const stockQty = getStockItem(ss.supplierName, row.item_code)?.total ?? 0
          const dates = supplierSlotDates[ss.supplierName] ?? []
          const plan = loadingPlan[ss.supplierName]?.[row.item_code] ?? []
          const remaining = getRemaining(ss.supplierName, row.item_code)
          return [stockQty || '', ...dates.map((_, i) => Number(plan[i]) || ''), stockQty > 0 ? parseFloat(remaining.toFixed(1)) : '']
        }),
        ...(supplierStocks.length > 0 ? (() => {
          const buf = parseFloat((row.next_month * bufferPct / 100).toFixed(1))
          const bufNeeded = parseFloat((buf - Z).toFixed(1))
          return [V || '', parseFloat(W.toFixed(1)), parseFloat(Z.toFixed(1)), buf, bufNeeded]
        })() : []),
      ]
      const exRow = ws.addRow(rowData)
      const ddpColOffset = 3 + (hasThaiCost ? 1 : 0)
      if (hasThaiCost) {
        const thaiPrice = thaiCostMap[row.item_code]
        if (thaiPrice != null) {
          const ddp1Price = row.ddp_prices[0]?.ddp_thb ?? null
          const isCheaper = ddp1Price != null && thaiPrice < ddp1Price
          exRow.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isCheaper ? 'FFFF4444' : 'FF99F6E4' } }
          exRow.getCell(3).font = { color: { argb: isCheaper ? 'FFFFFFFF' : 'FF0F766E' }, bold: isCheaper }
        }
      }
      Array.from({ length: exportDdpCols }, (_, i) => {
        const p = row.ddp_prices[i]
        if (!p) return
        const cell = exRow.getCell(ddpColOffset + i)
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + supBg(p.supplier) } }
        cell.font = { color: { argb: 'FF' + supFg(p.supplier) }, bold: true }
      })
    })

    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `Order_Plan_${new Date().toISOString().slice(0, 10)}.xlsx`
    a.click()
    URL.revokeObjectURL(url)
  }

  // ── Formatting ─────────────────────────────────────────────────────────
  function fmtN(n: number) { return n > 0 ? n.toLocaleString(undefined, { maximumFractionDigits: 0 }) : '—' }
  function fmtF(n: number) { return n.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) }
  function fmtDdp(n: number) { return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) }
  function fmtVal(v: number | string, dec = 1) {
    if (typeof v === 'string') return v
    return v.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec })
  }

  const activeDdpSuppliers = ddpSuppliers.filter(s => !ddpExcluded.has(s))
  const ddpCols = Math.min(activeDdpSuppliers.length, 8)
  const hasThaiCost = Object.keys(thaiCostMap).length > 0
  const hasStock = supplierStocks.length > 0

  // Clickable cell wrapper
  const C = ({ children, onClick, className = '' }: { children: React.ReactNode; onClick: () => void; className?: string }) => (
    <td className={`cursor-pointer select-none ${className}`} onClick={onClick}>{children}</td>
  )

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
            <p className="text-sm text-gray-500 mt-1">คลิกที่ตัวเลขใดก็ได้เพื่อดูที่มาของค่านั้น</p>
          </div>
          <div className="flex items-center gap-2">
            {rows.length > 0 && (
              <>
                <button onClick={() => setPwDialog({ action: 'save' })} className="px-4 py-2 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 transition-colors">บันทึก</button>
                <button onClick={exportExcel} className="px-4 py-2 bg-green-600 text-white text-sm rounded-lg hover:bg-green-700 transition-colors">↓ Export Excel</button>
                <button
                  onClick={async () => { setLoading(true); await buildPlanRows(parsedCache, selectedProject); setLoading(false) }}
                  className="px-4 py-2 border border-blue-300 text-blue-600 rounded-lg text-sm hover:bg-blue-50 transition-colors"
                  title="โหลดราคา DDP ใหม่จาก Cost Compare">
                  ↻ Refresh Prices
                </button>
              </>
            )}
            <button onClick={() => setShowHistory(true)} className="px-4 py-2 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 transition-colors flex items-center gap-1.5">
              ประวัติ {history.length > 0 && <span className="bg-gray-200 text-gray-600 text-xs rounded-full px-1.5">{history.length}</span>}
            </button>
          </div>
        </div>

        {/* Controls */}
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
                <label className="text-sm text-gray-600 font-medium block mb-1.5">Template Item Code</label>
                <div className="flex items-center gap-2 flex-wrap">
                  {templateInfo ? (
                    <>
                      <span className="text-xs font-semibold text-green-700 bg-green-50 border border-green-200 px-2.5 py-1 rounded-lg">✓ {templateInfo.count} items</span>
                      <span className="text-xs text-gray-400 truncate max-w-[150px]">{templateInfo.fileName}</span>
                      <button onClick={() => templateFileRef.current?.click()} className="text-xs text-blue-500 hover:text-blue-700 whitespace-nowrap">อัพเดต</button>
                      <button onClick={clearTemplate} className="text-gray-300 hover:text-red-400 text-xs">✕</button>
                    </>
                  ) : (
                    <button onClick={() => templateFileRef.current?.click()}
                      className="px-4 py-1.5 border border-dashed border-gray-300 rounded-lg text-sm text-gray-500 hover:bg-gray-50 hover:border-gray-400 transition-colors">
                      เลือก Template
                    </button>
                  )}
                  <input ref={templateFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleTemplateFile} />
                </div>
              </div>
              <div>
                <label className="text-sm text-gray-600 font-medium block mb-1.5">Usage (ย้อนหลัง 3 เดือน)</label>
                <div className="flex items-center gap-2 flex-wrap">
                  {usageData ? (
                    <>
                      <span className="text-xs font-semibold text-purple-700 bg-purple-50 border border-purple-200 px-2.5 py-1 rounded-lg">✓ {usageData.fileName}</span>
                      {(() => {
                        const yr = (usageData.fileName.match(/20\d\d/) ?? [])[0] ?? '2026'
                        const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
                        return (
                          <select
                            value={usageLabel}
                            onChange={e => setUsageLabel(e.target.value)}
                            className="text-xs text-purple-700 bg-purple-50 border border-purple-200 rounded-lg px-2 py-1 cursor-pointer outline-none font-semibold"
                          >
                            {months.map(m => (
                              <option key={m} value={`Usage ${m} ${yr}`}>{m} {yr}</option>
                            ))}
                          </select>
                        )
                      })()}
                      <button onClick={() => usageFileRef.current?.click()} className="text-xs text-blue-500 hover:text-blue-700 whitespace-nowrap">อัพเดต</button>
                      <button onClick={() => setUsageData(null)} className="text-gray-300 hover:text-red-400 text-xs">✕</button>
                    </>
                  ) : (
                    <button onClick={() => usageFileRef.current?.click()}
                      className="px-4 py-1.5 border border-dashed border-gray-300 rounded-lg text-sm text-gray-500 hover:bg-gray-50 hover:border-gray-400 transition-colors">
                      เลือกไฟล์ Usage
                    </button>
                  )}
                  <input ref={usageFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUsageFile} />
                </div>
              </div>
              <div>
                <label className="text-sm text-gray-600 font-medium block mb-1.5">stock_dashboard Excel</label>
                <div className="flex items-center gap-2">
                  <button onClick={() => fileRef.current?.click()} className="px-4 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 transition-colors">เลือกไฟล์</button>
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
                    <button onClick={() => updateFileRefs.current[ss.supplierName]?.click()} className="text-xs text-blue-500 hover:text-blue-700 whitespace-nowrap">อัพเดต</button>
                    <input type="file" accept=".xlsx,.xls" className="hidden"
                      ref={el => { updateFileRefs.current[ss.supplierName] = el }}
                      onChange={e => { const f = e.target.files?.[0]; if (f) parseSupplierStockFile(ss.supplierName, f) }} />
                    <button onClick={() => setSupplierStocks(prev => prev.filter(s => s.supplierName !== ss.supplierName))} className="text-gray-300 hover:text-red-400 text-xs">✕</button>
                  </div>
                ))}
                {addingSupplier ? (
                  <div className="flex items-center gap-2 bg-blue-50 rounded-lg px-3 py-2">
                    <select value={newSupplierName} onChange={e => setNewSupplierName(e.target.value)}
                      className="border border-gray-300 rounded px-2 py-1 text-xs outline-none focus:border-blue-400 bg-white flex-1">
                      <option value="">— เลือก Supplier —</option>
                      {allDbSuppliers.filter(s => !supplierStocks.some(ss => ss.supplierName === s)).map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                    {newSupplierName && (
                      <button onClick={() => { pendingSupplierRef.current = newSupplierName; addSupplierFileRef.current?.click() }}
                        className="text-xs px-3 py-1 bg-blue-600 text-white rounded-lg hover:bg-blue-700 whitespace-nowrap">Upload Stock</button>
                    )}
                    <button onClick={() => { setAddingSupplier(false); setNewSupplierName('') }} className="text-xs text-gray-400 hover:text-gray-600">ยกเลิก</button>
                  </div>
                ) : (
                  <button onClick={() => setAddingSupplier(true)} className="text-xs text-blue-500 hover:text-blue-700 px-3 py-2 border border-dashed border-blue-200 rounded-lg w-full text-left hover:border-blue-400 transition-colors">
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
            <div className="mt-4 border-t border-gray-100 pt-3 flex items-center gap-3 flex-wrap">
              <span className="text-xs text-gray-400">
                {rows.length} items
                {templateInfo ? <span className="text-green-600"> (template {templateInfo.count})</span> : ''}
                {selectedProject ? ` · DDP จาก ${selectedProject}` : ''}
                {hasStock ? ` · Stock: ${supplierStocks.map(s => s.supplierName).join(', ')}` : ''}
              </span>
              {ddpSuppliers.length > 0 && (
                <div className="flex items-center gap-1.5 flex-wrap">
                  <span className="text-xs text-gray-400 mr-1">DDP Ranking:</span>
                  {ddpSuppliers.slice(0, 8).map(s => {
                    const col = supplierColor(s)
                    const currencies = supplierAvailableCurrencies[s] ?? []
                    const pref = supplierCurrencyPref[s] ?? ''
                    const excluded = ddpExcluded.has(s)
                    return (
                      <span key={s} className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs transition-opacity ${excluded ? 'opacity-40 line-through' : ''} ${col.bg} ${col.text}`}>
                        <button
                          onClick={() => setDdpExcluded(prev => { const n = new Set(prev); excluded ? n.delete(s) : n.add(s); return n })}
                          title={excluded ? `เพิ่ม ${s} เข้า DDP ranking` : `ตัด ${s} ออกจาก DDP ranking`}
                          className="w-3 h-3 rounded-full border border-current/50 flex items-center justify-center text-[9px] shrink-0 hover:bg-current/20">
                          {excluded ? '+' : '✕'}
                        </button>
                        {s}
                        {!excluded && currencies.length > 1 && (
                          <select
                            value={pref}
                            onChange={e => setSupplierCurrencyPref(prev => ({ ...prev, [s]: e.target.value }))}
                            onClick={e => e.stopPropagation()}
                            className="ml-0.5 bg-transparent border border-current/30 rounded text-xs outline-none cursor-pointer py-px"
                          >
                            <option value="">latest</option>
                            {currencies.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        )}
                      </span>
                    )
                  })}
                  {hasThaiCost && (
                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs bg-teal-100 text-teal-700">
                      <span className="w-1.5 h-1.5 rounded-full bg-teal-500" />
                      {THAI_COST} (THB)
                    </span>
                  )}
                </div>
              )}
            </div>
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
                    <th className="px-3 py-2.5 text-left whitespace-nowrap font-semibold sticky left-0 bg-gray-50 z-20 border-r border-gray-200 min-w-[200px]">Item Code</th>
                    <th className="px-3 py-2.5 text-left whitespace-nowrap font-semibold sticky left-[200px] bg-gray-50 z-20 border-r border-gray-200 min-w-[220px]">Description</th>
                    {hasThaiCost && <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-teal-50 text-teal-700 border-r border-teal-100">ทุนไทย<br/><span className="font-normal text-xs">(THB)</span></th>}
                    {Array.from({ length: ddpCols }, (_, i) => <th key={i} className="px-2 py-2.5 text-left whitespace-nowrap font-semibold text-gray-500 w-44">DDP {i + 1}</th>)}
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">PO ไทย</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">Stock ไทย</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">PO ไทย/2</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">ลงเรือ</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">Fc. W1</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">Fc. W2</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">Fc. W3+4</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">Fc. Next Month</th>
                    {usageData && <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-purple-50 text-purple-700">{usageLabel}</th>}
                    {usageData && <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-purple-50 text-purple-700">Avg. Usage 3M</th>}
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-green-50 text-green-700">เหลือให้ W3W4</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-green-50 text-green-700">เหลือให้ Next</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-red-50 text-red-700">ต้องสั่งโหลด</th>
                    {supplierStocks.map(ss => {
                      const dates = supplierSlotDates[ss.supplierName] ?? []
                      return (<>
                        <th key={`h-${ss.supplierName}-stk`} className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-indigo-50 text-indigo-700 border-l border-indigo-100">Stock {ss.supplierName}</th>
                        {dates.map((date, i) => (
                          <th key={`h-${ss.supplierName}-d${i}`} className="px-2 py-2 text-center whitespace-nowrap bg-indigo-50 text-indigo-600">
                            <div className="flex items-center gap-1 justify-center">
                              <button className="text-xs text-indigo-700 hover:underline font-semibold" onClick={() => openSlotDateDialog(ss.supplierName, 'edit', i)}>{date}</button>
                              <button className="text-gray-300 hover:text-red-400 text-xs leading-none" onClick={() => openSlotDateDialog(ss.supplierName, 'delete', i)}>✕</button>
                            </div>
                          </th>
                        ))}
                        <th key={`h-${ss.supplierName}-add`} className="px-2 py-2 bg-indigo-50">
                          <button onClick={() => openSlotDateDialog(ss.supplierName, 'add', -1)}
                            className="text-indigo-400 hover:text-indigo-600 text-base font-bold leading-none px-1">+</button>
                        </th>
                        <th key={`h-${ss.supplierName}-rem`} className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-indigo-50 text-indigo-700">คงเหลือ</th>
                      </>)
                    })}
                    {hasStock && <>
                      <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-violet-50 text-violet-700 border-l border-violet-100">รวมโหลด</th>
                      <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-violet-50 text-violet-700">Stock หลังโหลด</th>
                      <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-orange-50 text-orange-700">PO Coverage for<br/>Fc. Next Month</th>
                      <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">
                        <input type="number" min={1} max={200} value={bufferPct} onChange={e => setBufferPct(Number(e.target.value) || 30)}
                          className="w-10 text-center bg-transparent border-b border-amber-400 outline-none font-bold text-amber-700 [appearance:textfield]" />
                        % Buffer Stock<br/><span className="font-normal text-xs">(Based on Fc. Next Month)</span>
                      </th>
                      <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">รวมที่ต้องสั่ง<br/><span className="font-normal text-xs">(Buffer + ขาด)</span></th>
                    </>}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row, i) => {
                    const V = computeV(row.item_code)
                    const W = computeW(row)
                    const Z = computeZ(row)
                    const wNeg = W < 0
                    const zNeg = Z < 0
                    const fp = (colType: string, sup?: string) => () => setFormulaPopup(buildFormula(row, colType, sup))

                    return (
                      <tr key={i} className="border-b border-gray-100 hover:bg-blue-50/10 group">
                        <td className="px-3 py-2 font-mono text-gray-800 whitespace-nowrap sticky left-0 bg-white border-r border-gray-200 z-10 min-w-[200px]">{row.item_code}</td>
                        <td className="px-3 py-2 text-gray-600 sticky left-[200px] bg-white border-r border-gray-200 z-10 min-w-[220px] max-w-[320px]">
                          <div className="overflow-x-auto whitespace-nowrap scrollbar-thin scrollbar-thumb-gray-200 hover:scrollbar-thumb-gray-400 cursor-default" title={row.description}>
                            {row.description || <span className="text-gray-400">—</span>}
                          </div>
                        </td>

                        {hasThaiCost && (() => {
                          const price = thaiCostMap[row.item_code]
                          const ddp1 = row.ddp_prices[0]?.ddp_thb ?? null
                          const isCheaper = price != null && ddp1 != null && price < ddp1
                          const savePct = isCheaper ? Math.round((1 - price / ddp1) * 100) : 0
                          return (
                            <td className={`px-3 py-2 text-right whitespace-nowrap border-r text-xs font-medium ${isCheaper ? 'bg-green-100 border-green-200' : 'bg-teal-50/30 border-teal-100 text-teal-700'}`}>
                              {price != null ? (
                                <div className="flex flex-col items-end gap-0.5">
                                  <span className={`font-mono ${isCheaper ? 'font-bold text-green-800' : ''}`}>
                                    {price.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </span>
                                  {isCheaper && (
                                    <span className="text-[10px] bg-green-600 text-white px-1.5 py-px rounded font-semibold leading-tight">
                                      ถูกกว่า {savePct}%
                                    </span>
                                  )}
                                </div>
                              ) : <span className="text-gray-300">—</span>}
                            </td>
                          )
                        })()}

                        {Array.from({ length: ddpCols }, (_, j) => {
                          const p = row.ddp_prices[j]
                          const col = p ? supplierColor(p.supplier) : null
                          return (
                            <td key={j} className={`px-2 py-1.5 w-44 cursor-pointer ${col ? col.hover : ''}`} onClick={() => p && showDdpFormula(row, p)}>
                              {p ? (
                                <span className={`flex items-center gap-1.5 px-2 py-0.5 rounded-md ${col!.bg} ${col!.text}`}>
                                  <span className={`w-1.5 h-1.5 rounded-full shrink-0 ${col!.dot}`} />
                                  <span className="font-medium text-xs flex-1">{p.supplier}</span>
                                  <span className="font-mono font-semibold text-xs">{fmtDdp(p.ddp_thb)}</span>
                                </span>
                              ) : <span className="text-gray-300 text-xs px-2">—</span>}
                            </td>
                          )
                        })}

                        <C className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap hover:bg-amber-100/40" onClick={fp('po_thai')}>{fmtN(row.po_thai)}</C>
                        <C className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap hover:bg-amber-100/40" onClick={fp('stock_thai')}>{fmtN(row.stock_thai)}</C>
                        <C className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap hover:bg-amber-100/40" onClick={fp('L')}>{fmtF(row.L)}</C>
                        <C className="px-3 py-2 text-right text-gray-700 whitespace-nowrap hover:bg-gray-100/80" onClick={fp('lonsua')}>{fmtN(row.lonsua)}</C>
                        <C className="px-3 py-2 text-right text-gray-700 whitespace-nowrap hover:bg-gray-100/80" onClick={fp('week1')}>{fmtN(row.week1)}</C>
                        <C className="px-3 py-2 text-right text-gray-700 whitespace-nowrap hover:bg-gray-100/80" onClick={fp('week2')}>{fmtN(row.week2)}</C>
                        <C className="px-3 py-2 text-right text-gray-700 whitespace-nowrap hover:bg-gray-100/80" onClick={fp('week3_4')}>{fmtN(row.week3_4)}</C>
                        <C className="px-3 py-2 text-right text-gray-700 whitespace-nowrap hover:bg-gray-100/80" onClick={fp('next_month')}>{fmtN(row.next_month)}</C>
                        {usageData && (() => {
                          const vals = usageData.items[row.item_code]
                          const last = vals ? (vals[2] ?? null) : null
                          const avg = vals ? vals.reduce((a, b) => a + b, 0) / vals.length : null
                          const openAvgPopup = vals ? () => {
                            const lbs = usageData.labels
                            setFormulaPopup({
                              itemCode: row.item_code,
                              description: row.description,
                              colName: 'Avg. Usage 3M',
                              source: `(${lbs[0]} + ${lbs[1]} + ${lbs[2]}) ÷ 3`,
                              lines: [
                                { op: '',  label: lbs[0] || 'เดือน 1', val: vals[0] },
                                { op: '+', label: lbs[1] || 'เดือน 2', val: vals[1] },
                                { op: '+', label: lbs[2] || 'เดือน 3', val: vals[2] },
                                { op: '=', label: 'Avg. Usage 3M', val: avg!, isResult: true },
                              ],
                            })
                          } : undefined
                          return (<>
                            <td className="px-3 py-2 text-right whitespace-nowrap bg-purple-50/20 text-purple-700 font-medium">
                              {last !== null ? last.toLocaleString() : <span className="text-gray-300">—</span>}
                            </td>
                            <td className={`px-3 py-2 text-right whitespace-nowrap bg-purple-50/20 text-purple-700 font-medium ${openAvgPopup ? 'cursor-pointer hover:bg-purple-100/40' : ''}`}
                                onClick={openAvgPopup}>
                              {avg !== null ? fmtF(avg) : <span className="text-gray-300">—</span>}
                            </td>
                          </>)
                        })()}
                        <C className={`px-3 py-2 text-right font-semibold whitespace-nowrap hover:bg-red-100/40 ${row.S < 0 ? 'bg-red-50 text-red-600' : 'bg-green-50/20 text-gray-700 hover:bg-green-100/40'}`} onClick={fp('S')}>
                          {row.S < 0 && <span className="mr-1 text-red-400">⚠</span>}{fmtF(row.S)}
                        </C>
                        <C className={`px-3 py-2 text-right font-semibold whitespace-nowrap ${row.T < 0 ? 'bg-red-50 text-red-600 hover:bg-red-100/40' : 'bg-green-50/20 text-gray-700 hover:bg-green-100/40'}`} onClick={fp('T')}>
                          {row.T < 0 && <span className="mr-1 text-red-400">⚠</span>}{fmtF(row.T)}
                        </C>
                        <C className={`px-3 py-2 text-right font-semibold bg-red-50/20 whitespace-nowrap hover:bg-red-100/40 ${row.U < 0 ? 'text-red-600' : 'text-gray-700'}`} onClick={fp('U')}>{fmtF(row.U)}</C>

                        {supplierStocks.map(ss => {
                          const stockItem = getStockItem(ss.supplierName, row.item_code)
                          const stockQty = stockItem?.total ?? 0
                          const finished = stockItem?.finished ?? 0
                          const dates = supplierSlotDates[ss.supplierName] ?? []
                          const plan = loadingPlan[ss.supplierName]?.[row.item_code] ?? []
                          const remaining = getRemaining(ss.supplierName, row.item_code)
                          const totalLoad = getLoadTotal(ss.supplierName, row.item_code)
                          const overFinished = finished > 0 && totalLoad > finished

                          return (<>
                            <td key={`${ss.supplierName}-stk`} className="px-3 py-2 text-right bg-indigo-50/20 whitespace-nowrap border-l border-indigo-50">
                              {stockQty > 0 ? (
                                <button onClick={() => stockItem && setStatusPopup({ itemCode: row.item_code, description: row.description, supplierName: ss.supplierName, stockItem })}
                                  className="text-indigo-600 hover:underline font-medium">{stockQty.toLocaleString()}</button>
                              ) : <span className="text-gray-300">—</span>}
                            </td>
                            {dates.map((_, i) => (
                              <td key={`${ss.supplierName}-d${i}`} className={`px-1.5 py-1 ${overFinished ? 'bg-orange-50/40' : 'bg-indigo-50/10'}`}>
                                <input type="number" min="0" value={plan[i] ?? ''} onChange={e => setLoadQty(ss.supplierName, row.item_code, i, e.target.value)} placeholder="0"
                                  className={`border rounded px-1.5 py-1 w-20 text-right text-xs outline-none focus:border-indigo-400 ${overFinished ? 'border-orange-400 bg-orange-50' : 'border-gray-200'}`} />
                              </td>
                            ))}
                            <td key={`${ss.supplierName}-add-placeholder`} className="px-1.5 py-1 bg-indigo-50/5 w-8">
                              {overFinished && (
                                <span title={`รวมโหลด ${totalLoad.toLocaleString()} เกินผลิตเสร็จ ${finished.toLocaleString()} ชิ้น`}
                                  className="text-orange-500 text-xs font-bold cursor-help">⚠</span>
                              )}
                            </td>
                            <C key={`${ss.supplierName}-rem`} className={`px-3 py-2 text-right whitespace-nowrap bg-indigo-50/20 font-medium hover:bg-indigo-100/40 ${stockQty > 0 && remaining < 0 ? 'text-red-500' : 'text-gray-700'}`}
                              onClick={fp('remaining', ss.supplierName)}>
                              {stockQty > 0 ? fmtF(remaining) : <span className="text-gray-400">—</span>}
                            </C>
                          </>)
                        })}

                        {hasStock && (() => {
                          const buf = row.next_month * bufferPct / 100
                          const bufNeeded = buf - Z
                          return (<>
                            <C className="px-3 py-2 text-right whitespace-nowrap bg-violet-50/20 font-medium text-gray-700 border-l border-violet-50 hover:bg-violet-100/40" onClick={fp('V')}>
                              {V > 0 ? V.toLocaleString() : <span className="text-gray-400">—</span>}
                            </C>
                            <C className={`px-3 py-2 text-right whitespace-nowrap bg-violet-50/20 font-semibold hover:bg-violet-100/40 ${wNeg ? 'text-red-600' : 'text-gray-700'}`} onClick={fp('W')}>
                              {wNeg && <span className="mr-1">⚠</span>}{fmtF(W)}
                            </C>
                            <C className="px-3 py-2 text-right whitespace-nowrap bg-orange-50/20 font-semibold hover:bg-orange-100/40" onClick={fp('Z')}>
                              {zNeg ? <span className="text-orange-600">ควรสั่ง {Math.ceil(Math.abs(Z)).toLocaleString()}</span>
                                : <span className="text-green-600 font-normal">พอ +{fmtF(Z)}</span>}
                            </C>
                            <C className="px-3 py-2 text-right whitespace-nowrap bg-amber-50/20 text-gray-600 text-sm hover:bg-amber-100/40" onClick={fp('buf')}>
                              {fmtF(buf)}
                            </C>
                            <C className={`px-3 py-2 text-right whitespace-nowrap bg-amber-50/20 font-semibold text-sm hover:bg-amber-100/40 ${bufNeeded > 0 ? 'text-amber-700' : 'text-gray-300'}`} onClick={fp('bufNeeded')}>
                              {bufNeeded > 0 ? `Buffer Qty ${Math.ceil(bufNeeded).toLocaleString()}` : '—'}
                            </C>
                          </>)
                        })()}
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* ── Formula popup ──────────────────────────────────────────────────── */}
      {formulaPopup && (
        <div className="fixed inset-0 bg-black/20 z-50 flex items-center justify-center p-4" onClick={() => setFormulaPopup(null)}>
          <div className="bg-white rounded-2xl shadow-2xl p-6 max-w-sm w-full" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-start mb-1">
              <div>
                <p className="text-xs text-gray-400 font-mono">{formulaPopup.itemCode}</p>
                <h3 className="text-base font-bold text-gray-900 mt-0.5">{formulaPopup.colName}</h3>
              </div>
              <button onClick={() => setFormulaPopup(null)} className="text-gray-300 hover:text-gray-600 text-lg leading-none mt-1">✕</button>
            </div>

            {formulaPopup.source && (
              <p className="text-xs text-blue-500 mb-3 bg-blue-50 rounded-lg px-3 py-1.5">{formulaPopup.source}</p>
            )}
            {formulaPopup.formulaStr && (
              <p className="text-xs text-gray-600 mb-3 font-mono bg-gray-50 rounded-lg px-3 py-1.5">{formulaPopup.formulaStr}</p>
            )}

            <div className="space-y-1 mt-2">
              {formulaPopup.lines.map((line, idx) => (
                <div key={idx}>
                  {line.isResult && <div className="border-t-2 border-gray-300 my-2" />}
                  <div className={`flex items-baseline gap-2 px-1 py-0.5 rounded ${line.isResult ? 'bg-gray-50' : ''}`}>
                    <span className="text-xs w-4 text-right shrink-0 text-gray-400 font-mono">{line.op}</span>
                    <span className={`flex-1 text-sm ${line.isResult ? 'font-bold text-gray-900' : 'text-gray-600'}`}>{line.label}</span>
                    <span className={`text-sm font-mono tabular-nums shrink-0 ${line.isResult ? 'font-bold text-gray-900' : 'text-gray-700'} ${typeof line.val === 'number' && line.val < 0 ? 'text-red-600' : ''}`}>
                      {typeof line.val === 'number' ? fmtVal(line.val, line.isResult ? 1 : 2) : line.val}
                    </span>
                  </div>
                  {line.note && (
                    <p className={`text-xs mt-1 px-5 font-medium ${typeof line.val === 'number' && line.val < 0 ? 'text-orange-600' : 'text-green-600'}`}>
                      → {line.note}
                    </p>
                  )}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ── Status popup ───────────────────────────────────────────────────── */}
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
                <div className="space-y-2 max-h-52 overflow-y-auto">
                  {statusPopup.stockItem.pos.map((po, i) => (
                    <div key={i} className="py-1.5 border-b border-gray-50 last:border-0">
                      <div className="flex justify-between items-baseline mb-1">
                        <span className="font-mono text-xs font-semibold text-gray-700">{po.poNo}</span>
                        <span className="text-xs text-gray-500">{(po.total || po.notProduced + po.inProduction + po.finished + po.onBoard).toLocaleString()} pcs</span>
                      </div>
                      <div className="flex gap-1.5 flex-wrap">
                        {po.notProduced > 0 && <span className="text-xs px-2 py-0.5 bg-gray-100 text-gray-600 rounded-md">ยังไม่ผลิต {po.notProduced.toLocaleString()}</span>}
                        {po.inProduction > 0 && <span className="text-xs px-2 py-0.5 bg-yellow-50 text-yellow-700 rounded-md">กำลังผลิต {po.inProduction.toLocaleString()}</span>}
                        {po.finished > 0 && <span className="text-xs px-2 py-0.5 bg-green-50 text-green-700 rounded-md">ผลิตเสร็จ {po.finished.toLocaleString()}</span>}
                        {po.onBoard > 0 && <span className="text-xs px-2 py-0.5 bg-blue-50 text-blue-700 rounded-md">On Board {po.onBoard.toLocaleString()}</span>}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── Password dialog ────────────────────────────────────────────────── */}
      {pwDialog && (
        <div className="fixed inset-0 bg-black/30 z-50 flex items-center justify-center p-4" onClick={() => { setPwDialog(null); setPwInput(''); setPwError(false) }}>
          <div className="bg-white rounded-xl shadow-xl p-6 max-w-xs w-full" onClick={e => e.stopPropagation()}>
            <h3 className="font-bold text-gray-900 mb-1">
              {pwDialog.action === 'save' ? 'บันทึกแผน' : 'ลบประวัติ'}
            </h3>
            <p className="text-xs text-gray-400 mb-4">ใส่รหัสผ่านเพื่อยืนยัน</p>
            <input
              type="password"
              value={pwInput}
              onChange={e => { setPwInput(e.target.value); setPwError(false) }}
              onKeyDown={e => e.key === 'Enter' && confirmPasswordAction()}
              placeholder="รหัสผ่าน"
              autoFocus
              className="border border-gray-300 rounded-lg px-3 py-2 w-full text-sm outline-none focus:border-blue-400 mb-1" />
            {pwError && <p className="text-xs text-red-500 mb-3">รหัสผ่านไม่ถูกต้อง</p>}
            {!pwError && <div className="mb-3" />}
            <div className="flex gap-2">
              <button onClick={() => { setPwDialog(null); setPwInput(''); setPwError(false) }}
                className="flex-1 px-4 py-2 border border-gray-300 rounded-lg text-sm text-gray-600 hover:bg-gray-50">ยกเลิก</button>
              <button onClick={confirmPasswordAction}
                className={`flex-1 px-4 py-2 text-white text-sm rounded-lg transition-colors ${pwDialog.action === 'delete' ? 'bg-red-500 hover:bg-red-600' : 'bg-blue-600 hover:bg-blue-700'}`}>
                {pwDialog.action === 'save' ? 'บันทึก' : 'ลบ'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Slot date dialog ───────────────────────────────────────────────── */}
      {slotDateDialog && (
        <div className="fixed inset-0 bg-black/30 z-50 flex items-center justify-center p-4" onClick={() => setSlotDateDialog(null)}>
          <div className="bg-white rounded-xl shadow-xl p-6 max-w-xs w-full" onClick={e => e.stopPropagation()}>
            {slotDateDialog.step === 'pw' ? (<>
              <h3 className="font-bold text-gray-900 mb-1">
                {slotDateDialog.action === 'add' ? 'เพิ่มวันโหลด' : slotDateDialog.action === 'edit' ? 'แก้ไขวันโหลด' : 'ลบวันโหลด'}
              </h3>
              <p className="text-xs text-gray-400 mb-4">ใส่รหัสผ่านเพื่อยืนยัน</p>
              <input type="password" value={slotDateDialog.pw} autoFocus placeholder="รหัสผ่าน"
                onChange={e => setSlotDateDialog(d => d ? { ...d, pw: e.target.value, pwError: false } : null)}
                onKeyDown={e => e.key === 'Enter' && confirmSlotDatePw()}
                className="border border-gray-300 rounded-lg px-3 py-2 w-full text-sm outline-none focus:border-blue-400 mb-1" />
              {slotDateDialog.pwError && <p className="text-xs text-red-500 mb-3">รหัสผ่านไม่ถูกต้อง</p>}
              {!slotDateDialog.pwError && <div className="mb-3" />}
              <div className="flex gap-2">
                <button onClick={() => setSlotDateDialog(null)} className="flex-1 px-4 py-2 border border-gray-300 rounded-lg text-sm text-gray-600 hover:bg-gray-50">ยกเลิก</button>
                <button onClick={confirmSlotDatePw} className={`flex-1 px-4 py-2 text-white text-sm rounded-lg ${slotDateDialog.action === 'delete' ? 'bg-red-500 hover:bg-red-600' : 'bg-blue-600 hover:bg-blue-700'}`}>ยืนยัน</button>
              </div>
            </>) : (<>
              <h3 className="font-bold text-gray-900 mb-1">
                {slotDateDialog.action === 'add' ? `เพิ่มวันโหลด — ${slotDateDialog.supplierName}` : `แก้ไขวันโหลด — ${slotDateDialog.supplierName}`}
              </h3>
              <p className="text-xs text-gray-400 mb-3">เช่น 8 Aug 2026</p>
              <input type="text" value={slotDateDialog.dateInput} autoFocus placeholder="8 Aug 2026"
                onChange={e => setSlotDateDialog(d => d ? { ...d, dateInput: e.target.value } : null)}
                onKeyDown={e => e.key === 'Enter' && confirmSlotDateAction()}
                className="border border-gray-300 rounded-lg px-3 py-2 w-full text-sm outline-none focus:border-blue-400 mb-4" />
              <div className="flex gap-2">
                <button onClick={() => setSlotDateDialog(null)} className="flex-1 px-4 py-2 border border-gray-300 rounded-lg text-sm text-gray-600 hover:bg-gray-50">ยกเลิก</button>
                <button onClick={confirmSlotDateAction} className="flex-1 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm rounded-lg">บันทึก</button>
              </div>
            </>)}
          </div>
        </div>
      )}

      {/* ── History modal ──────────────────────────────────────────────────── */}
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
                        {' · '}{s.itemCount} items{s.selectedProject ? ` · ${s.selectedProject}` : ''}
                      </p>
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <button onClick={() => restoreSession(s)} className="text-xs px-3 py-1.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors">โหลด</button>
                      <button onClick={() => setPwDialog({ action: 'delete', id: s.id })} className="text-xs text-gray-300 hover:text-red-400 transition-colors">✕</button>
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

'use client'

import { useState, useEffect, useRef, useMemo, useCallback, Fragment } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import NavBar from '@/components/NavBar'
import { isUnlocked } from '@/lib/auth'

interface LoadItem {
  item_code: string
  description: string
  qty_1: number
  qty_2: number
}

interface FLine {
  op: '' | '+' | '−' | '×' | '÷' | '='
  label: string
  val: number | string
  isResult?: boolean
}

interface FormulaPopup {
  itemCode: string
  description: string
  colName: string
  formulaStr?: string
  lines: FLine[]
}

interface LoadEntry {
  date: string
  item_code: string
  qty: number
  po_id?: string   // which PO this belongs to (undefined = legacy supplier-level)
  is_auto?: boolean  // true = system-suggested; false/undefined = manually entered
}

interface LoadSupplier {
  id: string
  name: string
  po_ids: string[]
  container_qtys: Record<string, number>
  loads: LoadEntry[]
  load_dates: string[]           // legacy — kept for backward compat
  po_load_dates: Record<string, string[]>  // po_id -> sorted dates (new)
}

interface LoadPlan {
  id?: string
  name: string
  type_1_name: string
  type_2_name: string
  forecast_1: number
  forecast_2: number
  items: LoadItem[]
  suppliers: LoadSupplier[]
  high_ratio_items?: string[]
  rate_default?: number
  rate_high?: number
  rate_rules?: { keyword: string; rate: 'high' | 'standard' }[]  // saved bulk rules
  updated_at?: string
}

interface POUpload {
  id: string
  supplier: string
  project: string
  po_rbs_ch_no: string | null
  filename: string | null
  created_at: string
  rows: { item_code: string; qty: number }[] | null
}

function mkPlan(): LoadPlan {
  return { name: 'Untitled Plan', type_1_name: 'Existing', type_2_name: 'New', forecast_1: 0, forecast_2: 0, items: [], suppliers: [], high_ratio_items: [], rate_default: 70, rate_high: 90, rate_rules: [] }
}

// Re-derive high_ratio_items by replaying all rules in order
function applyRateRules(items: LoadItem[], rules: { keyword: string; rate: 'high' | 'standard' }[]): string[] {
  const hi = new Set<string>()
  for (const rule of rules) {
    const kw = rule.keyword.toLowerCase()
    for (const item of items) {
      if (item.item_code.toLowerCase().includes(kw) || item.description.toLowerCase().includes(kw)) {
        if (rule.rate === 'high') hi.add(item.item_code)
        else hi.delete(item.item_code)
      }
    }
  }
  return Array.from(hi)
}

function mkSupplier(): LoadSupplier {
  return { id: crypto.randomUUID(), name: '', po_ids: [], container_qtys: {}, loads: [], load_dates: [], po_load_dates: {} }
}

// Ensure all optional fields are initialized when loading from Supabase
function normalizePlan<T extends LoadPlan>(raw: T): T {
  return {
    ...raw,
    high_ratio_items: raw.high_ratio_items ?? [],
    rate_rules: raw.rate_rules ?? [],
    rate_default: raw.rate_default ?? 70,
    rate_high: raw.rate_high ?? 90,
    suppliers: (raw.suppliers ?? []).map(s => ({
      ...s,
      po_load_dates: (s as any).po_load_dates ?? {},
      load_dates: s.load_dates ?? [],
      loads: s.loads ?? [],
      po_ids: s.po_ids ?? [],
      container_qtys: s.container_qtys ?? {},
    })),
  }
}

function fmtDate(d: string) {
  return new Date(d + 'T00:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
}

function poLabel(po: POUpload | undefined, fallback: string): string {
  if (!po) return fallback.slice(0, 8)
  return po.po_rbs_ch_no || po.filename?.replace(/\.xlsx?$/i, '') || po.id.slice(0, 8)
}

// ── Supplier row component ──
function SupplierRow({ sup, editable, poUploads, showPoSelector, onUpdate, onUploadContainerQty, onTogglePoSelector, onRemove, onAddPoDate, onRemovePoDate }: {
  sup: LoadSupplier
  editable: boolean
  poUploads: POUpload[]
  showPoSelector: boolean
  onUpdate: (upd: Partial<LoadSupplier>) => void
  onUploadContainerQty: (f: File) => void
  onTogglePoSelector: () => void
  onRemove: () => void
  onAddPoDate: (poId: string, date: string) => void
  onRemovePoDate: (poId: string, date: string) => void
}) {
  const fileRef = useRef<HTMLInputElement>(null)
  const [newDate, setNewDate] = useState('')
  const [newDatePoId, setNewDatePoId] = useState('')
  const containerCount = Object.keys(sup.container_qtys).length
  const dis = !editable

  const selectedPos = sup.po_ids
    .map(pid => poUploads.find(p => p.id === pid))
    .filter(Boolean) as POUpload[]

  // All PO-date entries regardless of whether PO is loaded (prevents display loss)
  const poDateEntries = Object.entries(sup.po_load_dates ?? {}).filter(([, dates]) => dates.length > 0)

  return (
    <div className="rounded-xl border border-gray-200 bg-white p-3">
      <div className="flex items-center gap-2 flex-wrap">
        <input
          value={sup.name}
          onChange={e => onUpdate({ name: e.target.value })}
          placeholder="Supplier name"
          disabled={dis}
          className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm w-36 focus:outline-none focus:border-blue-400 bg-white disabled:bg-gray-50 disabled:text-gray-400"
        />

        {/* PO selector */}
        <div className="relative">
          <button
            onClick={dis ? undefined : onTogglePoSelector}
            disabled={dis}
            className="px-2 py-1.5 border border-gray-300 rounded-lg text-xs text-gray-700 hover:bg-gray-50 bg-white font-medium disabled:opacity-40"
          >
            POs ({sup.po_ids.length}) ▾
          </button>
          {showPoSelector && !dis && (
            <div className="absolute left-0 top-full mt-1 z-50 bg-white border border-gray-200 rounded-xl shadow-xl w-80 max-h-60 overflow-y-auto">
              <div className="sticky top-0 bg-gray-50 px-3 py-2 text-xs text-gray-500 border-b font-semibold">
                {sup.name ? `POs ของ ${sup.name}` : 'Select POs'}
              </div>
              {(() => {
                const filtered = sup.name.trim()
                  ? poUploads.filter(po => po.supplier.toLowerCase().includes(sup.name.trim().toLowerCase()))
                  : poUploads
                return filtered.length === 0
                  ? <div className="p-3 text-xs text-gray-400">{sup.name ? `ไม่พบ PO ของ "${sup.name}"` : 'No POs available'}</div>
                  : filtered.map(po => (
                    <label key={po.id} className="flex items-center gap-2 px-3 py-2 hover:bg-blue-50 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={sup.po_ids.includes(po.id)}
                        onChange={e => onUpdate({ po_ids: e.target.checked ? [...sup.po_ids, po.id] : sup.po_ids.filter(id => id !== po.id) })}
                      />
                      <span className="text-xs leading-tight">
                        <span className="font-mono font-semibold">{po.po_rbs_ch_no || po.filename?.replace(/\.xlsx?$/i, '') || po.id.slice(0, 8)}</span>
                        <span className="text-gray-400 ml-1">· {po.supplier}</span>
                      </span>
                    </label>
                  ))
              })()}
            </div>
          )}
        </div>

        {/* Container QTY upload — optional */}
        <button
          onClick={() => !dis && fileRef.current?.click()}
          disabled={dis}
          title="Excel: A = item_code · B = container qty (optional)"
          className="px-2 py-1.5 border border-gray-300 rounded-lg text-xs text-gray-700 hover:bg-gray-50 bg-white font-medium disabled:opacity-40"
        >
          {containerCount > 0 ? `Container QTY (${containerCount})` : 'Container QTY (optional)'}
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) onUploadContainerQty(f); e.target.value = '' }} />

        {/* Add load date per PO */}
        {selectedPos.length > 0 ? (
          <>
            <select
              value={newDatePoId}
              onChange={e => setNewDatePoId(e.target.value)}
              disabled={dis}
              className="border border-gray-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none bg-white disabled:bg-gray-50 disabled:opacity-40 max-w-[140px]"
            >
              <option value="">เลือก PO</option>
              {selectedPos.map(po => (
                <option key={po.id} value={po.id}>
                  {poLabel(po, po.id)}
                </option>
              ))}
            </select>
            <input type="date" value={newDate} onChange={e => setNewDate(e.target.value)}
              disabled={dis}
              className="border border-gray-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none bg-white disabled:bg-gray-50 disabled:opacity-40" />
            <button
              onClick={() => { if (newDate && newDatePoId) { onAddPoDate(newDatePoId, newDate); setNewDate('') } }}
              disabled={dis || !newDate || !newDatePoId}
              className="px-2 py-1.5 bg-gray-100 hover:bg-gray-200 text-xs rounded-lg font-medium disabled:opacity-40"
            >
              + Date
            </button>
          </>
        ) : (
          <span className="text-xs text-gray-400 italic">เลือก PO ก่อนเพื่อเพิ่มวันโหลด</span>
        )}

        <button onClick={dis ? undefined : onRemove} disabled={dis} className="text-gray-400 hover:text-red-500 text-lg leading-none px-1 ml-auto disabled:opacity-30">×</button>
      </div>

      {/* Dates grouped by PO — iterated from po_load_dates directly (not via selectedPos) */}
      {poDateEntries.length > 0 && (
        <div className="mt-2 space-y-1.5">
          {poDateEntries.map(([poId, dates]) => {
            const po = poUploads.find(p => p.id === poId)
            const label = po ? poLabel(po, poId) : poId.slice(0, 8)
            return (
              <div key={poId} className="flex items-start gap-1.5 flex-wrap">
                <span className="text-[10px] font-mono font-semibold text-blue-600 mt-0.5 shrink-0">{label}:</span>
                {dates.map(d => (
                  <span key={d} className="inline-flex items-center gap-1 bg-white border border-blue-200 rounded-full px-2.5 py-0.5 text-xs text-gray-700">
                    {fmtDate(d)}
                    {!dis && <button onClick={() => onRemovePoDate(poId, d)} className="text-gray-400 hover:text-red-500 ml-0.5">×</button>}
                  </span>
                ))}
              </div>
            )
          })}
        </div>
      )}

      {/* Legacy supplier-level dates (backward compat) */}
      {sup.load_dates.length > 0 && (
        <div className="flex gap-1.5 flex-wrap mt-2">
          {sup.load_dates.map(d => (
            <span key={d} className="inline-flex items-center gap-1 bg-gray-100 border border-gray-200 rounded-full px-2.5 py-0.5 text-xs text-gray-500">
              {fmtDate(d)}
              {!dis && <button
                onClick={() => onUpdate({ load_dates: sup.load_dates.filter(x => x !== d), loads: sup.loads.filter(l => l.date !== d && !l.po_id) })}
                className="text-gray-400 hover:text-red-500 ml-0.5"
              >×</button>}
            </span>
          ))}
        </div>
      )}
    </div>
  )
}

// ── Main page ──
export default function LoadPlanPage() {
  const [unlocked, setUnlocked] = useState(false)
  const [plans, setPlans] = useState<(LoadPlan & { id: string })[]>([])
  const [plan, setPlan] = useState<LoadPlan | null>(null)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [lockMsg, setLockMsg] = useState(false)
  const [parseInfo, setParseInfo] = useState<{ type: 1 | 2; label: string } | null>(null)
  const [uploadedFiles, setUploadedFiles] = useState<{ type1?: string; type2?: string }>({})
  const [expandedSupIds, setExpandedSupIds] = useState<Set<string>>(new Set())
  const [poUploads, setPoUploads] = useState<POUpload[]>([])
  const [showPoSelector, setShowPoSelector] = useState<string | null>(null)
  const [rateFilter, setRateFilter] = useState('')
  const [formulaPopup, setFormulaPopup] = useState<FormulaPopup | null>(null)
  const templateRef1 = useRef<HTMLInputElement>(null)
  const templateRef2 = useRef<HTMLInputElement>(null)

  // Aggregate PO qty per supplier (sum all POs)
  const allPoQtyMaps = useMemo(() => {
    const result = new Map<string, Map<string, number>>()
    if (!plan) return result
    for (const sup of plan.suppliers) {
      const m = new Map<string, number>()
      for (const poId of sup.po_ids) {
        const po = poUploads.find(p => p.id === poId)
        if (!po?.rows) continue
        for (const r of po.rows) {
          if (r.item_code) m.set(r.item_code, (m.get(r.item_code) ?? 0) + (r.qty ?? 0))
        }
      }
      result.set(sup.id, m)
    }
    return result
  }, [plan, poUploads])

  const totalLoadsMap = useMemo(() => {
    const map = new Map<string, number>()
    if (!plan) return map
    for (const s of plan.suppliers) {
      for (const l of s.loads) map.set(l.item_code, (map.get(l.item_code) ?? 0) + l.qty)
    }
    return map
  }, [plan])

  // Sum of all PO quantities across every supplier per item
  const totalPoQtyMap = useMemo(() => {
    const map = new Map<string, number>()
    for (const [, m] of allPoQtyMaps) {
      for (const [ic, qty] of m) map.set(ic, (map.get(ic) ?? 0) + qty)
    }
    return map
  }, [allPoQtyMaps])

  // All unique load dates across all suppliers/POs (sorted)
  const allUniqueDates = useMemo(() => {
    if (!plan) return []
    const s = new Set<string>()
    for (const sup of plan.suppliers) {
      for (const poId of sup.po_ids) {
        for (const d of (sup.po_load_dates?.[poId] ?? [])) s.add(d)
      }
      for (const d of sup.load_dates) s.add(d)
    }
    return Array.from(s).sort()
  }, [plan])

  // date -> item_code -> total qty loaded on that date across all suppliers
  const dailyTotalsMap = useMemo(() => {
    const map = new Map<string, Map<string, number>>()
    if (!plan) return map
    for (const sup of plan.suppliers) {
      for (const l of sup.loads) {
        if (!l.date) continue
        let dm = map.get(l.date)
        if (!dm) { dm = new Map(); map.set(l.date, dm) }
        dm.set(l.item_code, (dm.get(l.item_code) ?? 0) + l.qty)
      }
    }
    return map
  }, [plan])

  useEffect(() => {
    setUnlocked(isUnlocked())
    supabase.from('load_plans').select('*').order('created_at', { ascending: false }).then(({ data }) => {
      if (data) setPlans((data as any[]).map(normalizePlan) as (LoadPlan & { id: string })[])
    })
    supabase.from('po_uploads').select('id, supplier, project, po_rbs_ch_no, filename, created_at, rows').then(({ data }) => {
      if (data) setPoUploads(data as POUpload[])
    })
  }, [])

  const refreshPlans = useCallback(async () => {
    const { data } = await supabase.from('load_plans').select('*').order('created_at', { ascending: false })
    if (data) setPlans((data as any[]).map(normalizePlan) as (LoadPlan & { id: string })[])
  }, [])

  function showLockMsg() {
    setLockMsg(true)
    setTimeout(() => setLockMsg(false), 3000)
  }

  const deletePlan = useCallback(async (id: string) => {
    if (!isUnlocked()) { showLockMsg(); return }
    await supabase.from('load_plans').delete().eq('id', id)
    refreshPlans()
    if ((plan as any)?.id === id) setPlan(null)
  }, [plan, refreshPlans])

  const savePlan = useCallback(async () => {
    if (!plan) return
    if (!isUnlocked()) { showLockMsg(); return }
    setSaving(true)
    const payload = {
      name: plan.name,
      type_1_name: plan.type_1_name,
      type_2_name: plan.type_2_name,
      forecast_1: plan.forecast_1,
      forecast_2: plan.forecast_2,
      items: plan.items,
      suppliers: plan.suppliers,
      high_ratio_items: plan.high_ratio_items ?? [],
      rate_default: plan.rate_default ?? 70,
      rate_high: plan.rate_high ?? 90,
      rate_rules: plan.rate_rules ?? [],
      updated_at: new Date().toISOString(),
    }
    if ((plan as LoadPlan & { id?: string }).id) {
      await supabase.from('load_plans').update(payload).eq('id', (plan as any).id)
    } else {
      const { data } = await supabase.from('load_plans').insert(payload).select('id').single()
      if (data) setPlan(p => p ? { ...p, id: (data as any).id } : p)
    }
    setSaving(false)
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
    refreshPlans()
  }, [plan, refreshPlans])

  // Parse template file — auto-detects item_code / description / qty columns from header row
  function handleTemplateUpload(file: File, typeNum: 1 | 2) {
    setUploadedFiles(prev => ({ ...prev, [typeNum === 1 ? 'type1' : 'type2']: file.name }))
    file.arrayBuffer().then(buf => {
      const wb = XLSX.read(buf, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][]

      let headerRow = 0
      let itemCol = 1
      let descCol = 2
      let qtyCol  = 3

      {
        let bestRow = -1, bestFi = -1, bestDi = -1, bestQi = -1, foundItem = false
        for (let ri = 0; ri < Math.min(5, rows.length) && !foundItem; ri++) {
          const r = rows[ri] as unknown[]
          let fi = -1, di = -1, qi = -1
          for (let ci = 0; ci < r.length; ci++) {
            const h = String(r[ci] ?? '').toLowerCase().replace(/[\s_\-]/g, '')
            if (fi < 0 && (h.includes('itemno') || h.includes('itemcode') || h === 'item')) fi = ci
            if (di < 0 && (h.includes('desc') || h.includes('name') || h.includes('สินค้า'))) di = ci
            if (qi < 0 && (h.includes('qty') || h.includes('assumed') || h.includes('จำนวน'))) qi = ci
          }
          if (fi >= 0) {
            bestRow = ri; bestFi = fi; bestDi = di; bestQi = qi; foundItem = true
          } else if (qi >= 0 && bestRow < 0) {
            bestRow = ri; bestFi = fi; bestDi = di; bestQi = qi
          }
        }
        if (bestRow >= 0) {
          headerRow = bestRow
          if (bestFi >= 0) itemCol = bestFi
          if (bestDi >= 0) descCol = bestDi
          if (bestQi >= 0) qtyCol = bestQi
        }
      }

      function parseQty(v: unknown): number {
        if (v === null || v === undefined || v === '' || v === '-') return 0
        if (typeof v === 'number') return isNaN(v) ? 0 : v
        const s = String(v).trim().replace(/,/g, '')
        return s === '' || s === '-' ? 0 : (Number(s) || 0)
      }

      {
        const firstDataRows = rows.slice(headerRow + 1, headerRow + 6) as unknown[][]
        const nonZeroAt = (col: number) => firstDataRows.some(r => parseQty(r[col]) > 0)
        if (!nonZeroAt(qtyCol)) {
          for (const delta of [-1, 1, -2, 2, -3, 3]) {
            const tryCol = qtyCol + delta
            if (tryCol >= 0 && tryCol !== itemCol && tryCol !== descCol && nonZeroAt(tryCol)) {
              qtyCol = tryCol
              break
            }
          }
        }
      }

      const parsed = new Map<string, { description: string; qty: number }>()
      for (let i = headerRow + 1; i < rows.length; i++) {
        const r = rows[i] as unknown[]
        const item_code = String(r[itemCol] ?? '').trim()
        if (!item_code) continue
        const qty = parseQty(r[qtyCol])
        parsed.set(item_code, { description: String(r[descCol] ?? '').trim(), qty })
      }

      const firstEntry = parsed.size > 0 ? parsed.entries().next().value : null
      const colLetters = (n: number) => String.fromCharCode(65 + n)
      setParseInfo({
        type: typeNum,
        label: `col ${colLetters(itemCol)}=item · col ${colLetters(descCol)}=desc · col ${colLetters(qtyCol)}=qty` +
          (firstEntry ? ` | "${firstEntry[0]}" qty=${firstEntry[1].qty}` : ` | 0 items`)
      })

      setPlan(p => {
        if (!p) return p
        const existingMap = new Map(p.items.map(it => [it.item_code, it]))
        for (const [ic, { description, qty }] of parsed) {
          const ex = existingMap.get(ic)
          if (ex) {
            existingMap.set(ic, { ...ex, description: description || ex.description, [`qty_${typeNum}`]: qty } as LoadItem)
          } else {
            existingMap.set(ic, { item_code: ic, description, qty_1: typeNum === 1 ? qty : 0, qty_2: typeNum === 2 ? qty : 0 })
          }
        }
        return { ...p, items: Array.from(existingMap.values()) }
      })
    })
  }

  function handleContainerQtyUpload(file: File, suppId: string) {
    file.arrayBuffer().then(buf => {
      const wb = XLSX.read(buf, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][]
      const qtys: Record<string, number> = {}
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i] as unknown[]
        const ic = String(r[0] ?? '').trim()
        if (ic) qtys[ic] = Number(r[1]) || 0
      }
      setPlan(p => p ? { ...p, suppliers: p.suppliers.map(s => s.id === suppId ? { ...s, container_qtys: qtys } : s) } : p)
    })
  }

  function updateSup(id: string, upd: Partial<LoadSupplier>) {
    setPlan(p => p ? { ...p, suppliers: p.suppliers.map(s => s.id === id ? { ...s, ...upd } : s) } : p)
  }

  function toggleExpanded(supId: string) {
    setExpandedSupIds(prev => {
      const next = new Set(prev)
      if (next.has(supId)) next.delete(supId)
      else next.add(supId)
      return next
    })
  }

  function toggleHighRatio(itemCode: string) {
    if (!unlocked) return
    setPlan(p => {
      if (!p) return p
      const hi = p.high_ratio_items ?? []
      const newHi = hi.includes(itemCode) ? hi.filter(c => c !== itemCode) : [...hi, itemCode]
      return { ...p, high_ratio_items: newHi }
    })
  }

  // Auto-distribute a PO's item quantities evenly across its dates (rounded to 10s, remainder to first date)
  function recalcAutoForPo(p: LoadPlan, suppId: string, poId: string): LoadPlan {
    const sup = p.suppliers.find(s => s.id === suppId)
    if (!sup) return p
    const dates = sup.po_load_dates?.[poId] ?? []
    const po = poUploads.find(x => x.id === poId)
    if (!po?.rows || !dates.length) return p

    const withoutAuto = sup.loads.filter(l => !(l.po_id === poId && l.is_auto))
    const newAutoLoads: LoadEntry[] = []
    const n = dates.length

    for (const row of po.rows) {
      const { item_code, qty } = row
      if (!qty) continue
      // Don't override any manual entry for this item+PO
      const hasManual = withoutAuto.some(l => l.po_id === poId && l.item_code === item_code)
      if (hasManual) continue
      // floor each non-first date to nearest 10, first date gets remainder
      const base = Math.floor(Math.floor(qty / n) / 100) * 100
      dates.forEach((d, i) => {
        const autoQty = i === 0 ? qty - base * (n - 1) : base
        if (autoQty > 0) newAutoLoads.push({ po_id: poId, date: d, item_code, qty: autoQty, is_auto: true })
      })
    }
    return { ...p, suppliers: p.suppliers.map(s => s.id === suppId ? { ...s, loads: [...withoutAuto, ...newAutoLoads] } : s) }
  }

  function addPoDate(suppId: string, poId: string, date: string) {
    setPlan(prev => {
      if (!prev) return prev
      const sup = prev.suppliers.find(s => s.id === suppId)
      if (!sup) return prev
      const existing = sup.po_load_dates?.[poId] ?? []
      if (existing.includes(date)) return prev
      const newDates = [...existing, date].sort()
      const withDates = {
        ...prev,
        suppliers: prev.suppliers.map(s => s.id === suppId
          ? { ...s, po_load_dates: { ...(s.po_load_dates ?? {}), [poId]: newDates } }
          : s)
      }
      return recalcAutoForPo(withDates, suppId, poId)
    })
  }

  function removePoDate(suppId: string, poId: string, date: string) {
    setPlan(prev => {
      if (!prev) return prev
      const sup = prev.suppliers.find(s => s.id === suppId)
      if (!sup) return prev
      const newDates = (sup.po_load_dates?.[poId] ?? []).filter(d => d !== date)
      const withRemoved = {
        ...prev,
        suppliers: prev.suppliers.map(s => s.id === suppId
          ? {
              ...s,
              po_load_dates: { ...(s.po_load_dates ?? {}), [poId]: newDates },
              loads: s.loads.filter(l => !(l.po_id === poId && l.date === date)),
            }
          : s)
      }
      return recalcAutoForPo(withRemoved, suppId, poId)
    })
  }

  function updateLoadQty(suppId: string, poId: string, date: string, itemCode: string, qty: number) {
    setPlan(p => {
      if (!p) return p
      const sup = p.suppliers.find(s => s.id === suppId)
      if (!sup) return p
      const idx = sup.loads.findIndex(l => l.po_id === poId && l.date === date && l.item_code === itemCode)
      let newLoads: LoadEntry[]
      if (qty <= 0) {
        newLoads = sup.loads.filter(l => !(l.po_id === poId && l.date === date && l.item_code === itemCode))
      } else if (idx >= 0) {
        newLoads = sup.loads.map((l, i) => i === idx ? { ...l, qty, is_auto: false } : l)
      } else {
        newLoads = [...sup.loads, { po_id: poId, date, item_code: itemCode, qty, is_auto: false }]
      }
      return { ...p, suppliers: p.suppliers.map(s => s.id === suppId ? { ...s, loads: newLoads } : s) }
    })
  }

  function exportToExcel() {
    if (!plan) return
    const wb = XLSX.utils.book_new()
    const hiItems = plan.high_ratio_items ?? []
    const rdPct = plan.rate_default ?? 70
    const rhPct = plan.rate_high ?? 90

    const meta = [
      [`Plan: ${plan.name || 'Untitled Plan'}`],
      [`${plan.type_1_name}: ${plan.forecast_1} branches   ${plan.type_2_name}: ${plan.forecast_2} branches`],
      [`Exported: ${new Date().toLocaleString('en-GB')}`],
      [],
    ]

    const headers: string[] = ['Item Code', 'Description', `Assumed Qty (${plan.type_1_name})`, `Assumed Qty (${plan.type_2_name})`, 'Ratio', 'Sum Assumed Qty']
    for (const sup of plan.suppliers) {
      headers.push(`PO QTY (${sup.name})`)
      for (const poId of sup.po_ids) {
        const po = poUploads.find(p => p.id === poId)
        const label = poLabel(po, poId)
        for (const d of (sup.po_load_dates?.[poId] ?? [])) {
          headers.push(`Load ${d} [${label}] (${sup.name})`)
        }
      }
      for (const d of sup.load_dates) headers.push(`Load ${d} (${sup.name})`)
    }
    headers.push('LEFT')

    const dataRows = plan.items.map(item => {
      const isHi = hiItems.includes(item.item_code)
      const ratio = (isHi ? rhPct : rdPct) / 100
      const a1 = Math.round(item.qty_1 * plan.forecast_1 * ratio)
      const a2 = Math.round(item.qty_2 * plan.forecast_2 * ratio)
      const sumA = a1 + a2
      const loaded = totalLoadsMap.get(item.item_code) ?? 0
      const row: (string | number)[] = [item.item_code, item.description, a1 || 0, a2 || 0, `${isHi ? rhPct : rdPct}%`, sumA]
      for (const sup of plan.suppliers) {
        const m = allPoQtyMaps.get(sup.id) ?? new Map<string, number>()
        row.push(m.get(item.item_code) ?? 0)
        for (const poId of sup.po_ids) {
          for (const d of (sup.po_load_dates?.[poId] ?? [])) {
            const e = sup.loads.find(l => l.po_id === poId && l.date === d && l.item_code === item.item_code)
            row.push(e?.qty ?? 0)
          }
        }
        for (const d of sup.load_dates) {
          const e = sup.loads.find(l => !l.po_id && l.date === d && l.item_code === item.item_code)
          row.push(e?.qty ?? 0)
        }
      }
      row.push(sumA - loaded)
      return row
    })

    const wsData: unknown[][] = [...meta, headers, ...dataRows]
    const ws = XLSX.utils.aoa_to_sheet(wsData)
    ws['!cols'] = [{ wch: 26 }, { wch: 40 }, ...Array(headers.length - 2).fill({ wch: 14 })]
    XLSX.utils.book_append_sheet(wb, ws, 'Load Plan')
    XLSX.writeFile(wb, `${plan.name || 'Load Plan'}.xlsx`)
  }

  // ── Plan list / dashboard view ──
  if (!plan) {
    return (
      <>
        <NavBar onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />
        {lockMsg && (
          <div className="fixed top-16 right-4 z-50 bg-amber-50 border border-amber-300 text-amber-800 text-sm px-4 py-2.5 rounded-xl shadow-lg font-medium">
            🔒 กรุณาปลดล็อคก่อน (คลิกกุญแจมุมขวาบน)
          </div>
        )}
        <main className="p-6 max-w-5xl mx-auto">
          <div className="flex items-center justify-between mb-8">
            <div>
              <h1 className="text-2xl font-bold text-gray-900">Branch Load Plans</h1>
              <p className="text-sm text-gray-400 mt-0.5">แผนการโหลดสินค้าตามสาขา</p>
            </div>
            <button
              onClick={() => { setPlan(normalizePlan(mkPlan())); setExpandedSupIds(new Set()) }}
              className="px-4 py-2 bg-blue-600 text-white text-sm font-semibold rounded-xl hover:bg-blue-700 shadow-sm"
            >
              + New Plan
            </button>
          </div>
          {plans.length === 0 ? (
            <div className="text-center py-24 text-gray-400">
              <div className="text-5xl mb-4">📦</div>
              <p className="text-lg font-medium text-gray-500">ยังไม่มีแผนการโหลด</p>
              <p className="text-sm mt-1">สร้าง New Plan เพื่อเริ่มต้น</p>
            </div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {plans.map(p => {
                const updAt = p.updated_at ? new Date(p.updated_at).toLocaleString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—'
                const totalBranches = (p.forecast_1 ?? 0) + (p.forecast_2 ?? 0)
                const itemCount = (p.items ?? []).length
                const supCount = (p.suppliers ?? []).length
                return (
                  <div key={p.id} className="bg-white border border-gray-200 rounded-2xl shadow-sm hover:shadow-md hover:border-blue-300 transition-all flex flex-col">
                    <div
                      onClick={() => { setPlan(normalizePlan(p)); setExpandedSupIds(new Set()) }}
                      className="p-5 cursor-pointer flex-1"
                    >
                      <div className="font-bold text-gray-900 text-base leading-snug mb-3">{p.name || 'Untitled Plan'}</div>
                      <div className="flex gap-3 mb-3">
                        <div className="flex-1 bg-blue-50 rounded-xl px-3 py-2 text-center">
                          <div className="text-lg font-bold text-blue-700 tabular-nums">{(p.forecast_1 ?? 0).toLocaleString()}</div>
                          <div className="text-[10px] text-blue-500 font-medium leading-tight mt-0.5">{p.type_1_name || 'Existing'}<br />branches</div>
                        </div>
                        <div className="flex-1 bg-indigo-50 rounded-xl px-3 py-2 text-center">
                          <div className="text-lg font-bold text-indigo-700 tabular-nums">{(p.forecast_2 ?? 0).toLocaleString()}</div>
                          <div className="text-[10px] text-indigo-500 font-medium leading-tight mt-0.5">{p.type_2_name || 'New'}<br />branches</div>
                        </div>
                        <div className="flex-1 bg-gray-50 rounded-xl px-3 py-2 text-center">
                          <div className="text-lg font-bold text-gray-700 tabular-nums">{totalBranches.toLocaleString()}</div>
                          <div className="text-[10px] text-gray-400 font-medium leading-tight mt-0.5">รวม<br />สาขา</div>
                        </div>
                      </div>
                      <div className="flex flex-wrap gap-2 text-xs text-gray-500">
                        <span className="bg-gray-100 rounded-full px-2 py-0.5">{itemCount} items</span>
                        <span className="bg-gray-100 rounded-full px-2 py-0.5">{supCount} suppliers</span>
                      </div>
                    </div>
                    <div className="border-t border-gray-100 px-4 py-3 flex items-center justify-between gap-2">
                      <span className="text-[11px] text-gray-400">แก้ไขล่าสุด {updAt}</span>
                      <button
                        onClick={() => { setPlan(normalizePlan(p)); setExpandedSupIds(new Set()) }}
                        className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg transition-colors shrink-0"
                      >
                        เปิดดู →
                      </button>
                      <button
                        onClick={e => { e.stopPropagation(); if (window.confirm(`ลบ "${p.name || 'Untitled Plan'}" ?`)) deletePlan(p.id) }}
                        className="text-xs text-gray-400 hover:text-red-500 transition-colors px-1.5 py-0.5 rounded"
                        title="ลบ plan นี้"
                      >
                        ลบ
                      </button>
                    </div>
                  </div>
                )
              })}
            </div>
          )}
        </main>
      </>
    )
  }

  const hiItems = plan.high_ratio_items ?? []
  const rateDefault = plan.rate_default ?? 70
  const rateHigh = plan.rate_high ?? 90

  // ── Plan editor view ──
  return (
    <>
      <NavBar onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />
      {lockMsg && (
        <div className="fixed top-16 right-4 z-50 bg-amber-50 border border-amber-300 text-amber-800 text-sm px-4 py-2.5 rounded-xl shadow-lg font-medium">
          🔒 กรุณาปลดล็อคก่อน (คลิกกุญแจมุมขวาบน)
        </div>
      )}
      <main className="p-6 min-h-screen">
        {/* Lock banner */}
        {!unlocked && (
          <div className="mb-4 flex items-center gap-3 bg-amber-50 border border-amber-200 rounded-xl px-4 py-2.5 text-sm text-amber-800">
            <span>🔒</span>
            <span>โหมดดูข้อมูล — ปลดล็อคที่มุมขวาบนเพื่อแก้ไข</span>
          </div>
        )}
        {/* Top bar */}
        <div className="flex items-center gap-3 mb-6 flex-wrap">
          <button onClick={() => setPlan(null)} className="text-sm text-gray-500 hover:text-gray-800">← Plans</button>
          <input
            value={plan.name}
            readOnly={!unlocked}
            onChange={e => setPlan(p => p ? { ...p, name: e.target.value } : p)}
            className={`text-xl font-bold text-gray-900 border-b-2 border-transparent focus:outline-none bg-transparent flex-1 min-w-48 ${unlocked ? 'hover:border-gray-300 focus:border-blue-500' : 'cursor-default'}`}
          />
          <button
            onClick={exportToExcel}
            className="px-4 py-2 text-sm font-semibold rounded-xl border border-gray-300 text-gray-700 hover:bg-gray-50 transition-colors"
            title="Export เป็น Excel"
          >
            Export Excel
          </button>
          <button
            onClick={savePlan}
            disabled={saving}
            className={`px-5 py-2 text-sm font-semibold rounded-xl transition-colors disabled:opacity-50 ${saved ? 'bg-green-600 text-white' : unlocked ? 'bg-blue-600 hover:bg-blue-700 text-white' : 'bg-gray-200 text-gray-500'}`}
          >
            {saving ? 'Saving…' : saved ? 'Saved ✓' : unlocked ? 'Save' : '🔒 Save'}
          </button>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
          {/* Branch types & forecast */}
          <div className="bg-white border border-gray-200 rounded-xl p-4">
            <h2 className="text-sm font-semibold text-gray-700 mb-3">Branch Types &amp; Forecast</h2>
            <div className="space-y-2.5">
              {[1, 2].map(n => (
                <div key={n} className="flex items-center gap-2 flex-wrap">
                  <input
                    value={n === 1 ? plan.type_1_name : plan.type_2_name}
                    readOnly={!unlocked}
                    onChange={e => setPlan(p => p ? { ...p, [`type_${n}_name`]: e.target.value } : p)}
                    placeholder={`Type ${n} name`}
                    className="border border-gray-200 rounded-lg px-3 py-1.5 text-sm w-32 focus:outline-none focus:border-blue-400 disabled:bg-gray-50 read-only:bg-gray-50"
                  />
                  <span className="text-xs text-gray-400">Forecast</span>
                  <input
                    type="text" inputMode="numeric" pattern="[0-9]*"
                    value={n === 1 ? (plan.forecast_1 || '') : (plan.forecast_2 || '')}
                    readOnly={!unlocked}
                    onFocus={e => { if (unlocked) e.target.select() }}
                    onChange={e => {
                      if (!unlocked) return
                      const v = e.target.value.replace(/[^0-9]/g, '')
                      setPlan(p => p ? { ...p, [`forecast_${n}`]: v === '' ? 0 : parseInt(v, 10) } : p)
                    }}
                    placeholder="0"
                    className="border border-gray-200 rounded-lg px-3 py-1.5 text-sm w-24 focus:outline-none focus:border-blue-400 read-only:bg-gray-50"
                  />
                  <span className="text-xs text-gray-400">branches</span>
                </div>
              ))}
              {/* Assumed Qty rates */}
              <div className="border-t border-gray-100 pt-2.5 flex items-center gap-3 flex-wrap">
                <span className="text-xs text-gray-500 font-medium">Order Ratio</span>
                {[
                  { key: 'rate_default' as const, label: 'Standard', val: plan.rate_default ?? 70, color: 'text-gray-600' },
                  { key: 'rate_high' as const, label: 'High', val: plan.rate_high ?? 90, color: 'text-orange-600' },
                ].map(({ key, label, val, color }) => (
                  <div key={key} className="flex items-center gap-1.5">
                    <span className={`text-xs ${color}`}>{label}</span>
                    <input
                      type="text" inputMode="numeric" pattern="[0-9]*"
                      value={val || ''}
                      readOnly={!unlocked}
                      onFocus={e => { if (unlocked) e.target.select() }}
                      onChange={e => {
                        if (!unlocked) return
                        const v = e.target.value.replace(/[^0-9]/g, '')
                        setPlan(p => p ? { ...p, [key]: v === '' ? 0 : parseInt(v, 10) } : p)
                      }}
                      className={`border border-gray-200 rounded-lg px-2 py-1 text-sm w-14 text-center focus:outline-none focus:border-blue-400 read-only:bg-gray-50 ${color} font-semibold`}
                    />
                    <span className="text-xs text-gray-400">%</span>
                  </div>
                ))}
              </div>
              {/* Rate rules */}
              {unlocked && plan.items.length > 0 && (() => {
                const curPlan = plan!
                const kw = rateFilter.trim().toLowerCase()
                const matched = kw
                  ? curPlan.items.filter(it =>
                      it.item_code.toLowerCase().includes(kw) ||
                      it.description.toLowerCase().includes(kw)
                    )
                  : []
                const rules = curPlan.rate_rules ?? []

                function addRule(rate: 'high' | 'standard') {
                  if (!matched.length || !kw) return
                  const keyword = rateFilter.trim()
                  // Replace existing rule with same keyword, or append
                  const existing = rules.findIndex(r => r.keyword.toLowerCase() === keyword.toLowerCase())
                  const newRules = existing >= 0
                    ? rules.map((r, i) => i === existing ? { keyword: r.keyword, rate } : r)
                    : [...rules, { keyword, rate }]
                  const newHi = applyRateRules(curPlan.items, newRules)
                  setPlan(p => p ? { ...p, rate_rules: newRules, high_ratio_items: newHi } : p)
                  setRateFilter('')
                }

                function removeRule(idx: number) {
                  const newRules = rules.filter((_, i) => i !== idx)
                  const newHi = applyRateRules(curPlan.items, newRules)
                  setPlan(p => p ? { ...p, rate_rules: newRules, high_ratio_items: newHi } : p)
                }

                return (
                  <div className="border-t border-gray-100 pt-2.5 space-y-2">
                    {/* Saved rule chips */}
                    {rules.length > 0 && (
                      <div className="flex flex-wrap gap-1.5">
                        {rules.map((rule, idx) => (
                          <span key={idx} className={`inline-flex items-center gap-1 text-[11px] font-semibold px-2 py-0.5 rounded-full border ${
                            rule.rate === 'high'
                              ? 'bg-orange-50 border-orange-200 text-orange-700'
                              : 'bg-gray-100 border-gray-200 text-gray-600'
                          }`}>
                            {rule.keyword} → {rule.rate === 'high' ? `${rateHigh}%` : `${rateDefault}%`}
                            <button onClick={() => removeRule(idx)} className="hover:text-red-500 ml-0.5 leading-none">×</button>
                          </span>
                        ))}
                      </div>
                    )}
                    {/* Input row */}
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="text-xs text-gray-500 font-medium shrink-0">+ Rule</span>
                      <input
                        type="text"
                        value={rateFilter}
                        onChange={e => setRateFilter(e.target.value)}
                        onKeyDown={e => { if (e.key === 'Enter' && matched.length) addRule('high') }}
                        placeholder="keyword เช่น pole, shelf..."
                        className="border border-gray-200 rounded-lg px-2 py-1 text-xs flex-1 min-w-[140px] focus:outline-none focus:border-orange-400"
                      />
                      {kw && <span className="text-xs text-gray-400 shrink-0">{matched.length} items</span>}
                      <button
                        onClick={() => addRule('high')}
                        disabled={!matched.length}
                        className="px-2.5 py-1 text-xs rounded-lg bg-orange-100 text-orange-700 border border-orange-200 font-semibold hover:bg-orange-200 disabled:opacity-30 shrink-0"
                      >
                        → High {rateHigh}%
                      </button>
                      <button
                        onClick={() => addRule('standard')}
                        disabled={!matched.length}
                        className="px-2.5 py-1 text-xs rounded-lg bg-gray-100 text-gray-600 border border-gray-200 font-semibold hover:bg-gray-200 disabled:opacity-30 shrink-0"
                      >
                        → Standard {rateDefault}%
                      </button>
                    </div>
                    {/* Preview matched items */}
                    {kw && matched.length > 0 && (
                      <div className="flex flex-wrap gap-1">
                        {matched.slice(0, 6).map(it => (
                          <span key={it.item_code} className="text-[10px] font-mono bg-orange-50 border border-orange-100 rounded px-1.5 py-0.5 text-orange-700">{it.item_code}</span>
                        ))}
                        {matched.length > 6 && <span className="text-[10px] text-gray-400">+{matched.length - 6} more</span>}
                      </div>
                    )}
                    {kw && matched.length === 0 && (
                      <p className="text-[11px] text-gray-400">ไม่พบ item ที่ตรงกับ "{rateFilter}"</p>
                    )}
                  </div>
                )
              })()}
            </div>
          </div>

          {/* Item template */}
          <div className="bg-white border border-gray-200 rounded-xl p-4">
            <h2 className="text-sm font-semibold text-gray-700 mb-1">Item Template</h2>
            <p className="text-[11px] text-gray-400 mb-3">หา header อัตโนมัติ — ต้องมีคอลัมน์ <span className="font-mono">Item_No / Item Code</span> · <span className="font-mono">Description</span> · <span className="font-mono">Qty / Assumed Qty</span> (ลำดับคอลัมน์ไม่ตายตัว · "-" = 0)</p>
            <div className="flex flex-col gap-2">
              <div className="flex items-center gap-2 flex-wrap">
                <button
                  onClick={() => unlocked && templateRef1.current?.click()}
                  disabled={!unlocked}
                  className="px-3 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 font-medium disabled:opacity-40"
                >
                  Upload {plan.type_1_name} Template
                </button>
                <input ref={templateRef1} type="file" accept=".xlsx,.xls" className="hidden"
                  onChange={e => { const f = e.target.files?.[0]; if (f) handleTemplateUpload(f, 1); e.target.value = '' }} />
                {uploadedFiles.type1 && (
                  <span className="text-[11px] text-green-600 font-mono truncate max-w-[220px]" title={uploadedFiles.type1}>
                    ✓ {uploadedFiles.type1}
                  </span>
                )}
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                <button
                  onClick={() => unlocked && templateRef2.current?.click()}
                  disabled={!unlocked}
                  className="px-3 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 font-medium disabled:opacity-40"
                >
                  Upload {plan.type_2_name} Template
                </button>
                <input ref={templateRef2} type="file" accept=".xlsx,.xls" className="hidden"
                  onChange={e => { const f = e.target.files?.[0]; if (f) handleTemplateUpload(f, 2); e.target.value = '' }} />
                {uploadedFiles.type2 && (
                  <span className="text-[11px] text-green-600 font-mono truncate max-w-[220px]" title={uploadedFiles.type2}>
                    ✓ {uploadedFiles.type2}
                  </span>
                )}
              </div>
              {plan.items.length > 0 && (
                <span className="text-sm text-green-600 font-semibold">{plan.items.length} items loaded</span>
              )}
              {parseInfo && (
                <span className="text-[10px] font-mono text-gray-400 leading-tight">
                  [{parseInfo.type === 1 ? plan.type_1_name : plan.type_2_name}] {parseInfo.label}
                </span>
              )}
            </div>
          </div>
        </div>

        {/* Suppliers */}
        <div className="bg-white border border-gray-200 rounded-xl p-4 mb-6">
          <div className="flex items-center justify-between mb-3">
            <h2 className="text-sm font-semibold text-gray-700">Suppliers</h2>
            <button
              onClick={() => unlocked && setPlan(p => p ? { ...p, suppliers: [...p.suppliers, mkSupplier()] } : p)}
              disabled={!unlocked}
              className="px-3 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 text-sm font-semibold rounded-lg border border-blue-200 disabled:opacity-40"
            >
              + Add Supplier
            </button>
          </div>
          {plan.suppliers.length === 0
            ? <p className="text-sm text-gray-400 py-2">No suppliers yet. Add one above.</p>
            : (
              <div className="space-y-2">
                {plan.suppliers.map((sup, si) => (
                  <SupplierRow
                    key={sup.id}
                    sup={sup}
                    editable={unlocked}
                    poUploads={poUploads}
                    showPoSelector={showPoSelector === sup.id}
                    onUpdate={upd => updateSup(sup.id, upd)}
                    onUploadContainerQty={f => handleContainerQtyUpload(f, sup.id)}
                    onTogglePoSelector={() => setShowPoSelector(v => v === sup.id ? null : sup.id)}
                    onRemove={() => setPlan(p => p ? { ...p, suppliers: p.suppliers.filter((_, i) => i !== si) } : p)}
                    onAddPoDate={(poId, date) => addPoDate(sup.id, poId, date)}
                    onRemovePoDate={(poId, date) => removePoDate(sup.id, poId, date)}
                  />
                ))}
              </div>
            )}
        </div>

        {/* Main table */}
        {plan.items.length > 0 ? (
          <div className="bg-white border border-gray-200 rounded-xl overflow-hidden">
            {unlocked && hiItems.length > 0 && (
              <div className="px-4 py-2 bg-orange-50 border-b border-orange-100 text-xs text-orange-700">
                <span className="font-semibold">{hiItems.length} items</span> ใช้ rate {rateHigh}% — คลิกที่ badge ใน Rate column เพื่อเปลี่ยน
              </div>
            )}
            {unlocked && hiItems.length === 0 && (
              <div className="px-4 py-2 bg-gray-50 border-b border-gray-100 text-xs text-gray-400">
                ทุก item ใช้ rate {rateDefault}% — คลิกที่ badge ใน Rate column เพื่อเปลี่ยนรายการที่ต้องการเป็น {rateHigh}%
              </div>
            )}
            <div className="overflow-x-auto">
              <table className="w-full text-xs border-collapse min-w-max">
                <thead>
                  <tr className="bg-slate-800 text-white">
                    <th className="text-left px-3 py-3 font-semibold whitespace-nowrap sticky left-0 bg-slate-800 z-10">Item Code</th>
                    <th className="text-left px-3 py-3 font-semibold whitespace-nowrap">Description</th>
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap">
                      Assumed Qty<br /><span className="font-normal opacity-70 text-[10px]">{plan.type_1_name}</span>
                    </th>
                    {plan.forecast_2 > 0 && (
                      <th className="text-right px-3 py-3 font-semibold whitespace-nowrap">
                        Assumed Qty<br /><span className="font-normal opacity-70 text-[10px]">{plan.type_2_name}</span>
                      </th>
                    )}
                    <th className="text-center px-2 py-3 font-semibold whitespace-nowrap bg-amber-800 text-[10px]">
                      Ratio<br /><span className="font-normal opacity-70">×%</span>
                    </th>
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap">Sum Assumed<br />Qty</th>
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-emerald-900 border-l-2 border-emerald-700 text-[10px]">
                      PO Cover<br /><span className="font-normal opacity-70">PO − Assumed</span>
                    </th>
                    {/* Per-supplier columns */}
                    {plan.suppliers.map(sup => {
                      const isExp = expandedSupIds.has(sup.id)
                      const hasCQ = Object.keys(sup.container_qtys).length > 0
                      if (!isExp) {
                        return (
                          <th key={sup.id}
                            onClick={() => toggleExpanded(sup.id)}
                            className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-blue-900 hover:bg-blue-800 cursor-pointer border-l border-slate-600 select-none"
                            title={`คลิกเพื่อขยาย ${sup.name || 'Supplier'}`}
                          >
                            PO QTY<br /><span className="font-normal text-[10px] opacity-70">{sup.name || '—'} ▸</span>
                          </th>
                        )
                      }
                      return (
                        <Fragment key={sup.id}>
                          {hasCQ && (
                            <th className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-blue-950 border-l-2 border-blue-500">
                              Full Ctn QTY<br /><span className="font-normal text-[10px] opacity-70">{sup.name || '—'}</span>
                            </th>
                          )}
                          {/* Per-PO: each PO gets its own QTY column + date columns */}
                          {sup.po_ids.map((poId, pi) => {
                            const po = poUploads.find(p => p.id === poId)
                            const label = poLabel(po, poId)
                            const dates = sup.po_load_dates?.[poId] ?? []
                            return (
                              <Fragment key={poId}>
                                <th
                                  onClick={pi === 0 ? () => toggleExpanded(sup.id) : undefined}
                                  className={`text-right px-3 py-3 font-semibold whitespace-nowrap bg-blue-900 border-l-2 border-blue-600 select-none ${pi === 0 ? 'hover:bg-blue-800 cursor-pointer' : ''}`}
                                  title={pi === 0 ? `คลิกเพื่อซ่อน ${sup.name || 'Supplier'}` : undefined}
                                >
                                  <span className="font-mono text-[9px] opacity-60 block">{label}{pi === 0 ? ' ▾' : ''}</span>
                                  PO QTY
                                </th>
                                {dates.map(d => (
                                  <th key={`${poId}-${d}`} className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-indigo-900 border-l border-indigo-700">
                                    <span className="font-mono text-[9px] opacity-60 block">{label}</span>
                                    <span className="font-normal text-[10px] opacity-70">{fmtDate(d)}</span>
                                  </th>
                                ))}
                              </Fragment>
                            )
                          })}
                          {/* Legacy supplier-level dates */}
                          {sup.load_dates.map(d => (
                            <th key={`legacy-${d}`} className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-indigo-800 border-l border-indigo-600">
                              Load<br /><span className="font-normal text-[10px] opacity-70">{fmtDate(d)}</span>
                            </th>
                          ))}
                        </Fragment>
                      )
                    })}
                    {/* Daily totals per unique date across all suppliers */}
                    {allUniqueDates.map(d => (
                      <th key={`dt-${d}`} className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-teal-900 border-l border-teal-700">
                        รวมโหลด<br /><span className="font-normal text-[10px] opacity-70">{fmtDate(d)}</span>
                      </th>
                    ))}
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-rose-900 border-l-2 border-rose-700">PO LEFT</th>
                  </tr>
                </thead>
                <tbody>
                  {plan.items.map((item, ii) => {
                    const isHiItem = hiItems.includes(item.item_code)
                    const ratio = (isHiItem ? rateHigh : rateDefault) / 100
                    const a1 = Math.round(item.qty_1 * plan.forecast_1 * ratio)
                    const a2 = Math.round(item.qty_2 * plan.forecast_2 * ratio)
                    const sumA = a1 + a2
                    const loaded = totalLoadsMap.get(item.item_code) ?? 0
                    const totalPo = totalPoQtyMap.get(item.item_code) ?? 0
                    const left = totalPo - loaded
                    const base = ii % 2 === 0 ? 'bg-white' : 'bg-gray-50'
                    const isHi = isHiItem
                    return (
                      <tr key={item.item_code} className={`border-t border-gray-100 ${base}`}>
                        <td className="px-3 py-2 font-mono font-semibold text-gray-900 sticky left-0 z-10" style={{ background: 'inherit' }}>{item.item_code}</td>
                        <td className="px-3 py-2 text-gray-600 max-w-[280px] truncate">{item.description || '—'}</td>
                        <td
                          className="px-3 py-2 text-right tabular-nums text-gray-700 cursor-pointer hover:bg-blue-50 hover:text-blue-700"
                          onClick={() => setFormulaPopup({
                            itemCode: item.item_code, description: item.description,
                            colName: `Assumed Qty (${plan.type_1_name})`,
                            formulaStr: `Assumed Qty/branch × Forecast ${plan.type_1_name} × Ratio`,
                            lines: [
                              { op: '', label: `Assumed Qty / branch`, val: item.qty_1 },
                              { op: '×', label: `Forecast (${plan.type_1_name})`, val: plan.forecast_1 },
                              { op: '×', label: `Ratio (${isHi ? 'High' : 'Standard'})`, val: `${isHi ? rateHigh : rateDefault}%` },
                              { op: '=', label: 'Assumed Qty', val: a1, isResult: true },
                            ]
                          })}
                        >{a1 > 0 ? a1.toLocaleString() : '—'}</td>
                        {plan.forecast_2 > 0 && (
                          <td
                            className="px-3 py-2 text-right tabular-nums text-gray-700 cursor-pointer hover:bg-blue-50 hover:text-blue-700"
                            onClick={() => setFormulaPopup({
                              itemCode: item.item_code, description: item.description,
                              colName: `Assumed Qty (${plan.type_2_name})`,
                              formulaStr: `Assumed Qty/branch × Forecast ${plan.type_2_name} × Ratio`,
                              lines: [
                                { op: '', label: `Assumed Qty / branch`, val: item.qty_2 },
                                { op: '×', label: `Forecast (${plan.type_2_name})`, val: plan.forecast_2 },
                                { op: '×', label: `Ratio (${isHi ? 'High' : 'Standard'})`, val: `${isHi ? rateHigh : rateDefault}%` },
                                { op: '=', label: 'Assumed Qty', val: a2, isResult: true },
                              ]
                            })}
                          >{a2 > 0 ? a2.toLocaleString() : '—'}</td>
                        )}
                        {/* Rate toggle */}
                        <td className="px-2 py-1 text-center bg-amber-50/50">
                          <button
                            onClick={() => toggleHighRatio(item.item_code)}
                            title={unlocked ? (isHi ? 'คลิกเพื่อเปลี่ยนเป็น 70%' : 'คลิกเพื่อเปลี่ยนเป็น 90%') : undefined}
                            className={`text-[10px] font-bold px-1.5 py-0.5 rounded border transition-colors ${
                              isHi
                                ? 'bg-orange-100 text-orange-700 border-orange-300'
                                : 'bg-gray-100 text-gray-500 border-gray-200'
                            } ${unlocked ? 'cursor-pointer hover:opacity-70' : 'cursor-default'}`}
                          >
                            {isHi ? `${rateHigh}%` : `${rateDefault}%`}
                          </button>
                        </td>
                        <td
                          className="px-3 py-2 text-right tabular-nums font-semibold text-gray-900 cursor-pointer hover:bg-blue-50 hover:text-blue-700"
                          onClick={() => setFormulaPopup({
                            itemCode: item.item_code, description: item.description,
                            colName: 'Sum Assumed Qty',
                            formulaStr: `Assumed (${plan.type_1_name}) + Assumed (${plan.type_2_name})`,
                            lines: [
                              { op: '', label: `Assumed (${plan.type_1_name})`, val: a1 },
                              { op: '+', label: `Assumed (${plan.type_2_name})`, val: a2 },
                              { op: '=', label: 'Sum Assumed Qty', val: sumA, isResult: true },
                            ]
                          })}
                        >{sumA > 0 ? sumA.toLocaleString() : '0'}</td>
                        {/* PO Coverage: sum all PO qty − sumA */}
                        {(() => {
                          const cov = totalPo - sumA
                          return (
                            <td
                              className={`px-3 py-2 text-right tabular-nums font-bold border-l-2 cursor-pointer ${
                                cov >= 0 ? 'text-green-700 bg-green-50 border-emerald-300 hover:bg-green-100' : 'text-red-600 bg-red-50 border-red-300 hover:bg-red-100'
                              }`}
                              onClick={() => setFormulaPopup({
                                itemCode: item.item_code, description: item.description,
                                colName: 'PO Cover',
                                formulaStr: 'Total PO QTY (ทุกซัพ) − Sum Assumed Qty',
                                lines: [
                                  { op: '', label: 'Total PO QTY (ทุกซัพ)', val: totalPo },
                                  { op: '−', label: 'Sum Assumed Qty', val: sumA },
                                  { op: '=', label: cov >= 0 ? 'Buffer ที่เหลือ' : 'ขาด (ต้องสั่งเพิ่ม)', val: cov, isResult: true },
                                ]
                              })}
                            >
                              {cov > 0 ? '+' : ''}{cov !== 0 ? cov.toLocaleString() : totalPo === 0 ? '—' : '0'}
                            </td>
                          )
                        })()}
                        {/* Per-supplier cells */}
                        {plan.suppliers.map(sup => {
                          const supPoMap = allPoQtyMaps.get(sup.id) ?? new Map<string, number>()
                          const isExp = expandedSupIds.has(sup.id)
                          const hasCQ = Object.keys(sup.container_qtys).length > 0
                          const poQty = supPoMap.get(item.item_code) ?? 0
                          if (!isExp) {
                            return (
                              <td key={sup.id} className="px-3 py-2 text-right tabular-nums text-blue-900 bg-blue-50 border-l border-blue-100">
                                {poQty > 0 ? poQty.toLocaleString() : '—'}
                              </td>
                            )
                          }
                          return (
                            <Fragment key={sup.id}>
                              {hasCQ && (
                                <td className="px-3 py-2 text-right tabular-nums text-blue-900 bg-blue-50/70 border-l-2 border-blue-200">
                                  {sup.container_qtys[item.item_code]?.toLocaleString() ?? '—'}
                                </td>
                              )}
                              {/* Per-PO: individual PO qty + its load dates */}
                              {sup.po_ids.map(poId => {
                                const po = poUploads.find(p => p.id === poId)
                                const thisPo = po?.rows ?? []
                                const thisPoQty = thisPo.filter(r => r.item_code === item.item_code).reduce((s, r) => s + (r.qty ?? 0), 0)
                                const dates = sup.po_load_dates?.[poId] ?? []
                                return (
                                  <Fragment key={poId}>
                                    <td className="px-3 py-2 text-right tabular-nums text-blue-900 bg-blue-50 border-l-2 border-blue-300 font-semibold">
                                      {thisPoQty > 0 ? thisPoQty.toLocaleString() : '—'}
                                    </td>
                                    {dates.map(d => {
                                      const entry = sup.loads.find(l => l.po_id === poId && l.date === d && l.item_code === item.item_code)
                                      const isAuto = entry?.is_auto === true
                                      return (
                                        <td key={`${poId}-${d}`} className="px-1.5 py-1 bg-indigo-50 border-l border-indigo-100">
                                          <input
                                            type="number" min="0"
                                            value={entry?.qty ?? ''}
                                            placeholder="0"
                                            disabled={!unlocked}
                                            onChange={e => updateLoadQty(sup.id, poId, d, item.item_code, Number(e.target.value))}
                                            className={`w-20 text-right px-2 py-0.5 border rounded focus:outline-none focus:border-indigo-500 tabular-nums text-xs disabled:bg-gray-50 disabled:opacity-60 ${isAuto ? 'bg-sky-100 border-sky-300 text-sky-800 italic' : 'bg-white border-indigo-200'}`}
                                            title={isAuto ? 'ค่า auto-suggested — แก้ไขได้' : undefined}
                                          />
                                        </td>
                                      )
                                    })}
                                  </Fragment>
                                )
                              })}
                              {/* Legacy supplier-level dates */}
                              {sup.load_dates.map(d => {
                                const entry = sup.loads.find(l => !l.po_id && l.date === d && l.item_code === item.item_code)
                                return (
                                  <td key={`legacy-${d}`} className="px-1.5 py-1 bg-indigo-50/70 border-l border-indigo-100">
                                    <input
                                      type="number" min="0"
                                      value={entry?.qty ?? ''}
                                      placeholder="0"
                                      disabled={!unlocked}
                                      onChange={e => updateLoadQty(sup.id, '', d, item.item_code, Number(e.target.value))}
                                      className="w-20 text-right px-2 py-0.5 border border-indigo-200 rounded focus:outline-none focus:border-indigo-500 tabular-nums bg-white text-xs disabled:bg-gray-50 disabled:opacity-60"
                                    />
                                  </td>
                                )
                              })}
                            </Fragment>
                          )
                        })}
                        {/* Daily totals per date */}
                        {allUniqueDates.map(d => {
                          const total = dailyTotalsMap.get(d)?.get(item.item_code) ?? 0
                          return (
                            <td
                              key={`dt-${d}`}
                              className="px-3 py-2 text-right tabular-nums font-semibold text-teal-800 bg-teal-50 border-l border-teal-100 cursor-pointer hover:bg-teal-100"
                              onClick={() => {
                                const breakdown = plan.suppliers.flatMap(s =>
                                  s.loads.filter(l => l.date === d && l.item_code === item.item_code && l.qty > 0)
                                    .map(l => ({ supName: s.name, qty: l.qty, poId: l.po_id }))
                                )
                                setFormulaPopup({
                                  itemCode: item.item_code, description: item.description,
                                  colName: `รวมโหลด ${d}`,
                                  formulaStr: `Σ ทุกซัพที่โหลดวันที่ ${d}`,
                                  lines: [
                                    ...breakdown.map((b, i) => ({
                                      op: (i === 0 ? '' : '+') as FLine['op'],
                                      label: b.supName + (b.poId ? ` (${b.poId.slice(0, 10)})` : ''),
                                      val: b.qty,
                                    })),
                                    { op: '=', label: 'รวม', val: total, isResult: true },
                                  ]
                                })
                              }}
                            >
                              {total > 0 ? total.toLocaleString() : '—'}
                            </td>
                          )
                        })}
                        <td
                          className={`px-3 py-2 text-right tabular-nums font-bold border-l-2 cursor-pointer ${left < 0 ? 'text-red-600 bg-red-50 border-red-200 hover:bg-red-100' : left === 0 ? 'text-green-700 bg-green-50 border-green-200 hover:bg-green-100' : 'text-gray-900 border-rose-200 hover:bg-gray-100'}`}
                          onClick={() => setFormulaPopup({
                            itemCode: item.item_code, description: item.description,
                            colName: 'PO LEFT',
                            formulaStr: 'Sum PO QTY (ทุกซัพ) − รวมโหลดแล้ว',
                            lines: [
                              { op: '', label: 'Sum PO QTY (ทุกซัพ)', val: totalPo },
                              { op: '−', label: 'รวมโหลดแล้ว', val: loaded },
                              { op: '=', label: left < 0 ? 'โหลดเกิน PO' : left === 0 ? 'โหลดครบ PO' : 'PO ที่ยังโหลดได้', val: left, isResult: true },
                            ]
                          })}
                        >
                          {left.toLocaleString()}
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          <div className="bg-white border-2 border-dashed border-gray-200 rounded-xl py-16 text-center text-gray-400">
            <p className="text-3xl mb-2">📋</p>
            <p className="font-medium">Upload an item template to see the planning table</p>
            <p className="text-xs mt-1">Excel: item_code · description · qty/{plan.type_1_name} · qty/{plan.type_2_name}</p>
          </div>
        )}
      </main>

      {/* Formula popup */}
      {formulaPopup && (
        <div className="fixed inset-0 bg-black/20 z-50 flex items-center justify-center p-4" onClick={() => setFormulaPopup(null)}>
          <div className="bg-white rounded-2xl shadow-2xl p-6 max-w-sm w-full" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-start mb-1">
              <div>
                <p className="text-xs text-gray-400 font-mono">{formulaPopup.itemCode}</p>
                <p className="text-xs text-gray-500 mt-0.5 truncate max-w-[220px]">{formulaPopup.description}</p>
                <h3 className="text-base font-bold text-gray-900 mt-1">{formulaPopup.colName}</h3>
              </div>
              <button onClick={() => setFormulaPopup(null)} className="text-gray-300 hover:text-gray-600 text-lg leading-none mt-1">✕</button>
            </div>
            {formulaPopup.formulaStr && (
              <p className="text-xs text-gray-500 mb-3 font-mono bg-gray-50 rounded-lg px-3 py-1.5 mt-2">{formulaPopup.formulaStr}</p>
            )}
            <div className="space-y-1 mt-2">
              {formulaPopup.lines.map((line, idx) => (
                <div key={idx}>
                  {line.isResult && <div className="border-t-2 border-gray-200 my-2" />}
                  <div className={`flex items-baseline gap-2 px-1 py-0.5 rounded ${line.isResult ? 'bg-gray-50' : ''}`}>
                    <span className="text-xs w-4 text-right shrink-0 text-gray-400 font-mono">{line.op}</span>
                    <span className={`flex-1 text-sm ${line.isResult ? 'font-bold text-gray-900' : 'text-gray-600'}`}>{line.label}</span>
                    <span className={`text-sm font-mono tabular-nums shrink-0 ${line.isResult ? 'font-bold text-gray-900' : 'text-gray-700'} ${typeof line.val === 'number' && line.val < 0 ? 'text-red-600' : ''}`}>
                      {typeof line.val === 'number' ? line.val.toLocaleString() : line.val}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </>
  )
}

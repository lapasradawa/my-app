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

interface LoadEntry {
  date: string
  item_code: string
  qty: number
}

interface LoadSupplier {
  id: string
  name: string
  po_ids: string[]
  container_qtys: Record<string, number>
  loads: LoadEntry[]
  load_dates: string[]
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
  return { name: 'Untitled Plan', type_1_name: 'Existing', type_2_name: 'New', forecast_1: 0, forecast_2: 0, items: [], suppliers: [] }
}

function mkSupplier(): LoadSupplier {
  return { id: crypto.randomUUID(), name: '', po_ids: [], container_qtys: {}, loads: [], load_dates: [] }
}

function fmtDate(d: string) {
  return new Date(d + 'T00:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
}

// ── Supplier row component ──
function SupplierRow({ sup, poUploads, showPoSelector, onUpdate, onUploadContainerQty, onTogglePoSelector, onRemove }: {
  sup: LoadSupplier
  poUploads: POUpload[]
  showPoSelector: boolean
  onUpdate: (upd: Partial<LoadSupplier>) => void
  onUploadContainerQty: (f: File) => void
  onTogglePoSelector: () => void
  onRemove: () => void
}) {
  const fileRef = useRef<HTMLInputElement>(null)
  const [newDate, setNewDate] = useState('')
  const containerCount = Object.keys(sup.container_qtys).length

  return (
    <div className="rounded-xl border border-gray-200 bg-white p-3">
      <div className="flex items-center gap-2 flex-wrap">
        <input
          value={sup.name}
          onChange={e => onUpdate({ name: e.target.value })}
          placeholder="Supplier name"
          className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm w-36 focus:outline-none focus:border-blue-400 bg-white"
        />

        {/* PO selector */}
        <div className="relative">
          <button
            onClick={onTogglePoSelector}
            className="px-2 py-1.5 border border-gray-300 rounded-lg text-xs text-gray-700 hover:bg-gray-50 bg-white font-medium"
          >
            POs ({sup.po_ids.length}) ▾
          </button>
          {showPoSelector && (
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
          onClick={() => fileRef.current?.click()}
          title="Excel: A = item_code · B = container qty (optional)"
          className="px-2 py-1.5 border border-gray-300 rounded-lg text-xs text-gray-700 hover:bg-gray-50 bg-white font-medium"
        >
          {containerCount > 0 ? `Container QTY (${containerCount})` : 'Container QTY (optional)'}
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) onUploadContainerQty(f); e.target.value = '' }} />

        {/* Add load date */}
        <input type="date" value={newDate} onChange={e => setNewDate(e.target.value)}
          className="border border-gray-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:border-blue-400 bg-white" />
        <button
          onClick={() => {
            if (!newDate || sup.load_dates.includes(newDate)) return
            onUpdate({ load_dates: [...sup.load_dates, newDate].sort() })
            setNewDate('')
          }}
          disabled={!newDate}
          className="px-2 py-1.5 bg-gray-100 hover:bg-gray-200 text-xs rounded-lg font-medium disabled:opacity-40"
        >
          + Date
        </button>

        <button onClick={onRemove} className="text-gray-400 hover:text-red-500 text-lg leading-none px-1 ml-auto">×</button>
      </div>

      {sup.load_dates.length > 0 && (
        <div className="flex gap-1.5 flex-wrap mt-2">
          {sup.load_dates.map(d => (
            <span key={d} className="inline-flex items-center gap-1 bg-white border border-gray-200 rounded-full px-2.5 py-0.5 text-xs text-gray-700">
              {fmtDate(d)}
              <button
                onClick={() => onUpdate({ load_dates: sup.load_dates.filter(x => x !== d), loads: sup.loads.filter(l => l.date !== d) })}
                className="text-gray-400 hover:text-red-500 ml-0.5"
              >×</button>
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
  const templateRef1 = useRef<HTMLInputElement>(null)
  const templateRef2 = useRef<HTMLInputElement>(null)

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

  useEffect(() => {
    setUnlocked(isUnlocked())
    supabase.from('load_plans').select('*').order('created_at', { ascending: false }).then(({ data }) => {
      if (data) setPlans(data as (LoadPlan & { id: string })[])
    })
    supabase.from('po_uploads').select('id, supplier, project, po_rbs_ch_no, filename, created_at, rows').then(({ data }) => {
      if (data) setPoUploads(data as POUpload[])
    })
  }, [])

  const refreshPlans = useCallback(async () => {
    const { data } = await supabase.from('load_plans').select('*').order('created_at', { ascending: false })
    if (data) setPlans(data as (LoadPlan & { id: string })[])
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

      // Scan first 5 rows. Prioritise rows that have item_code column (fi >= 0).
      // A row with only qty keyword (e.g. a merged title "Assumed Qty…") is kept as
      // a fallback but the scan continues until a proper header row is found.
      let headerRow = 0
      let itemCol = 1   // fallback: col B
      let descCol = 2   // fallback: col C
      let qtyCol  = 3   // fallback: col D

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

      // Helper: parse a cell value as number, handling comma-formatted strings like "1,216"
      function parseQty(v: unknown): number {
        if (v === null || v === undefined || v === '' || v === '-') return 0
        if (typeof v === 'number') return isNaN(v) ? 0 : v
        const s = String(v).trim().replace(/,/g, '')
        return s === '' || s === '-' ? 0 : (Number(s) || 0)
      }

      // Verify qtyCol: if first several data rows all give 0, search nearby columns for non-zero numeric values
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

      // Show diagnostic: what columns were detected, first item's qty
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

  function updateLoadQty(suppId: string, date: string, itemCode: string, qty: number) {
    setPlan(p => {
      if (!p) return p
      const sup = p.suppliers.find(s => s.id === suppId)
      if (!sup) return p
      const idx = sup.loads.findIndex(l => l.date === date && l.item_code === itemCode)
      let newLoads: LoadEntry[]
      if (qty <= 0) {
        newLoads = sup.loads.filter(l => !(l.date === date && l.item_code === itemCode))
      } else if (idx >= 0) {
        newLoads = sup.loads.map((l, i) => i === idx ? { ...l, qty } : l)
      } else {
        newLoads = [...sup.loads, { date, item_code: itemCode, qty }]
      }
      return { ...p, suppliers: p.suppliers.map(s => s.id === suppId ? { ...s, loads: newLoads } : s) }
    })
  }

  function exportToExcel() {
    if (!plan) return
    const wb = XLSX.utils.book_new()

    // Meta rows
    const meta = [
      [`Plan: ${plan.name || 'Untitled Plan'}`],
      [`${plan.type_1_name}: ${plan.forecast_1} branches   ${plan.type_2_name}: ${plan.forecast_2} branches`],
      [`Exported: ${new Date().toLocaleString('en-GB')}`],
      [],
    ]

    // Column headers
    const headers: string[] = ['Item Code', 'Description', `Assumed Qty (${plan.type_1_name})`, `Assumed Qty (${plan.type_2_name})`, 'Sum Assumed Qty']
    for (const sup of plan.suppliers) {
      headers.push(`PO QTY (${sup.name})`)
      for (const d of sup.load_dates) headers.push(`Load ${d} (${sup.name})`)
    }
    headers.push('LEFT')

    // Data rows
    const dataRows = plan.items.map(item => {
      const a1 = item.qty_1 * plan.forecast_1
      const a2 = item.qty_2 * plan.forecast_2
      const sumA = a1 + a2
      const loaded = totalLoadsMap.get(item.item_code) ?? 0
      const row: (string | number)[] = [item.item_code, item.description, a1 || 0, a2 || 0, sumA]
      for (const sup of plan.suppliers) {
        const m = allPoQtyMaps.get(sup.id) ?? new Map<string, number>()
        row.push(m.get(item.item_code) ?? 0)
        for (const d of sup.load_dates) {
          const e = sup.loads.find(l => l.date === d && l.item_code === item.item_code)
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
              onClick={() => { setPlan(mkPlan()); setExpandedSupIds(new Set()) }}
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
                    {/* Card header */}
                    <div
                      onClick={() => { setPlan(p); setExpandedSupIds(new Set()) }}
                      className="p-5 cursor-pointer flex-1"
                    >
                      <div className="font-bold text-gray-900 text-base leading-snug mb-3">{p.name || 'Untitled Plan'}</div>
                      {/* Branch stats */}
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
                      {/* Meta */}
                      <div className="flex flex-wrap gap-2 text-xs text-gray-500">
                        <span className="bg-gray-100 rounded-full px-2 py-0.5">{itemCount} items</span>
                        <span className="bg-gray-100 rounded-full px-2 py-0.5">{supCount} suppliers</span>
                      </div>
                    </div>
                    {/* Card footer */}
                    <div className="border-t border-gray-100 px-4 py-3 flex items-center justify-between gap-2">
                      <div>
                        <span className="text-[11px] text-gray-400">แก้ไขล่าสุด {updAt}</span>
                      </div>
                      <button
                        onClick={() => { setPlan(p); setExpandedSupIds(new Set()) }}
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
        {/* Top bar */}
        <div className="flex items-center gap-3 mb-6 flex-wrap">
          <button onClick={() => setPlan(null)} className="text-sm text-gray-500 hover:text-gray-800">← Plans</button>
          <input
            value={plan.name}
            onChange={e => setPlan(p => p ? { ...p, name: e.target.value } : p)}
            className="text-xl font-bold text-gray-900 border-b-2 border-transparent hover:border-gray-300 focus:border-blue-500 focus:outline-none bg-transparent flex-1 min-w-48"
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
                    onChange={e => setPlan(p => p ? { ...p, [`type_${n}_name`]: e.target.value } : p)}
                    placeholder={`Type ${n} name`}
                    className="border border-gray-200 rounded-lg px-3 py-1.5 text-sm w-32 focus:outline-none focus:border-blue-400"
                  />
                  <span className="text-xs text-gray-400">Forecast</span>
                  <input
                    type="text" inputMode="numeric" pattern="[0-9]*"
                    value={n === 1 ? (plan.forecast_1 || '') : (plan.forecast_2 || '')}
                    onFocus={e => e.target.select()}
                    onChange={e => {
                      const v = e.target.value.replace(/[^0-9]/g, '')
                      setPlan(p => p ? { ...p, [`forecast_${n}`]: v === '' ? 0 : parseInt(v, 10) } : p)
                    }}
                    placeholder="0"
                    className="border border-gray-200 rounded-lg px-3 py-1.5 text-sm w-24 focus:outline-none focus:border-blue-400"
                  />
                  <span className="text-xs text-gray-400">branches</span>
                </div>
              ))}
            </div>
          </div>

          {/* Item template */}
          <div className="bg-white border border-gray-200 rounded-xl p-4">
            <h2 className="text-sm font-semibold text-gray-700 mb-1">Item Template</h2>
            <p className="text-[11px] text-gray-400 mb-3">หา header อัตโนมัติ — ต้องมีคอลัมน์ <span className="font-mono">Item_No / Item Code</span> · <span className="font-mono">Description</span> · <span className="font-mono">Qty / Assumed Qty</span> (ลำดับคอลัมน์ไม่ตายตัว · "-" = 0)</p>
            <div className="flex flex-col gap-2">
              <div className="flex items-center gap-2 flex-wrap">
                <button
                  onClick={() => templateRef1.current?.click()}
                  className="px-3 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 font-medium"
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
                  onClick={() => templateRef2.current?.click()}
                  className="px-3 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 font-medium"
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
              onClick={() => setPlan(p => p ? { ...p, suppliers: [...p.suppliers, mkSupplier()] } : p)}
              className="px-3 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 text-sm font-semibold rounded-lg border border-blue-200"
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
                    poUploads={poUploads}
                    showPoSelector={showPoSelector === sup.id}
                    onUpdate={upd => updateSup(sup.id, upd)}
                    onUploadContainerQty={f => handleContainerQtyUpload(f, sup.id)}
                    onTogglePoSelector={() => setShowPoSelector(v => v === sup.id ? null : sup.id)}
                    onRemove={() => setPlan(p => p ? { ...p, suppliers: p.suppliers.filter((_, i) => i !== si) } : p)}
                  />
                ))}
              </div>
            )}
        </div>

        {/* Main table */}
        {plan.items.length > 0 ? (
          <div className="bg-white border border-gray-200 rounded-xl overflow-hidden">
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
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap">Sum Assumed<br />Qty</th>
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
                          <th
                            onClick={() => toggleExpanded(sup.id)}
                            className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-blue-900 hover:bg-blue-800 cursor-pointer border-l border-blue-700 select-none"
                            title={`คลิกเพื่อซ่อน ${sup.name || 'Supplier'}`}
                          >
                            PO QTY<br /><span className="font-normal text-[10px] opacity-70">{sup.name || '—'} ▾</span>
                          </th>
                          {sup.load_dates.map(d => (
                            <th key={d} className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-indigo-900 border-l border-indigo-700">
                              Load<br /><span className="font-normal text-[10px] opacity-70">{fmtDate(d)}</span>
                            </th>
                          ))}
                        </Fragment>
                      )
                    })}
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-rose-900 border-l-2 border-rose-700">LEFT</th>
                  </tr>
                </thead>
                <tbody>
                  {plan.items.map((item, ii) => {
                    const a1 = item.qty_1 * plan.forecast_1
                    const a2 = item.qty_2 * plan.forecast_2
                    const sumA = a1 + a2
                    const loaded = totalLoadsMap.get(item.item_code) ?? 0
                    const left = sumA - loaded
                    const base = ii % 2 === 0 ? 'bg-white' : 'bg-gray-50'
                    return (
                      <tr key={item.item_code} className={`border-t border-gray-100 ${base}`}>
                        <td className="px-3 py-2 font-mono font-semibold text-gray-900 sticky left-0 z-10" style={{ background: 'inherit' }}>{item.item_code}</td>
                        <td className="px-3 py-2 text-gray-600 max-w-[280px] truncate">{item.description || '—'}</td>
                        <td className="px-3 py-2 text-right tabular-nums text-gray-700">{a1 > 0 ? a1.toLocaleString() : '—'}</td>
                        {plan.forecast_2 > 0 && (
                          <td className="px-3 py-2 text-right tabular-nums text-gray-700">{a2 > 0 ? a2.toLocaleString() : '—'}</td>
                        )}
                        <td className="px-3 py-2 text-right tabular-nums font-semibold text-gray-900">{sumA > 0 ? sumA.toLocaleString() : '0'}</td>
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
                              <td className="px-3 py-2 text-right tabular-nums text-blue-900 bg-blue-50 border-l border-blue-100">
                                {poQty > 0 ? poQty.toLocaleString() : '—'}
                              </td>
                              {sup.load_dates.map(d => {
                                const entry = sup.loads.find(l => l.date === d && l.item_code === item.item_code)
                                return (
                                  <td key={d} className="px-1.5 py-1 bg-indigo-50 border-l border-indigo-100">
                                    <input
                                      type="number" min="0"
                                      value={entry?.qty ?? ''}
                                      placeholder="0"
                                      onChange={e => updateLoadQty(sup.id, d, item.item_code, Number(e.target.value))}
                                      className="w-20 text-right px-2 py-0.5 border border-indigo-200 rounded focus:outline-none focus:border-indigo-500 tabular-nums bg-white text-xs"
                                    />
                                  </td>
                                )
                              })}
                            </Fragment>
                          )
                        })}
                        <td className={`px-3 py-2 text-right tabular-nums font-bold border-l-2 ${left < 0 ? 'text-red-600 bg-red-50 border-red-200' : left === 0 ? 'text-green-700 bg-green-50 border-green-200' : 'text-gray-900 border-rose-200'}`}>
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
    </>
  )
}

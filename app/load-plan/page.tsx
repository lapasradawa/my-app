'use client'

import { useState, useEffect, useRef, useMemo, useCallback } from 'react'
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
function SupplierRow({ sup, isActive, poUploads, showPoSelector, onToggleActive, onUpdate, onUploadContainerQty, onTogglePoSelector, onRemove }: {
  sup: LoadSupplier
  isActive: boolean
  poUploads: POUpload[]
  showPoSelector: boolean
  onToggleActive: () => void
  onUpdate: (upd: Partial<LoadSupplier>) => void
  onUploadContainerQty: (f: File) => void
  onTogglePoSelector: () => void
  onRemove: () => void
}) {
  const fileRef = useRef<HTMLInputElement>(null)
  const [newDate, setNewDate] = useState('')
  const containerCount = Object.keys(sup.container_qtys).length

  return (
    <div className={`rounded-xl border p-3 transition-colors ${isActive ? 'border-blue-400 bg-blue-50' : 'border-gray-200 bg-white'}`}>
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
              <div className="sticky top-0 bg-gray-50 px-3 py-2 text-xs text-gray-500 border-b font-semibold">Select POs</div>
              {poUploads.length === 0
                ? <div className="p-3 text-xs text-gray-400">No POs available</div>
                : poUploads.map(po => (
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
                ))}
            </div>
          )}
        </div>

        {/* Container QTY upload */}
        <button
          onClick={() => fileRef.current?.click()}
          title="Excel: A = item_code · B = container qty"
          className="px-2 py-1.5 border border-gray-300 rounded-lg text-xs text-gray-700 hover:bg-gray-50 bg-white font-medium"
        >
          {containerCount > 0 ? `Container QTY (${containerCount})` : 'Upload Container QTY'}
        </button>
        <span className="text-[10px] text-gray-400 font-mono">A: item_code · B: qty/container</span>
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

        <button
          onClick={onToggleActive}
          className={`px-3 py-1.5 text-xs rounded-lg font-semibold ml-auto transition-colors ${isActive ? 'bg-blue-600 text-white shadow-sm' : 'bg-gray-100 hover:bg-gray-200 text-gray-700'}`}
        >
          {isActive ? '▼ Working' : '▶ Work'}
        </button>
        <button onClick={onRemove} className="text-gray-400 hover:text-red-500 text-lg leading-none px-1">×</button>
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
  const [, setUnlocked] = useState(false)
  const [plans, setPlans] = useState<(LoadPlan & { id: string })[]>([])
  const [plan, setPlan] = useState<LoadPlan | null>(null)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [activeSupIdx, setActiveSupIdx] = useState<number | null>(null)
  const [poUploads, setPoUploads] = useState<POUpload[]>([])
  const [showPoSelector, setShowPoSelector] = useState<string | null>(null)
  const templateRef1 = useRef<HTMLInputElement>(null)
  const templateRef2 = useRef<HTMLInputElement>(null)

  const activeSup = (plan && activeSupIdx !== null) ? (plan.suppliers[activeSupIdx] ?? null) : null

  const poQtyMap = useMemo(() => {
    const map = new Map<string, number>()
    if (!activeSup) return map
    for (const poId of activeSup.po_ids) {
      const po = poUploads.find(p => p.id === poId)
      if (!po?.rows) continue
      for (const r of po.rows) {
        if (r.item_code) map.set(r.item_code, (map.get(r.item_code) ?? 0) + (r.qty ?? 0))
      }
    }
    return map
  }, [activeSup, poUploads])

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

  const savePlan = useCallback(async () => {
    if (!plan) return
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

  // Parse template file: col A=NO (skip), col B=item_code, col C=description, col D=qty
  // Updates only qty_1 or qty_2 depending on typeNum, merges with existing items
  function handleTemplateUpload(file: File, typeNum: 1 | 2) {
    file.arrayBuffer().then(buf => {
      const wb = XLSX.read(buf, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][]
      const parsed = new Map<string, { description: string; qty: number }>()
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i] as unknown[]
        const item_code = String(r[1] ?? '').trim()
        if (!item_code) continue
        const raw = String(r[3] ?? '').trim()
        const qty = raw === '-' || raw === '' ? 0 : Number(raw) || 0
        parsed.set(item_code, { description: String(r[2] ?? '').trim(), qty })
      }
      setPlan(p => {
        if (!p) return p
        // Merge with existing items
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

  // ── Plan list view ──
  if (!plan) {
    return (
      <>
        <NavBar onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />
        <main className="p-6 max-w-3xl mx-auto">
          <div className="flex items-center justify-between mb-6">
            <h1 className="text-2xl font-bold text-gray-900">Branch Load</h1>
            <button
              onClick={() => { setPlan(mkPlan()); setActiveSupIdx(null) }}
              className="px-4 py-2 bg-blue-600 text-white text-sm font-semibold rounded-xl hover:bg-blue-700"
            >
              + New Plan
            </button>
          </div>
          {plans.length === 0 ? (
            <div className="text-center py-20 text-gray-400">
              <p className="text-4xl mb-3">📦</p>
              <p className="text-lg font-medium">No plans yet</p>
              <p className="text-sm mt-1">Create a new plan to get started</p>
            </div>
          ) : (
            <div className="space-y-2">
              {plans.map(p => (
                <div key={p.id} onClick={() => { setPlan(p); setActiveSupIdx(null) }}
                  className="border border-gray-200 rounded-xl p-4 cursor-pointer hover:border-blue-400 hover:shadow-sm transition-all bg-white">
                  <div className="font-semibold text-gray-900">{p.name}</div>
                  <div className="text-xs text-gray-500 mt-1 flex flex-wrap gap-4">
                    <span>{p.type_1_name}: {(p.forecast_1 ?? 0).toLocaleString()} branches</span>
                    <span>{p.type_2_name}: {(p.forecast_2 ?? 0).toLocaleString()} branches</span>
                    <span>{(p.items ?? []).length} items</span>
                    <span>{(p.suppliers ?? []).length} suppliers</span>
                  </div>
                </div>
              ))}
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
            onClick={savePlan}
            disabled={saving}
            className={`ml-auto px-5 py-2 text-sm font-semibold rounded-xl transition-colors disabled:opacity-50 ${saved ? 'bg-green-600 text-white' : 'bg-blue-600 hover:bg-blue-700 text-white'}`}
          >
            {saving ? 'Saving…' : saved ? 'Saved ✓' : 'Save'}
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
                    type="number" min="0"
                    value={n === 1 ? plan.forecast_1 : plan.forecast_2}
                    onChange={e => setPlan(p => p ? { ...p, [`forecast_${n}`]: Number(e.target.value) || 0 } : p)}
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
            <p className="text-[11px] text-gray-400 font-mono mb-3">A: NO (skip) · B: item_code · C: description · D: qty per branch · "-" = 0</p>
            <div className="flex flex-col gap-2">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => templateRef1.current?.click()}
                  className="px-3 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 font-medium"
                >
                  Upload {plan.type_1_name} Template
                </button>
                <input ref={templateRef1} type="file" accept=".xlsx,.xls" className="hidden"
                  onChange={e => { const f = e.target.files?.[0]; if (f) handleTemplateUpload(f, 1); e.target.value = '' }} />
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => templateRef2.current?.click()}
                  className="px-3 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 font-medium"
                >
                  Upload {plan.type_2_name} Template
                </button>
                <input ref={templateRef2} type="file" accept=".xlsx,.xls" className="hidden"
                  onChange={e => { const f = e.target.files?.[0]; if (f) handleTemplateUpload(f, 2); e.target.value = '' }} />
              </div>
              {plan.items.length > 0 && (
                <span className="text-sm text-green-600 font-semibold">{plan.items.length} items loaded</span>
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
                    isActive={activeSupIdx === si}
                    poUploads={poUploads}
                    showPoSelector={showPoSelector === sup.id}
                    onToggleActive={() => { setActiveSupIdx(v => v === si ? null : si); setShowPoSelector(null) }}
                    onUpdate={upd => updateSup(sup.id, upd)}
                    onUploadContainerQty={f => handleContainerQtyUpload(f, sup.id)}
                    onTogglePoSelector={() => setShowPoSelector(v => v === sup.id ? null : sup.id)}
                    onRemove={() => {
                      setPlan(p => p ? { ...p, suppliers: p.suppliers.filter((_, i) => i !== si) } : p)
                      if (activeSupIdx === si) setActiveSupIdx(null)
                      else if (activeSupIdx !== null && activeSupIdx > si) setActiveSupIdx(v => v !== null ? v - 1 : null)
                    }}
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
                    <th className="text-left px-3 py-3 font-semibold whitespace-nowrap">Item Code</th>
                    <th className="text-left px-3 py-3 font-semibold whitespace-nowrap">Description</th>
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap">
                      Assumed Qty<br /><span className="font-normal opacity-70 text-[10px]">{plan.type_1_name} branch type</span>
                    </th>
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap">
                      Assumed Qty<br /><span className="font-normal opacity-70 text-[10px]">{plan.type_2_name} branch type</span>
                    </th>
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap">Sum Assumed<br />Qty</th>
                    {activeSup && (
                      <>
                        <th className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-blue-900 border-l-2 border-blue-500">
                          Full Container<br />QTY
                        </th>
                        <th className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-blue-900">
                          PO QTY
                        </th>
                        {activeSup.load_dates.map(d => (
                          <th key={d} className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-indigo-900">
                            Load<br />{fmtDate(d)}
                          </th>
                        ))}
                      </>
                    )}
                    <th className="text-right px-3 py-3 font-semibold whitespace-nowrap bg-rose-900">LEFT</th>
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
                        <td className={`px-3 py-2 font-mono font-semibold text-gray-900`}>{item.item_code}</td>
                        <td className="px-3 py-2 text-gray-600 max-w-[280px] truncate">{item.description || '—'}</td>
                        <td className="px-3 py-2 text-right tabular-nums text-gray-700">{a1.toLocaleString()}</td>
                        <td className="px-3 py-2 text-right tabular-nums text-gray-700">{item.qty_2 > 0 ? a2.toLocaleString() : '—'}</td>
                        <td className="px-3 py-2 text-right tabular-nums font-semibold text-gray-900">{sumA.toLocaleString()}</td>
                        {activeSup && (
                          <>
                            <td className="px-3 py-2 text-right tabular-nums text-blue-900 bg-blue-50 border-l-2 border-blue-200">
                              {activeSup.container_qtys[item.item_code]?.toLocaleString() ?? '—'}
                            </td>
                            <td className="px-3 py-2 text-right tabular-nums text-blue-900 bg-blue-50">
                              {(poQtyMap.get(item.item_code) ?? 0) > 0 ? (poQtyMap.get(item.item_code) ?? 0).toLocaleString() : '—'}
                            </td>
                            {activeSup.load_dates.map(d => {
                              const entry = activeSup.loads.find(l => l.date === d && l.item_code === item.item_code)
                              return (
                                <td key={d} className="px-1.5 py-1 bg-indigo-50">
                                  <input
                                    type="number" min="0"
                                    value={entry?.qty ?? ''}
                                    placeholder="0"
                                    onChange={e => updateLoadQty(activeSup.id, d, item.item_code, Number(e.target.value))}
                                    className="w-20 text-right px-2 py-0.5 border border-indigo-200 rounded focus:outline-none focus:border-indigo-500 tabular-nums bg-white text-xs"
                                  />
                                </td>
                              )
                            })}
                          </>
                        )}
                        <td className={`px-3 py-2 text-right tabular-nums font-bold ${left < 0 ? 'text-red-600 bg-red-50' : left === 0 ? 'text-green-700 bg-green-50' : 'text-gray-900'}`}>
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

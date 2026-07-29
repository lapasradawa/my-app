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

export default function OrderPlanPage() {
  const [projects, setProjects] = useState<string[]>([])
  const [selectedProject, setSelectedProject] = useState('')
  const [settings, setSettings] = useState<Settings>({ cny_rate: 4.85, usd_rate: 33.00, ddp_multiplier: 1.11 })
  const [rows, setRows] = useState<PlanRow[]>([])
  const [parsedCache, setParsedCache] = useState<ParsedRow[]>([])
  const [allSuppliers, setAllSuppliers] = useState<string[]>([])
  const [loading, setLoading] = useState(false)
  const [fileName, setFileName] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    loadProjects()
    loadSettings()
  }, [])

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

  async function loadProjects() {
    const { data } = await supabase.from('po_items').select('project')
    if (data) {
      setProjects([...new Set((data as { project: string }[]).map(r => r.project))].sort())
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

    // Find header row: look for row where col A is "#" or col B contains "item"/"code"
    let dataStart = 1
    for (let i = 0; i < Math.min(raw.length, 10); i++) {
      const row = raw[i] as unknown[]
      const colA = String(row[0] ?? '').trim()
      const colB = String(row[1] ?? '').toLowerCase()
      if (colA === '#' || colB.includes('item') || colB.includes('code')) {
        dataStart = i + 1
        break
      }
    }

    const parsed: ParsedRow[] = []
    for (let i = dataStart; i < raw.length; i++) {
      const row = raw[i] as unknown[]
      const itemCode = String(row[1] ?? '').trim()  // col B
      if (!itemCode || itemCode.toLowerCase().includes('total') || itemCode.toLowerCase().includes('รวม')) continue

      parsed.push({
        item_code: itemCode,
        description: String(row[2] ?? '').trim(),    // col C
        stock_thai: Number(row[3]) || 0,             // col D
        po_thai: Number(row[6]) || 0,                // col G
        lonsua: Number(row[12]) || 0,                // col M
        week1: Number(row[13]) || 0,                 // col N
        week2: Number(row[14]) || 0,                 // col O
        week3_4: Number(row[15]) || 0,               // col P
        next_month: Number(row[16]) || 0,            // col Q
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
      const { data } = await supabase
        .from('po_items')
        .select('item_code, supplier, fob_price, currency')
        .eq('project', project)
        .order('uploaded_at', { ascending: false })

      if (data) {
        for (const item of data as { item_code: string; supplier: string; fob_price: number; currency: string }[]) {
          supplierSet.add(item.supplier)
          if (!priceMap.has(item.item_code)) priceMap.set(item.item_code, new Map())
          const sup = priceMap.get(item.item_code)!
          if (!sup.has(item.supplier)) sup.set(item.supplier, { fob_price: item.fob_price, currency: item.currency })
        }
      }
    }

    setAllSuppliers([...supplierSet].sort())

    const { cny_rate, usd_rate, ddp_multiplier } = settings

    function toDdpThb(fob: number, currency: string) {
      return fob * (currency === 'USD' ? usd_rate : cny_rate) * ddp_multiplier
    }

    const planRows: PlanRow[] = parsed.map(r => {
      const supplierPrices = priceMap.get(r.item_code)
      const ddp_prices: DdpPrice[] = []
      if (supplierPrices) {
        for (const [supplier, { fob_price, currency }] of supplierPrices.entries()) {
          ddp_prices.push({ supplier, ddp_thb: toDdpThb(fob_price, currency) })
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

  async function handleProjectChange(project: string) {
    setSelectedProject(project)
    if (parsedCache.length > 0) {
      setLoading(true)
      await buildPlanRows(parsedCache, project)
      setLoading(false)
    }
  }

  function fmtN(n: number, dec = 0) {
    if (n === 0) return '—'
    return n.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec })
  }

  function fmtCalc(n: number) {
    return n.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })
  }

  const ddpCols = Math.min(allSuppliers.length, 5)

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
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">Order Plan (สั่งโหลด)</h1>
          <p className="text-sm text-gray-500 mt-1">อัพโหลด stock_dashboard Excel และเลือก Project เพื่อดูแผนการสั่งโหลด</p>
        </div>

        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mb-6">
          <div className="flex items-center gap-6 flex-wrap">
            <div>
              <label className="text-sm text-gray-600 font-medium block mb-1.5">Project (ราคา DDP)</label>
              <select
                value={selectedProject}
                onChange={e => handleProjectChange(e.target.value)}
                className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 bg-white min-w-[160px]">
                <option value="">— เลือก Project —</option>
                {projects.map(p => <option key={p} value={p}>{p}</option>)}
              </select>
            </div>
            <div>
              <label className="text-sm text-gray-600 font-medium block mb-1.5">stock_dashboard Excel</label>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => fileRef.current?.click()}
                  className="px-4 py-1.5 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-50 transition-colors">
                  เลือกไฟล์
                </button>
                {fileName && <span className="text-sm text-gray-500 truncate max-w-xs">{fileName}</span>}
                <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile} />
              </div>
            </div>
          </div>
          {rows.length > 0 && (
            <p className="mt-3 text-xs text-gray-400">{rows.length} items โหลดแล้ว{selectedProject ? ` · ราคา DDP จาก ${selectedProject}` : ' · ยังไม่ได้เลือก Project สำหรับราคา DDP'}</p>
          )}
        </div>

        {loading && (
          <div className="text-center py-16 text-gray-400 text-sm">กำลังโหลด...</div>
        )}

        {!loading && rows.length === 0 && (
          <div className="text-center py-16 text-gray-400 text-sm">อัพโหลด stock_dashboard เพื่อดูแผน</div>
        )}

        {!loading && rows.length > 0 && (
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="text-xs border-collapse">
                <thead>
                  <tr className="bg-gray-50 text-gray-500 border-b border-gray-200">
                    <th className="px-3 py-2.5 text-left whitespace-nowrap font-semibold sticky left-0 bg-gray-50 z-10 border-r border-gray-200">Item Code</th>
                    <th className="px-3 py-2.5 text-left whitespace-nowrap font-semibold">Description</th>
                    {Array.from({ length: ddpCols }, (_, i) => (
                      <th key={i} className="px-3 py-2.5 text-right whitespace-nowrap font-semibold text-blue-600">
                        DDP {i + 1}
                      </th>
                    ))}
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">J: PO ไทย</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">K: Stock ไทย</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-amber-50 text-amber-700">L: PO/2</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">M: ลงเรือ</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">N: W1</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">O: W2</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">P: W3+4</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold">Q: Next</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-green-50 text-green-700">S: เผื่อ W3W4</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-green-50 text-green-700">T: เผื่อ Next</th>
                    <th className="px-3 py-2.5 text-right whitespace-nowrap font-semibold bg-red-50 text-red-700">U: ต้องสั่ง</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row, i) => (
                    <tr key={i} className="border-b border-gray-100 hover:bg-blue-50/20">
                      <td className="px-3 py-2 font-mono text-gray-800 whitespace-nowrap sticky left-0 bg-white border-r border-gray-100 z-10">{row.item_code}</td>
                      <td className="px-3 py-2 text-gray-600 whitespace-nowrap max-w-[220px] overflow-hidden text-ellipsis">{row.description || <span className="text-gray-300">—</span>}</td>
                      {Array.from({ length: ddpCols }, (_, j) => {
                        const p = row.ddp_prices[j]
                        return (
                          <td key={j} className="px-3 py-2 text-right whitespace-nowrap">
                            {p ? (
                              <span>
                                <span className="text-gray-400">{p.supplier} </span>
                                <span className="text-gray-700 font-medium">{p.ddp_thb.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })}</span>
                              </span>
                            ) : <span className="text-gray-300">—</span>}
                          </td>
                        )
                      })}
                      <td className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap">{fmtN(row.po_thai)}</td>
                      <td className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap">{fmtN(row.stock_thai)}</td>
                      <td className="px-3 py-2 text-right text-gray-700 bg-amber-50/20 whitespace-nowrap">{fmtCalc(row.L)}</td>
                      <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.lonsua)}</td>
                      <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.week1)}</td>
                      <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.week2)}</td>
                      <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.week3_4)}</td>
                      <td className="px-3 py-2 text-right text-gray-700 whitespace-nowrap">{fmtN(row.next_month)}</td>
                      <td className="px-3 py-2 text-right font-medium bg-green-50/20 whitespace-nowrap text-gray-700">{fmtCalc(row.S)}</td>
                      <td className="px-3 py-2 text-right font-medium bg-green-50/20 whitespace-nowrap text-gray-700">{fmtCalc(row.T)}</td>
                      <td className={`px-3 py-2 text-right font-semibold bg-red-50/20 whitespace-nowrap ${row.U < 0 ? 'text-red-600' : 'text-gray-700'}`}>
                        {fmtCalc(row.U)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

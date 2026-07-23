'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { supabase } from '@/lib/supabase'
import LockButton from '@/components/LockButton'

interface PoBuilderItem {
  item_code: string
  description: string
  fob_price: number | null
  currency: string
  qty: string
}

export default function PoBuilderPage() {
  const [allSuppliers, setAllSuppliers] = useState<string[]>([])
  const [poSupplier, setPoSupplier] = useState('')
  const [poItemInput, setPoItemInput] = useState('')
  const [poItems, setPoItems] = useState<PoBuilderItem[]>([])
  const [poAdding, setPoAdding] = useState(false)
  const [, setUnlocked] = useState(false)

  useEffect(() => { loadSuppliers() }, [])

  async function loadSuppliers() {
    const { data } = await supabase.from('po_items').select('supplier')
    if (data) {
      setAllSuppliers([...new Set((data as { supplier: string }[]).map(r => r.supplier))].sort())
    }
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
        .select('fob_price, currency')
        .ilike('item_code', item.item_code)
        .eq('supplier', newSupplier)
        .order('uploaded_at', { ascending: false })
        .limit(1)
      const found = data?.[0] as { fob_price: number; currency: string } | undefined
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
    ws['!cols'] = [{ wch: 5 }, { wch: 27 }, { wch: 60 }, { wch: 10 }, { wch: 15 }, { wch: 15 }]

    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'PO')
    XLSX.writeFile(wb, `PO_${poSupplier}_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  function fmtN(n: number) {
    return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-6 text-sm sticky top-0 z-20 shadow-sm flex-wrap">
        <span className="font-bold text-gray-900">Import PO</span>
        <Link href="/" className="text-gray-500 hover:text-gray-800 transition-colors">PO Matching</Link>
        <Link href="/dashboard" className="text-gray-500 hover:text-gray-800 transition-colors">Dashboard</Link>
        <Link href="/calendar" className="text-gray-500 hover:text-gray-800 transition-colors">Calendar</Link>
        <Link href="/report" className="text-gray-500 hover:text-gray-800 transition-colors">Report</Link>
        <Link href="/compare" className="text-gray-500 hover:text-gray-800 transition-colors">Cost Compare</Link>
        <Link href="/po-builder" className="text-blue-600">PO Builder</Link>
        <Link href="/guide" className="text-gray-500 hover:text-gray-800 transition-colors">Guide</Link>
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
        <div className="ml-auto">
          <LockButton onUnlock={() => setUnlocked(true)} onLock={() => setUnlocked(false)} />
        </div>
      </nav>

      <div className="max-w-4xl mx-auto px-6 py-8">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">PO Builder</h1>
          <p className="text-sm text-gray-500 mt-1">เลือก Supplier และเพิ่ม Item Code จากทุก Project เพื่อสร้าง PO</p>
        </div>

        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
          <div className="flex items-center gap-3 mb-5 flex-wrap">
            <label className="text-sm text-gray-600 font-medium whitespace-nowrap">Supplier:</label>
            <select
              value={poSupplier}
              onChange={e => changePoSupplier(e.target.value)}
              className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 bg-white min-w-[180px]">
              <option value="">— เลือก Supplier —</option>
              {allSuppliers.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>

          {!poSupplier ? (
            <p className="text-sm text-gray-400 text-center py-8">เลือก Supplier เพื่อเริ่มสร้าง PO</p>
          ) : (
            <>
              <div className="flex items-center gap-2 mb-5">
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
                <p className="text-sm text-gray-400 text-center py-8">เพิ่ม Item Code เพื่อสร้าง PO</p>
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
                            {item.fob_price !== null
                              ? `${fmtN(item.fob_price)} ${item.currency}`
                              : <span className="text-gray-300">ไม่พบราคา</span>}
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
    </div>
  )
}

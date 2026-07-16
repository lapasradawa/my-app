'use client'

import { useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import { useParams, useRouter } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import { exportQCReportExcel, type QCItem } from '@/lib/qc-excel'

const NAV = [
  { href: '/', label: 'PO Matching' }, { href: '/dashboard', label: 'Dashboard' },
  { href: '/calendar', label: 'Calendar' }, { href: '/report', label: 'Report' },
  { href: '/summary', label: 'Summary' }, { href: '/compare', label: 'Cost Compare' },
  { href: '/po-builder', label: 'PO Builder' }, { href: '/qc', label: 'QC Report' },
  { href: '/guide', label: 'Guide' },
]

const CA_OPTIONS = ['REPLACEMENT IN NEXT SHIPMENT', 'CREDIT NOTE / REFUND', 'REWORK / REPAIR', 'OTHER']

interface QCReport {
  id: string; report_no: string; supplier_company: string; destuffing_date: string
  issue_found_date: string; invoice_no: string; po_no: string; attachment_desc: string
  description: string; items: QCItem[]; corrective_actions: string[]; corrective_action_comment: string
  root_cause: string; preventive_action: string; verification_accepted: boolean | null
  verification_comment: string; status: string; closure_file_url: string | null
}

const EMPTY: Omit<QCReport, 'id'> = {
  report_no: '', supplier_company: '', destuffing_date: '', issue_found_date: '',
  invoice_no: '', po_no: '', attachment_desc: 'PO, Photos', description: '',
  items: [], corrective_actions: [], corrective_action_comment: '',
  root_cause: '', preventive_action: '', verification_accepted: null,
  verification_comment: '', status: 'open', closure_file_url: null,
}

const EMPTY_ITEM: QCItem = { item_code: '', product_description: '', qty: 0, unit_price: 0, total: 0, qty_defective: 0, remark: '' }

function SectionHeader({ title }: { title: string }) {
  return (
    <div className="bg-gray-100 border border-gray-300 px-4 py-2 text-sm font-bold text-gray-800 rounded-t-lg mt-6">
      {title}
    </div>
  )
}

export default function QCDetailPage() {
  const { id } = useParams<{ id: string }>()
  const router = useRouter()
  const closureRef = useRef<HTMLInputElement>(null)
  const [form, setForm] = useState<Omit<QCReport, 'id'>>(EMPTY)
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [uploadingClosure, setUploadingClosure] = useState(false)
  const [poItems, setPoItems] = useState<{ item_code: string; description: string; fob_price: number | null; currency: string | null }[]>([])
  const [itemSearch, setItemSearch] = useState<string[]>([])
  const [openDropdown, setOpenDropdown] = useState<number | null>(null)

  useEffect(() => {
    supabase.from('qc_reports').select('*').eq('id', id).single().then(({ data }) => {
      if (data) setForm(data as Omit<QCReport, 'id'>)
      setLoading(false)
    })
    supabase.from('po_items').select('item_code, description, fob_price, currency')
      .order('item_code').then(({ data }) => {
        if (data) {
          const unique = Array.from(new Map(data.map(r => [r.item_code, r])).values())
          setPoItems(unique)
          setItemSearch(unique.map(r => r.item_code))
        }
      })
  }, [id])

  function set<K extends keyof typeof form>(k: K, v: typeof form[K]) {
    setForm(prev => ({ ...prev, [k]: v }))
    setSaved(false)
  }

  function setItem(i: number, k: keyof QCItem, v: string | number) {
    const items = [...form.items]
    items[i] = { ...items[i], [k]: v }
    if (k === 'qty' || k === 'unit_price') {
      const q = k === 'qty' ? Number(v) : items[i].qty
      const p = k === 'unit_price' ? Number(v) : items[i].unit_price
      items[i].total = Math.round(q * p * 100) / 100
    }
    set('items', items)
  }

  function addItem() { set('items', [...form.items, { ...EMPTY_ITEM }]) }
  function removeItem(i: number) { set('items', form.items.filter((_, j) => j !== i)) }

  function toggleCA(opt: string) {
    const arr = form.corrective_actions.includes(opt)
      ? form.corrective_actions.filter(x => x !== opt)
      : [...form.corrective_actions, opt]
    set('corrective_actions', arr)
  }

  async function save() {
    setSaving(true)
    await supabase.from('qc_reports').update({ ...form, updated_at: new Date().toISOString() }).eq('id', id)
    setSaved(true)
    setSaving(false)
  }

  async function uploadClosure(file: File) {
    setUploadingClosure(true)
    const path = `qc-closure/${id}-${Date.now()}.${file.name.split('.').pop()}`
    const { error } = await supabase.storage.from('payment-proofs').upload(path, file, { upsert: true })
    if (!error) {
      const { data: urlData } = supabase.storage.from('payment-proofs').getPublicUrl(path)
      const url = urlData.publicUrl
      await supabase.from('qc_reports').update({ closure_file_url: url, status: 'closed', updated_at: new Date().toISOString() }).eq('id', id)
      set('closure_file_url', url)
      set('status', 'closed')
    } else {
      alert('อัปโหลดไม่สำเร็จ: ' + error.message)
    }
    setUploadingClosure(false)
  }

  async function handleExport() {
    setExporting(true)
    await exportQCReportExcel({ ...form })
    setExporting(false)
  }

  async function deleteReport() {
    if (!confirm(`ลบ Report ${form.report_no} ใช่ไหม? ไม่สามารถกู้คืนได้`)) return
    await supabase.from('qc_reports').delete().eq('id', id)
    router.push('/qc')
  }

  const inputCls = 'border border-gray-300 rounded-lg px-3 py-2 text-sm outline-none focus:border-blue-400 w-full bg-white'
  const labelCls = 'block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1'

  if (loading) return <div className="min-h-screen bg-gray-50 flex items-center justify-center text-gray-400 text-sm">กำลังโหลด...</div>

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-5 text-sm flex-wrap">
        <span className="font-bold text-gray-900 mr-2">Import PO</span>
        {NAV.map(n => (
          <Link key={n.href} href={n.href}
            className={n.href === '/qc' ? 'text-blue-600 font-semibold' : 'text-gray-500 hover:text-gray-800 transition-colors'}>
            {n.label}
          </Link>
        ))}
      </nav>

      <div className="max-w-5xl mx-auto px-6 py-6">
        {/* Top bar */}
        <div className="flex items-center justify-between mb-4 flex-wrap gap-3">
          <div className="flex items-center gap-3">
            <Link href="/qc" className="text-sm text-gray-500 hover:text-gray-800">← QC Report</Link>
            <span className="text-gray-300">|</span>
            <h1 className="text-lg font-bold text-gray-900 font-mono">{form.report_no}</h1>
            <span className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold ${form.status === 'closed' ? 'bg-green-100 text-green-700' : 'bg-orange-100 text-orange-700'}`}>
              {form.status === 'closed' ? '✓ ปิดจบแล้ว' : '● เปิดอยู่'}
            </span>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={handleExport} disabled={exporting}
              className="bg-green-600 hover:bg-green-700 text-white text-sm font-semibold px-4 py-2 rounded-lg transition-colors disabled:opacity-50">
              {exporting ? 'กำลัง Export...' : '↓ Export Excel'}
            </button>
            <button onClick={save} disabled={saving}
              className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold px-4 py-2 rounded-lg transition-colors disabled:opacity-50">
              {saving ? 'กำลังบันทึก...' : saved ? '✓ บันทึกแล้ว' : 'บันทึก'}
            </button>
            <button onClick={deleteReport} className="text-xs text-red-400 hover:text-red-600 px-2 py-2">ลบ</button>
          </div>
        </div>

        {/* Header fields */}
        <div className="bg-white rounded-xl border border-gray-200 overflow-hidden mb-1">
          {/* Fixed info row */}
          <div className="grid grid-cols-3 divide-x divide-gray-100 bg-gray-50 border-b border-gray-200 text-xs">
            <div className="px-4 py-2.5">
              <span className="text-gray-400 font-semibold uppercase tracking-wide text-[10px]">Subject</span>
              <div className="font-semibold text-gray-700 mt-0.5">QUALITY CLAIM</div>
            </div>
            <div className="px-4 py-2.5">
              <span className="text-gray-400 font-semibold uppercase tracking-wide text-[10px]">Customer Company</span>
              <div className="font-semibold text-gray-700 mt-0.5">RETAIL BUSINESS SOLUTION CO., LTD</div>
            </div>
            <div className="px-4 py-2.5">
              <span className="text-gray-400 font-semibold uppercase tracking-wide text-[10px]">Address</span>
              <div className="text-gray-600 mt-0.5">387 SUKHONTHASAWAT RD., LADPRAO, LADPRAO, BANGKOK, THAILAND 10230</div>
            </div>
          </div>
          {/* Editable fields */}
          <div className="p-5 grid grid-cols-3 gap-4">
            <div>
              <label className={labelCls}>Report No.</label>
              <input className={inputCls} value={form.report_no} onChange={e => set('report_no', e.target.value)} />
            </div>
            <div>
              <label className={labelCls}>Supplier Company</label>
              <input className={inputCls} value={form.supplier_company} onChange={e => set('supplier_company', e.target.value)}
                placeholder="SHANGHAI YONGGUAN COMMERCIAL EQUIPMENT CO., LTD" />
            </div>
            <div>
              <label className={labelCls}>Invoice No.</label>
              <input className={inputCls} value={form.invoice_no} onChange={e => set('invoice_no', e.target.value)} placeholder="YG260039" />
            </div>
            <div>
              <label className={labelCls}>PO No.</label>
              <input className={inputCls} value={form.po_no} onChange={e => set('po_no', e.target.value)} placeholder="RBSYG04-GEN5" />
            </div>
            <div>
              <label className={labelCls}>Destuffing Date</label>
              <input type="date" className={inputCls} value={form.destuffing_date} onChange={e => set('destuffing_date', e.target.value)} />
            </div>
            <div>
              <label className={labelCls}>Issue Found Date</label>
              <input type="date" className={inputCls} value={form.issue_found_date} onChange={e => set('issue_found_date', e.target.value)} />
            </div>
          </div>
        </div>

        {/* Part 1: ISSUE */}
        <SectionHeader title="Part 1 : ISSUE" />
        <div className="bg-white border border-gray-300 border-t-0 rounded-b-lg p-5 mb-1">
          <label className={labelCls}>Description</label>
          <textarea className={`${inputCls} min-h-[80px] resize-y`} value={form.description}
            onChange={e => set('description', e.target.value)} placeholder="อธิบายปัญหาที่พบ..." />

          <div className="mt-4 overflow-x-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-gray-100">
                  {['NO.', 'ITEM CODE', 'PRODUCT DESCRIPTION', 'QTY (PCS)', 'UNIT PRICE', 'TOTAL', 'QTY DEFECTIVE', 'REMARK', ''].map(h => (
                    <th key={h} className="border border-gray-300 px-2 py-1.5 text-left font-semibold text-gray-600 whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {form.items.map((item, i) => (
                  <tr key={i}>
                    <td className="border border-gray-200 px-2 py-1 text-center text-gray-500">{i + 1}</td>
                    {/* item_code combobox */}
                    <td className="border border-gray-200 p-0 relative">
                      <input
                        className="w-full px-2 py-1.5 text-xs outline-none focus:bg-blue-50 min-w-[140px]"
                        value={item.item_code}
                        onChange={e => {
                          setItem(i, 'item_code', e.target.value)
                          const q = e.target.value.toLowerCase()
                          setItemSearch(poItems.filter(p => p.item_code.toLowerCase().includes(q)).map(p => p.item_code))
                          setOpenDropdown(i)
                        }}
                        onFocus={() => {
                          setItemSearch(poItems.map(p => p.item_code))
                          setOpenDropdown(i)
                        }}
                        onBlur={() => setTimeout(() => setOpenDropdown(null), 150)}
                        placeholder="item code…"
                      />
                      {openDropdown === i && itemSearch.length > 0 && (
                        <ul className="absolute z-50 left-0 top-full max-h-48 overflow-y-auto bg-white border border-gray-300 shadow-lg rounded text-xs w-56">
                          {itemSearch.slice(0, 50).map(code => {
                            const p = poItems.find(x => x.item_code === code)!
                            return (
                              <li key={code}
                                className="px-2 py-1.5 hover:bg-blue-50 cursor-pointer border-b border-gray-100 last:border-0"
                                onMouseDown={() => {
                                  setItem(i, 'item_code', code)
                                  if (p.description) setItem(i, 'product_description', p.description)
                                  if (p.fob_price != null) setItem(i, 'unit_price', p.fob_price)
                                  setOpenDropdown(null)
                                }}>
                                <div className="font-mono font-semibold text-gray-800">{code}</div>
                                {p.description && <div className="text-gray-400 truncate max-w-[200px]">{p.description}</div>}
                              </li>
                            )
                          })}
                        </ul>
                      )}
                    </td>
                    {/* product_description */}
                    <td className="border border-gray-200 p-0">
                      <input className="w-full px-2 py-1.5 text-xs outline-none focus:bg-blue-50 min-w-[120px]"
                        value={item.product_description} onChange={e => setItem(i, 'product_description', e.target.value)} />
                    </td>
                    {(['qty', 'unit_price', 'total', 'qty_defective'] as const).map(k => (
                      <td key={k} className="border border-gray-200 p-0">
                        <input type="number" className="w-full px-2 py-1.5 text-xs outline-none focus:bg-blue-50 text-right min-w-[70px]"
                          value={item[k] || ''} onChange={e => setItem(i, k, parseFloat(e.target.value) || 0)}
                          readOnly={k === 'total'} style={k === 'total' ? { background: '#f9f9f9' } : {}} />
                      </td>
                    ))}
                    <td className="border border-gray-200 p-0">
                      <input className="w-full px-2 py-1.5 text-xs outline-none focus:bg-blue-50 min-w-[100px]"
                        value={item.remark} onChange={e => setItem(i, 'remark', e.target.value)} />
                    </td>
                    <td className="border border-gray-200 px-1 text-center">
                      <button onClick={() => removeItem(i)} className="text-red-400 hover:text-red-600 text-xs px-1">✕</button>
                    </td>
                  </tr>
                ))}
                {form.items.length > 0 && (
                  <tr className="bg-gray-50 font-semibold text-xs">
                    <td colSpan={3} className="border border-gray-300 px-2 py-1.5 text-center">Total</td>
                    <td className="border border-gray-300 px-2 py-1.5 text-right">{form.items.reduce((s, i) => s + (i.qty || 0), 0).toLocaleString()}</td>
                    <td className="border border-gray-300"></td>
                    <td className="border border-gray-300 px-2 py-1.5 text-right">{form.items.reduce((s, i) => s + (i.total || 0), 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                    <td className="border border-gray-300 px-2 py-1.5 text-right">{form.items.reduce((s, i) => s + (i.qty_defective || 0), 0).toLocaleString()}</td>
                    <td colSpan={2} className="border border-gray-300"></td>
                  </tr>
                )}
              </tbody>
            </table>
            <button onClick={addItem} className="mt-2 text-xs text-blue-600 hover:text-blue-800 font-medium">+ เพิ่ม Item</button>
          </div>
        </div>

        {/* Part 2: CORRECTIVE ACTION */}
        <SectionHeader title="Part 2 : CORRECTIVE ACTION" />
        <div className="bg-white border border-gray-300 border-t-0 rounded-b-lg p-5 mb-1">
          <div className="grid grid-cols-2 gap-6">
            <div>
              <label className={labelCls}>Corrective Action</label>
              <div className="flex flex-col gap-2 mt-1">
                {CA_OPTIONS.map(opt => (
                  <label key={opt} className="flex items-center gap-2 text-sm cursor-pointer">
                    <input type="checkbox" checked={form.corrective_actions.includes(opt)}
                      onChange={() => toggleCA(opt)} className="w-4 h-4 rounded" />
                    <span className="text-gray-700">{opt}</span>
                  </label>
                ))}
              </div>
            </div>
            <div>
              <label className={labelCls}>Description / Comment</label>
              <textarea className={`${inputCls} min-h-[120px] resize-y`} value={form.corrective_action_comment}
                onChange={e => set('corrective_action_comment', e.target.value)} placeholder="รายละเอียดการแก้ไข..." />
            </div>
          </div>
        </div>

        {/* Part 3: PREVENTIVE ACTION */}
        <SectionHeader title="Part 3 : PREVENTIVE ACTION (FILLED BY SUPPLIER)" />
        <div className="bg-white border border-gray-300 border-t-0 rounded-b-lg p-5 mb-1">
          <div className="grid grid-cols-2 gap-6">
            <div>
              <label className={labelCls}>Root Cause</label>
              <textarea className={`${inputCls} min-h-[100px] resize-y`} value={form.root_cause}
                onChange={e => set('root_cause', e.target.value)} placeholder="สาเหตุของปัญหา..." />
            </div>
            <div>
              <label className={labelCls}>Action</label>
              <textarea className={`${inputCls} min-h-[100px] resize-y`} value={form.preventive_action}
                onChange={e => set('preventive_action', e.target.value)} placeholder="มาตรการป้องกัน..." />
            </div>
          </div>
        </div>

        {/* Part 4: VERIFICATION */}
        <SectionHeader title="Part 4 : VERIFICATION STATUS (FILLED BY RBS)" />
        <div className="bg-white border border-gray-300 border-t-0 rounded-b-lg p-5 mb-1">
          <div className="flex items-center gap-6">
            <label className="flex items-center gap-2 cursor-pointer">
              <input type="radio" name="verification" checked={form.verification_accepted === true}
                onChange={() => set('verification_accepted', true)} className="w-4 h-4" />
              <span className="text-sm font-semibold text-green-700">☑ ACCEPTED</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input type="radio" name="verification" checked={form.verification_accepted === false}
                onChange={() => set('verification_accepted', false)} className="w-4 h-4" />
              <span className="text-sm font-semibold text-red-600">☐ NOT ACCEPTED</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input type="radio" name="verification" checked={form.verification_accepted === null}
                onChange={() => set('verification_accepted', null)} className="w-4 h-4" />
              <span className="text-sm text-gray-500">รอตรวจสอบ</span>
            </label>
          </div>
          {form.verification_accepted === false && (
            <div className="mt-3">
              <label className={labelCls}>Comment</label>
              <input className={inputCls} value={form.verification_comment}
                onChange={e => set('verification_comment', e.target.value)} placeholder="เหตุผลที่ไม่ยอมรับ..." />
            </div>
          )}
        </div>

        {/* Status + Closure */}
        <div className="bg-white rounded-xl border border-gray-200 p-5 mt-6">
          <div className="flex items-start justify-between flex-wrap gap-4">
            <div>
              <div className="text-sm font-semibold text-gray-700 mb-3">สถานะการปิดจบงาน</div>
              <div className="flex items-center gap-3">
                <button
                  onClick={() => set('status', form.status === 'closed' ? 'open' : 'closed')}
                  className={`px-4 py-2 rounded-lg text-sm font-semibold border transition-colors ${form.status === 'closed' ? 'bg-green-600 text-white border-green-600' : 'bg-white text-gray-600 border-gray-300 hover:border-green-400'}`}>
                  {form.status === 'closed' ? '✓ ปิดจบแล้ว' : 'ทำเครื่องหมายว่าปิดจบ'}
                </button>
                {form.status === 'open' && <span className="text-xs text-gray-400">หรืออัปโหลดหลักฐานเพื่อปิดอัตโนมัติ</span>}
              </div>
            </div>
            <div>
              <div className="text-sm font-semibold text-gray-700 mb-3">หลักฐานการปิดจบ (PDF)</div>
              {form.closure_file_url ? (
                <div className="flex flex-col gap-1">
                  <a href={form.closure_file_url} target="_blank" rel="noopener noreferrer"
                    className="text-sm text-blue-600 hover:text-blue-800 font-medium">📄 ดู / ดาวน์โหลดหลักฐาน</a>
                  <button onClick={() => closureRef.current?.click()} className="text-xs text-gray-400 hover:text-gray-600 text-left">
                    {uploadingClosure ? 'กำลังอัปโหลด...' : 'เปลี่ยนไฟล์'}
                  </button>
                </div>
              ) : (
                <button onClick={() => closureRef.current?.click()} disabled={uploadingClosure}
                  className="flex items-center gap-2 border-2 border-dashed border-gray-200 rounded-lg px-4 py-2.5 text-sm text-gray-400 hover:border-blue-300 hover:text-blue-500 transition-colors disabled:opacity-50">
                  {uploadingClosure ? 'กำลังอัปโหลด...' : '+ อัปโหลดหลักฐาน (PDF)'}
                </button>
              )}
              <input ref={closureRef} type="file" accept=".pdf,.jpg,.jpeg,.png" className="hidden"
                onChange={e => { const f = e.target.files?.[0]; if (f) uploadClosure(f); e.target.value = '' }} />
            </div>
          </div>
        </div>

        {/* Bottom save */}
        <div className="flex justify-end mt-4 pb-8">
          <button onClick={save} disabled={saving}
            className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold px-6 py-2.5 rounded-lg transition-colors disabled:opacity-50">
            {saving ? 'กำลังบันทึก...' : saved ? '✓ บันทึกแล้ว' : 'บันทึก'}
          </button>
        </div>
      </div>
    </div>
  )
}

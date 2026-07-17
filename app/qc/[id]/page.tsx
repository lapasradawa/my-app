'use client'

import { useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import { useParams, useRouter } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import { exportQCReportExcel, type QCItem } from '@/lib/qc-excel'
import { isUnlocked, tryUnlock } from '@/lib/auth'

function PasswordGate({ onSuccess, onCancel }: { onSuccess: () => void; onCancel: () => void }) {
  const [pw, setPw] = useState('')
  const [err, setErr] = useState(false)
  function attempt() {
    if (tryUnlock(pw)) { onSuccess() }
    else { setErr(true); setPw('') }
  }
  return (
    <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center">
      <div className="bg-white rounded-2xl shadow-xl p-6 w-80">
        <h3 className="font-semibold text-gray-800 mb-3">ใส่รหัสผ่านเพื่อดำเนินการ</h3>
        <input autoFocus type="password" value={pw}
          onChange={e => { setPw(e.target.value); setErr(false) }}
          onKeyDown={e => e.key === 'Enter' && attempt()}
          placeholder="รหัสผ่าน"
          className={`w-full border rounded-lg px-3 py-2 text-sm outline-none mb-2 ${err ? 'border-red-400' : 'border-gray-300 focus:border-blue-400'}`} />
        {err && <p className="text-red-500 text-xs mb-2">รหัสผ่านไม่ถูกต้อง</p>}
        <div className="flex gap-2 justify-end">
          <button onClick={onCancel} className="px-3 py-1.5 text-sm text-gray-600 hover:bg-gray-100 rounded-lg">ยกเลิก</button>
          <button onClick={attempt} className="px-4 py-1.5 text-sm bg-blue-600 text-white rounded-lg hover:bg-blue-700">ยืนยัน</button>
        </div>
      </div>
    </div>
  )
}

const NAV = [
  { href: '/', label: 'PO Matching' }, { href: '/dashboard', label: 'Dashboard' },
  { href: '/calendar', label: 'Calendar' }, { href: '/report', label: 'Report' },
  { href: '/summary', label: 'Summary' }, { href: '/compare', label: 'Cost Compare' },
  { href: '/po-builder', label: 'PO Builder' }, { href: '/guide', label: 'Guide' },
]

const CA_OPTIONS = ['REPLACEMENT IN NEXT SHIPMENT', 'CREDIT NOTE / REFUND', 'REWORK / REPAIR', 'OTHER']

interface QCReport {
  id: string; report_no: string; supplier_company: string; destuffing_date: string
  issue_found_date: string; invoice_no: string; po_no: string; attachment_desc: string
  description: string; items: QCItem[]; corrective_actions: string[]; corrective_action_comment: string
  root_cause: string; preventive_action: string; verification_accepted: boolean | null
  verification_comment: string; status: string; closure_file_url: string | null
  photo_urls: string[]; issue_types: string[]
}

const EMPTY: Omit<QCReport, 'id'> = {
  report_no: '', supplier_company: '', destuffing_date: '', issue_found_date: '',
  invoice_no: '', po_no: '', attachment_desc: 'PO, Photos', description: '',
  items: [], corrective_actions: [], corrective_action_comment: '',
  root_cause: '', preventive_action: '', verification_accepted: null,
  verification_comment: '', status: 'open', closure_file_url: null,
  photo_urls: [], issue_types: [],
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
  const photoRef = useRef<HTMLInputElement>(null)
  const [uploadingPhotos, setUploadingPhotos] = useState(false)
  const [form, setForm] = useState<Omit<QCReport, 'id'>>(EMPTY)
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [uploadingClosure, setUploadingClosure] = useState(false)
  const [showPW, setShowPW] = useState(false)
  const [pwCallback, setPwCallback] = useState<(() => void) | null>(null)

  function requireUnlock(action: () => void) {
    if (isUnlocked()) { action() }
    else { setPwCallback(() => action); setShowPW(true) }
  }
  const [poItems, setPoItems] = useState<{ item_code: string; description: string; fob_price: number | null; currency: string | null }[]>([])
  const [invoicesData, setInvoicesData] = useState<{ id: string; invoice_no: string; supplier: string | null; rows: { code?: string }[] }[]>([])
  const [suppliers, setSuppliers] = useState<string[]>([])
  const [itemSearch, setItemSearch] = useState<string[]>([])
  const [openDropdown, setOpenDropdown] = useState<number | null>(null)
  const [openSupplierDrop, setOpenSupplierDrop] = useState(false)
  const [supplierSearch, setSupplierSearch] = useState('')
  const [issueTypeOptions, setIssueTypeOptions] = useState<string[]>([])
  const [issueTypeInput, setIssueTypeInput] = useState('')
  const [showIssueTypeDrop, setShowIssueTypeDrop] = useState(false)

  useEffect(() => {
    supabase.from('qc_reports').select('*').eq('id', id).single().then(({ data }) => {
      if (data) setForm({ ...data, issue_types: data.issue_types ?? [] } as Omit<QCReport, 'id'>)
      setLoading(false)
    })
    supabase.from('po_items').select('item_code, description, fob_price, currency')
      .order('item_code').then(({ data }) => {
        if (data) {
          const unique = Array.from(new Map(data.map(r => [r.item_code, r])).values())
          setPoItems(unique)
        }
      })
    supabase.from('invoices').select('id, invoice_no, supplier, rows').order('invoice_no').then(({ data }) => {
      if (data) {
        setInvoicesData(data as typeof invoicesData)
        const uniq = [...new Set(data.map(r => r.supplier).filter(Boolean) as string[])].sort()
        setSuppliers(uniq)
      }
    })
    supabase.from('qc_reports').select('issue_types').then(({ data }) => {
      if (data) {
        const all = data.flatMap(r => r.issue_types ?? [])
        setIssueTypeOptions([...new Set(all)].sort())
      }
    })
  }, [id])

  function getItemsForDropdown(invoice_no: string): { item_code: string; description: string; fob_price: number | null; currency: string | null; inInvoice: boolean }[] {
    if (!invoice_no) return poItems.map(p => ({ ...p, inInvoice: false }))
    const inv = invoicesData.find(i => i.invoice_no === invoice_no)
    const codes = new Set((inv?.rows ?? []).map(r => r.code).filter(Boolean) as string[])
    const inInv = poItems.filter(p => codes.has(p.item_code)).map(p => ({ ...p, inInvoice: true }))
    const others = poItems.filter(p => !codes.has(p.item_code)).map(p => ({ ...p, inInvoice: false }))
    return [...inInv, ...others]
  }

  function set<K extends keyof typeof form>(k: K, v: typeof form[K]) {
    setForm(prev => {
      const next = { ...prev, [k]: v }
      if (k === 'invoice_no' && typeof v === 'string') {
        const inv = invoicesData.find(i => i.invoice_no === v)
        if (inv?.supplier && !prev.supplier_company) next.supplier_company = inv.supplier
      }
      return next
    })
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

  async function uploadPhotos(files: FileList) {
    setUploadingPhotos(true)
    const newUrls: string[] = []
    for (const file of Array.from(files)) {
      const path = `qc-photos/${id}/${Date.now()}-${file.name}`
      const { error } = await supabase.storage.from('payment-proofs').upload(path, file, { upsert: true })
      if (!error) {
        const { data: urlData } = supabase.storage.from('payment-proofs').getPublicUrl(path)
        newUrls.push(urlData.publicUrl)
      }
    }
    if (newUrls.length > 0) {
      const updated = [...(form.photo_urls || []), ...newUrls]
      set('photo_urls', updated)
    }
    setUploadingPhotos(false)
  }

  function removePhoto(url: string) {
    set('photo_urls', (form.photo_urls || []).filter(u => u !== url))
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
          <Link key={n.href} href={n.href} className="text-gray-500 hover:text-gray-800 transition-colors">{n.label}</Link>
        ))}
        <div className="relative group">
          <span className="text-blue-600 font-semibold cursor-default">QC Report ▾</span>
          <div className="absolute left-0 top-full pt-1 hidden group-hover:block z-50">
            <div className="bg-white border border-gray-200 rounded-lg shadow-lg py-1 min-w-[140px]">
              <Link href="/qc" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">รายการ QC</Link>
              <Link href="/qc/summary" className="block px-4 py-2 text-sm text-gray-700 hover:bg-blue-50">QC Summary</Link>
            </div>
          </div>
        </div>
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
            <button onClick={() => requireUnlock(save)} disabled={saving}
              className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold px-4 py-2 rounded-lg transition-colors disabled:opacity-50">
              {saving ? 'กำลังบันทึก...' : saved ? '✓ บันทึกแล้ว' : 'บันทึก'}
            </button>
            <button onClick={() => requireUnlock(deleteReport)} className="text-xs text-red-400 hover:text-red-600 px-2 py-2">ลบ</button>
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
            <div className="relative">
              <label className={labelCls}>Supplier Company</label>
              <input className={inputCls} value={form.supplier_company}
                onChange={e => { set('supplier_company', e.target.value); setSupplierSearch(e.target.value.toLowerCase()); setOpenSupplierDrop(true) }}
                onFocus={() => { setSupplierSearch(''); setOpenSupplierDrop(true) }}
                onBlur={() => setTimeout(() => setOpenSupplierDrop(false), 150)}
                placeholder="SHANGHAI YONGGUAN..." />
              {openSupplierDrop && (
                <ul className="absolute z-50 left-0 top-full max-h-48 overflow-y-auto bg-white border border-gray-300 shadow-lg rounded text-xs w-full">
                  {suppliers.filter(s => !supplierSearch || s.toLowerCase().includes(supplierSearch)).slice(0, 30).map(s => (
                    <li key={s} className="px-3 py-2 hover:bg-blue-50 cursor-pointer border-b border-gray-100 last:border-0 truncate"
                      onMouseDown={() => { set('supplier_company', s); setOpenSupplierDrop(false) }}>
                      {s}
                    </li>
                  ))}
                </ul>
              )}
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

          {/* Issue Types */}
          <div className="mt-4">
            <label className={labelCls}>Issue Type</label>
            <div className="flex flex-wrap gap-2 mb-2">
              {(form.issue_types ?? []).map(t => (
                <span key={t} className="inline-flex items-center gap-1 px-2.5 py-1 bg-orange-100 text-orange-700 rounded-full text-xs font-semibold">
                  {t}
                  <button type="button" onClick={() => set('issue_types', form.issue_types.filter(x => x !== t))}
                    className="ml-0.5 text-orange-400 hover:text-orange-700 leading-none">×</button>
                </span>
              ))}
            </div>
            <div className="relative w-72">
              <input
                className={inputCls}
                value={issueTypeInput}
                onChange={e => { setIssueTypeInput(e.target.value); setShowIssueTypeDrop(true) }}
                onFocus={() => setShowIssueTypeDrop(true)}
                onBlur={() => setTimeout(() => setShowIssueTypeDrop(false), 150)}
                onKeyDown={e => {
                  if (e.key === 'Enter' && issueTypeInput.trim()) {
                    e.preventDefault()
                    const t = issueTypeInput.trim()
                    if (!form.issue_types.includes(t)) set('issue_types', [...form.issue_types, t])
                    if (!issueTypeOptions.includes(t)) setIssueTypeOptions(prev => [...prev, t].sort())
                    setIssueTypeInput('')
                    setShowIssueTypeDrop(false)
                  }
                }}
                placeholder="พิมพ์แล้วกด Enter หรือเลือกจาก dropdown…"
              />
              {showIssueTypeDrop && (
                <ul className="absolute z-50 left-0 top-full mt-1 max-h-48 overflow-y-auto bg-white border border-gray-300 shadow-lg rounded text-sm w-full">
                  {[
                    ...['Packing Quality', 'Working Process', 'Product Quality', 'Loading Process'],
                    ...issueTypeOptions,
                  ]
                    .filter((t, i, arr) => arr.indexOf(t) === i)
                    .filter(t => !issueTypeInput || t.toLowerCase().includes(issueTypeInput.toLowerCase()))
                    .filter(t => !(form.issue_types ?? []).includes(t))
                    .map(t => (
                      <li key={t}
                        className="px-3 py-2 hover:bg-blue-50 cursor-pointer border-b border-gray-100 last:border-0"
                        onMouseDown={() => {
                          set('issue_types', [...(form.issue_types ?? []), t])
                          if (!issueTypeOptions.includes(t)) setIssueTypeOptions(prev => [...prev, t].sort())
                          setIssueTypeInput('')
                          setShowIssueTypeDrop(false)
                        }}>
                        {t}
                      </li>
                    ))
                  }
                  {issueTypeInput.trim() && !['Packing Quality', 'Working Process', 'Product Quality', 'Loading Process', ...issueTypeOptions].includes(issueTypeInput.trim()) && (
                    <li className="px-3 py-2 text-blue-600 hover:bg-blue-50 cursor-pointer font-semibold"
                      onMouseDown={() => {
                        const t = issueTypeInput.trim()
                        set('issue_types', [...(form.issue_types ?? []), t])
                        setIssueTypeOptions(prev => [...prev, t].sort())
                        setIssueTypeInput('')
                        setShowIssueTypeDrop(false)
                      }}>
                      + เพิ่ม "{issueTypeInput.trim()}"
                    </li>
                  )}
                </ul>
              )}
            </div>
          </div>

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
                          const base = getItemsForDropdown(form.invoice_no)
                          setItemSearch(base.filter(p => p.item_code.toLowerCase().includes(q) || p.description?.toLowerCase().includes(q)).map(p => p.item_code))
                          setOpenDropdown(i)
                        }}
                        onFocus={() => {
                          const base = getItemsForDropdown(form.invoice_no)
                          const q = item.item_code.toLowerCase()
                          const filtered = q
                            ? base.filter(p => p.item_code.toLowerCase().includes(q) || p.description?.toLowerCase().includes(q))
                            : base
                          setItemSearch(filtered.map(p => p.item_code))
                          setOpenDropdown(i)
                        }}
                        onBlur={() => setTimeout(() => setOpenDropdown(null), 150)}
                        placeholder="item code…"
                      />
                      {openDropdown === i && itemSearch.length > 0 && (() => {
                        const allItems = getItemsForDropdown(form.invoice_no)
                        const displayed = itemSearch.slice(0, 60).map(code => allItems.find(x => x.item_code === code)!)
                        const hasInvoiceItems = displayed.some(p => p?.inInvoice)
                        const hasOtherItems = displayed.some(p => p && !p.inInvoice)
                        return (
                          <ul className="absolute z-50 left-0 top-full max-h-56 overflow-y-auto bg-white border border-gray-300 shadow-lg rounded text-xs w-56">
                            {displayed.map((p, idx) => {
                              if (!p) return null
                              const prevP = idx > 0 ? displayed[idx - 1] : null
                              const showSep = hasInvoiceItems && hasOtherItems && !p.inInvoice && (idx === 0 || prevP?.inInvoice)
                              return (
                                <li key={p.item_code}>
                                  {showSep && <div className="px-2 py-1 text-gray-400 bg-gray-50 border-b border-gray-200 font-medium" style={{ fontSize: 10 }}>— Items อื่นๆ —</div>}
                                  <div
                                    className="px-2 py-1.5 hover:bg-blue-50 cursor-pointer border-b border-gray-100 last:border-0"
                                    onMouseDown={() => {
                                      const items = [...form.items]
                                      items[i] = {
                                        ...items[i],
                                        item_code: p.item_code,
                                        product_description: p.description || items[i].product_description,
                                        unit_price: p.fob_price ?? items[i].unit_price,
                                        total: Math.round((items[i].qty) * (p.fob_price ?? items[i].unit_price) * 100) / 100,
                                      }
                                      set('items', items)
                                      setOpenDropdown(null)
                                    }}>
                                    <div className={`font-mono font-semibold ${p.inInvoice ? 'text-blue-700' : 'text-gray-800'}`}>{p.item_code}</div>
                                    {p.description && <div className="text-gray-400 truncate max-w-[200px]">{p.description}</div>}
                                  </div>
                                </li>
                              )
                            })}
                          </ul>
                        )
                      })()}
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

        {/* Part 5: PHOTO */}
        <SectionHeader title="Part 5 : PHOTO" />
        <div className="bg-white border border-gray-300 border-t-0 rounded-b-lg p-5 mb-1">
          {(form.photo_urls || []).length > 0 && (
            <div className="grid grid-cols-3 gap-3 mb-4">
              {(form.photo_urls || []).map((url, idx) => (
                <div key={idx} className="relative group rounded-lg overflow-hidden border border-gray-200">
                  <img src={url} alt={`photo-${idx + 1}`} className="w-full h-48 object-cover" />
                  <button
                    onClick={() => removePhoto(url)}
                    className="absolute top-1.5 right-1.5 bg-red-500 hover:bg-red-600 text-white rounded-full w-6 h-6 text-xs font-bold opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                    ✕
                  </button>
                  <div className="absolute bottom-1.5 left-1.5 bg-black/50 text-white text-xs px-1.5 py-0.5 rounded">{idx + 1}</div>
                </div>
              ))}
            </div>
          )}
          <button
            onClick={() => photoRef.current?.click()}
            disabled={uploadingPhotos}
            className="flex items-center gap-2 border-2 border-dashed border-gray-200 rounded-lg px-5 py-3 text-sm text-gray-400 hover:border-blue-300 hover:text-blue-500 transition-colors disabled:opacity-50">
            {uploadingPhotos ? 'กำลังอัปโหลด...' : '+ เพิ่มรูปภาพ (เลือกได้หลายรูปพร้อมกัน)'}
          </button>
          <input ref={photoRef} type="file" accept="image/*" multiple className="hidden"
            onChange={e => { if (e.target.files?.length) uploadPhotos(e.target.files); e.target.value = '' }} />
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
          <button onClick={() => requireUnlock(save)} disabled={saving}
            className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold px-6 py-2.5 rounded-lg transition-colors disabled:opacity-50">
            {saving ? 'กำลังบันทึก...' : saved ? '✓ บันทึกแล้ว' : 'บันทึก'}
          </button>
        </div>
      </div>
      {showPW && pwCallback && (
        <PasswordGate onSuccess={() => { setShowPW(false); pwCallback() }} onCancel={() => setShowPW(false)} />
      )}
    </div>
  )
}

'use client'

import { useEffect, useRef, useState } from 'react'
import { useParams, useRouter } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import { isUnlocked } from '@/lib/auth'
import NavBar from '@/components/NavBar'
import PasswordModal from '@/components/PasswordModal'

interface PORow {
  item_code: string
  description: string
  qty: number
  unit_price: number
  total: number
}

interface PODetail {
  id: string
  supplier: string
  project: string
  currency: string
  filename: string | null
  po_date: string | null
  po_rbs_ch_no: string | null
  po_rbs_th_no: string | null
  rows: PORow[]
  total_amount: number
  cost_saving: number | null
  cost_saving_pct: number | null
  cost_saving_file_url: string | null
  created_at: string
  updated_at: string
}

export default function PODetailPage() {
  const { id } = useParams<{ id: string }>()
  const router = useRouter()
  const costFileRef = useRef<HTMLInputElement>(null)

  const [po, setPO] = useState<PODetail | null>(null)
  const [loading, setLoading] = useState(true)
  const [unlocked, setUnlocked] = useState(false)
  const [showPasswordModal, setShowPasswordModal] = useState(false)
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null)

  // Date editing
  const [editingDate, setEditingDate] = useState(false)
  const [newDate, setNewDate] = useState('')
  const [savingDate, setSavingDate] = useState(false)

  // PO name editing
  const [editingNames, setEditingNames] = useState(false)
  const [newChNo, setNewChNo] = useState('')
  const [newThNo, setNewThNo] = useState('')
  const [savingNames, setSavingNames] = useState(false)

  // Cost saving
  const [editCost, setEditCost] = useState(false)
  const [costInput, setCostInput] = useState('')
  const [costPctInput, setCostPctInput] = useState('')
  const [savingCost, setSavingCost] = useState(false)
  const [uploadingCostFile, setUploadingCostFile] = useState(false)

  useEffect(() => {
    setUnlocked(isUnlocked())
    load()
  }, [id])

  async function load() {
    setLoading(true)
    const { data, error } = await supabase
      .from('po_uploads')
      .select('*')
      .eq('id', id)
      .single()
    if (error || !data) { router.replace('/po-builder'); return }
    const d = data as PODetail
    setPO(d)
    setNewDate(d.po_date ?? '')
    setNewChNo(d.po_rbs_ch_no ?? '')
    setNewThNo(d.po_rbs_th_no ?? '')
    setLoading(false)
  }

  function requireUnlock(action: () => void) {
    if (unlocked || isUnlocked()) {
      action()
    } else {
      setPendingAction(() => action)
      setShowPasswordModal(true)
    }
  }

  function onUnlockSuccess() {
    setUnlocked(true)
    setShowPasswordModal(false)
    if (pendingAction) { pendingAction(); setPendingAction(null) }
  }

  async function saveDate() {
    if (!po) return
    setSavingDate(true)
    await supabase
      .from('po_uploads')
      .update({ po_date: newDate || null, updated_at: new Date().toISOString() })
      .eq('id', po.id)
    setPO(p => p ? { ...p, po_date: newDate || null } : p)
    setEditingDate(false)
    setSavingDate(false)
  }

  async function saveNames() {
    if (!po) return
    setSavingNames(true)
    await supabase
      .from('po_uploads')
      .update({ po_rbs_ch_no: newChNo.trim() || null, po_rbs_th_no: newThNo.trim() || null, updated_at: new Date().toISOString() })
      .eq('id', po.id)
    setPO(p => p ? { ...p, po_rbs_ch_no: newChNo.trim() || null, po_rbs_th_no: newThNo.trim() || null } : p)
    setEditingNames(false)
    setSavingNames(false)
  }

  async function saveCostSaving() {
    if (!po) return
    const cost = costInput === '' ? null : parseFloat(costInput.replace(/,/g, ''))
    const pct = costPctInput === '' ? null : parseFloat(costPctInput.replace(/,/g, ''))
    setSavingCost(true)
    await supabase.from('po_uploads').update({ cost_saving: cost, cost_saving_pct: pct, updated_at: new Date().toISOString() }).eq('id', po.id)
    setPO(p => p ? { ...p, cost_saving: cost, cost_saving_pct: pct } : p)
    setEditCost(false)
    setSavingCost(false)
  }

  async function uploadCostFile(file: File) {
    if (!po) return
    setUploadingCostFile(true)
    try {
      const ext = file.name.split('.').pop()
      const path = `po-${po.id}/cost-saving-${Date.now()}.${ext}`
      const { error } = await supabase.storage.from('payment-proofs').upload(path, file, { upsert: true })
      if (error) { alert(`อัปโหลดไม่สำเร็จ: ${error.message}`); return }
      const { data: urlData } = supabase.storage.from('payment-proofs').getPublicUrl(path)
      await supabase.from('po_uploads').update({ cost_saving_file_url: urlData.publicUrl, updated_at: new Date().toISOString() }).eq('id', po.id)
      setPO(p => p ? { ...p, cost_saving_file_url: urlData.publicUrl } : p)
    } finally {
      setUploadingCostFile(false)
    }
  }

  async function deletePO() {
    if (!po) return
    if (!confirm(`ลบ PO ของ ${po.supplier} (${po.project}) ออกจากระบบ?`)) return
    await supabase.from('po_uploads').delete().eq('id', po.id)
    router.replace('/po-builder')
  }

  const fmt = (n: number, d = 2) =>
    n.toLocaleString(undefined, { minimumFractionDigits: d, maximumFractionDigits: d })

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50 flex flex-col">
        <NavBar />
        <div className="flex-1 flex items-center justify-center">
          <p className="text-sm text-gray-400">กำลังโหลด...</p>
        </div>
      </div>
    )
  }

  if (!po) return null

  const totalQty = po.rows.reduce((s, r) => s + r.qty, 0)

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      <NavBar
        onUnlock={() => setUnlocked(true)}
        onLock={() => setUnlocked(false)}
      />

      {showPasswordModal && (
        <PasswordModal
          onSuccess={onUnlockSuccess}
          onCancel={() => { setShowPasswordModal(false); setPendingAction(null) }}
        />
      )}

      <input ref={costFileRef} type="file" accept=".xlsx,.xls,.csv,.pdf" className="hidden"
        onChange={e => { const f = e.target.files?.[0]; if (f) uploadCostFile(f); e.target.value = '' }} />

      <div className="max-w-5xl mx-auto w-full px-6 py-8">
        {/* Back */}
        <button onClick={() => router.back()} className="text-sm text-gray-400 hover:text-gray-600 mb-4 flex items-center gap-1">
          ← กลับ
        </button>

        <div className="flex gap-6 flex-col lg:flex-row">
          {/* Left: main info + table */}
          <div className="flex-1 min-w-0">
            {/* Header card */}
            <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mb-5">
              <div className="flex items-start justify-between flex-wrap gap-4 mb-4">
                <div>
                  <div className="flex items-center gap-3 flex-wrap mb-1">
                    <h1 className="text-xl font-bold text-gray-900">{po.supplier}</h1>
                    <span className="text-xs bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full font-medium">{po.project}</span>
                    <span className="text-xs bg-gray-100 text-gray-600 px-2 py-0.5 rounded-full">{po.currency}</span>
                  </div>
                  {po.filename && <p className="text-xs text-gray-400">{po.filename}</p>}
                </div>
                <button
                  onClick={() => requireUnlock(deletePO)}
                  className="text-xs text-red-400 hover:text-red-600 border border-red-200 px-3 py-1 rounded-lg hover:bg-red-50 transition-colors"
                >
                  ลบ PO นี้
                </button>
              </div>

              {/* PO Names */}
              <div className="border border-gray-100 rounded-lg p-3 mb-4 bg-gray-50">
                <div className="flex items-center justify-between mb-2">
                  <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide">ชื่อ PO</p>
                  {!editingNames ? (
                    <button
                      onClick={() => requireUnlock(() => setEditingNames(true))}
                      className="text-xs text-blue-500 hover:text-blue-700"
                    >
                      {po.po_rbs_ch_no || po.po_rbs_th_no ? 'แก้ไข' : '+ เพิ่มชื่อ PO'}
                    </button>
                  ) : null}
                </div>
                {editingNames ? (
                  <div className="space-y-2">
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="text-xs text-gray-500 mb-1 block">PO RBS CH No.</label>
                        <input
                          type="text"
                          value={newChNo}
                          onChange={e => setNewChNo(e.target.value)}
                          placeholder="เช่น RBSYG01-GEN5"
                          className="w-full border border-gray-300 rounded-lg px-2 py-1.5 text-sm outline-none focus:border-blue-400"
                          autoFocus
                        />
                      </div>
                      <div>
                        <label className="text-xs text-gray-500 mb-1 block">PO RBS TH No.</label>
                        <input
                          type="text"
                          value={newThNo}
                          onChange={e => setNewThNo(e.target.value)}
                          placeholder="เช่น RBSTH-YG01"
                          className="w-full border border-gray-300 rounded-lg px-2 py-1.5 text-sm outline-none focus:border-blue-400"
                        />
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <button onClick={() => { setEditingNames(false); setNewChNo(po.po_rbs_ch_no ?? ''); setNewThNo(po.po_rbs_th_no ?? '') }}
                        className="text-sm px-3 py-1 text-gray-500 hover:bg-gray-100 rounded-lg">ยกเลิก</button>
                      <button onClick={saveNames} disabled={savingNames}
                        className="text-sm px-3 py-1 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50">
                        {savingNames ? 'บันทึก...' : 'บันทึก'}
                      </button>
                    </div>
                  </div>
                ) : (
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <p className="text-xs text-gray-400">PO RBS CH No.</p>
                      <p className="text-sm font-mono font-medium text-gray-800">{po.po_rbs_ch_no || <span className="text-gray-400 font-normal">—</span>}</p>
                    </div>
                    <div>
                      <p className="text-xs text-gray-400">PO RBS TH No.</p>
                      <p className="text-sm font-mono font-medium text-gray-800">{po.po_rbs_th_no || <span className="text-gray-400 font-normal">—</span>}</p>
                    </div>
                  </div>
                )}
              </div>

              {/* Stats */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 pt-4 border-t border-gray-100">
                <div>
                  <p className="text-xs text-gray-400 mb-0.5">วันที่เปิด PO</p>
                  {editingDate ? (
                    <div className="flex items-center gap-2 flex-wrap">
                      <input
                        type="date"
                        value={newDate}
                        onChange={e => setNewDate(e.target.value)}
                        className="border border-blue-300 rounded px-2 py-1 text-sm outline-none focus:border-blue-500"
                        autoFocus
                      />
                      <button onClick={saveDate} disabled={savingDate} className="text-xs text-blue-600 hover:text-blue-800 font-medium">
                        {savingDate ? '...' : 'บันทึก'}
                      </button>
                      <button onClick={() => { setEditingDate(false); setNewDate(po.po_date ?? '') }} className="text-xs text-gray-400 hover:text-gray-600">ยกเลิก</button>
                    </div>
                  ) : (
                    <div className="flex items-center gap-2">
                      <p className="text-sm font-medium text-gray-800">
                        {po.po_date
                          ? new Date(po.po_date).toLocaleDateString('th-TH', { day: '2-digit', month: 'short', year: 'numeric' })
                          : <span className="text-gray-400">—</span>}
                      </p>
                      <button onClick={() => requireUnlock(() => setEditingDate(true))} className="text-xs text-gray-400 hover:text-blue-500">แก้ไข</button>
                    </div>
                  )}
                </div>
                <div>
                  <p className="text-xs text-gray-400 mb-0.5">จำนวนรายการ</p>
                  <p className="text-sm font-medium text-gray-800">{po.rows.length} รายการ</p>
                </div>
                <div>
                  <p className="text-xs text-gray-400 mb-0.5">QTY รวม</p>
                  <p className="text-sm font-medium text-gray-800">{fmt(totalQty, 0)} ชิ้น</p>
                </div>
                <div>
                  <p className="text-xs text-gray-400 mb-0.5">ยอดรวม</p>
                  <p className="text-sm font-bold text-gray-900">{fmt(po.total_amount)} {po.currency}</p>
                </div>
              </div>
            </div>

            {/* Items table */}
            <div className="overflow-auto rounded-xl border border-gray-200 bg-white shadow-sm">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="bg-gray-50 text-gray-500 text-xs border-b border-gray-200">
                    <th className="px-4 py-3 text-center w-10">No.</th>
                    <th className="px-4 py-3 text-left min-w-[160px]">Item Code</th>
                    <th className="px-4 py-3 text-left">Description</th>
                    <th className="px-4 py-3 text-right w-24">QTY</th>
                    <th className="px-4 py-3 text-right w-36">Unit Price ({po.currency})</th>
                    <th className="px-4 py-3 text-right w-36">Total ({po.currency})</th>
                  </tr>
                </thead>
                <tbody>
                  {po.rows.map((row, i) => (
                    <tr key={i} className="border-b border-gray-100 hover:bg-gray-50/50">
                      <td className="px-4 py-2.5 text-center text-gray-400 text-xs">{i + 1}</td>
                      <td className="px-4 py-2.5 font-mono text-gray-900 text-xs whitespace-nowrap">{row.item_code}</td>
                      <td className="px-4 py-2.5 text-gray-600 text-xs">{row.description || <span className="text-gray-400">—</span>}</td>
                      <td className="px-4 py-2.5 text-right text-gray-700">{fmt(row.qty, 0)}</td>
                      <td className="px-4 py-2.5 text-right text-gray-700">{fmt(row.unit_price)}</td>
                      <td className="px-4 py-2.5 text-right font-medium text-gray-900">{fmt(row.total)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-gray-50 text-xs font-semibold border-t-2 border-gray-200 text-gray-700">
                    <td colSpan={3} className="px-4 py-3 text-right">TOTAL</td>
                    <td className="px-4 py-3 text-right">{fmt(totalQty, 0)}</td>
                    <td></td>
                    <td className="px-4 py-3 text-right">{fmt(po.total_amount)} {po.currency}</td>
                  </tr>
                </tfoot>
              </table>
            </div>

            <p className="text-xs text-gray-400 mt-3 text-right">
              อัปโหลดเมื่อ {new Date(po.created_at).toLocaleDateString('th-TH', { day: '2-digit', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' })}
            </p>
          </div>

          {/* Right sidebar: Cost Saving */}
          <div className="w-full lg:w-72 shrink-0">
            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <div className="flex items-center justify-between mb-2">
                <p className="text-xs text-gray-400">Cost Saving</p>
                {!editCost && (
                  <button
                    onClick={() => requireUnlock(() => {
                      setCostInput(po.cost_saving != null ? String(po.cost_saving) : '')
                      setCostPctInput(po.cost_saving_pct != null ? String(po.cost_saving_pct) : '')
                      setEditCost(true)
                    })}
                    className="text-xs text-blue-500 hover:text-blue-700"
                  >
                    {po.cost_saving != null ? 'แก้ไข' : '+ เพิ่มข้อมูล'}
                  </button>
                )}
              </div>

              {editCost ? (
                <div className="flex flex-col gap-2">
                  <input
                    type="number"
                    value={costInput}
                    onChange={e => setCostInput(e.target.value)}
                    placeholder="Cost Saving (THB)"
                    className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 w-full"
                    autoFocus
                  />
                  <input
                    type="number"
                    value={costPctInput}
                    onChange={e => setCostPctInput(e.target.value)}
                    placeholder="% Cost Saving"
                    className="border border-gray-300 rounded-lg px-3 py-1.5 text-sm outline-none focus:border-blue-400 w-full"
                  />
                  <div className="flex gap-2 mt-1">
                    <button onClick={() => setEditCost(false)}
                      className="text-sm px-3 py-1 text-gray-500 hover:bg-gray-100 rounded-lg">ยกเลิก</button>
                    <button onClick={saveCostSaving} disabled={savingCost}
                      className="text-sm px-3 py-1 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50">
                      {savingCost ? 'บันทึก...' : 'บันทึก'}
                    </button>
                  </div>
                </div>
              ) : (
                <div className="mb-3 space-y-1">
                  <p className="text-sm font-semibold text-gray-800">
                    {po.cost_saving != null
                      ? `${po.cost_saving.toLocaleString()} THB`
                      : <span className="text-gray-400 font-normal">—</span>}
                  </p>
                  <p className="text-sm font-semibold text-green-600">
                    {po.cost_saving_pct != null
                      ? `${po.cost_saving_pct}%`
                      : <span className="text-gray-400 font-normal">—</span>}
                  </p>
                </div>
              )}

              <div className="pt-2 border-t border-gray-100 mt-2">
                {po.cost_saving_file_url ? (
                  <div className="flex flex-col gap-1">
                    <a href={po.cost_saving_file_url} target="_blank" rel="noopener noreferrer"
                      className="inline-flex items-center gap-1.5 text-sm text-blue-600 hover:text-blue-800 font-medium">
                      <span>📊</span> ดาวน์โหลด Cost Saving
                    </a>
                    <button
                      onClick={() => requireUnlock(() => costFileRef.current?.click())}
                      className="text-xs text-gray-400 hover:text-gray-600 self-start"
                    >
                      {uploadingCostFile ? 'กำลังอัปโหลด...' : 'เปลี่ยนไฟล์'}
                    </button>
                  </div>
                ) : (
                  <button
                    onClick={() => requireUnlock(() => costFileRef.current?.click())}
                    disabled={uploadingCostFile}
                    className="flex items-center gap-2 border-2 border-dashed border-gray-200 rounded-lg px-3 py-2 text-xs text-gray-400 hover:border-blue-300 hover:text-blue-500 w-full justify-center transition-colors disabled:opacity-50"
                  >
                    {uploadingCostFile ? 'กำลังอัปโหลด...' : '+ อัปโหลด Excel Cost Saving'}
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}

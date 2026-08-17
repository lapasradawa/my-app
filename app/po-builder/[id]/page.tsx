'use client'

import { useEffect, useState } from 'react'
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
  rows: PORow[]
  total_amount: number
  created_at: string
  updated_at: string
}

export default function PODetailPage() {
  const { id } = useParams<{ id: string }>()
  const router = useRouter()

  const [po, setPO] = useState<PODetail | null>(null)
  const [loading, setLoading] = useState(true)
  const [editingDate, setEditingDate] = useState(false)
  const [newDate, setNewDate] = useState('')
  const [savingDate, setSavingDate] = useState(false)
  const [showPasswordModal, setShowPasswordModal] = useState(false)
  const [unlocked, setUnlocked] = useState(false)
  const [, setUnlockedState] = useState(false)

  useEffect(() => {
    setUnlocked(isUnlocked())
    setUnlockedState(isUnlocked())
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
    setPO(data as PODetail)
    setNewDate((data as PODetail).po_date ?? '')
    setLoading(false)
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

  function requireUnlock(action: () => void) {
    if (unlocked || isUnlocked()) { action() }
    else { setShowPasswordModal(true) }
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
      <NavBar onUnlock={() => { setUnlocked(true); setUnlockedState(true) }} onLock={() => { setUnlocked(false); setUnlockedState(false) }} />

      {showPasswordModal && (
        <PasswordModal
          onSuccess={() => { setUnlocked(true); setUnlockedState(true); setShowPasswordModal(false); setEditingDate(true) }}
          onCancel={() => setShowPasswordModal(false)}
        />
      )}

      <div className="max-w-5xl mx-auto w-full px-6 py-8">
        {/* Back */}
        <button onClick={() => router.back()} className="text-sm text-gray-400 hover:text-gray-600 mb-4 flex items-center gap-1">
          ← กลับ
        </button>

        {/* Header card */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mb-6">
          <div className="flex items-start justify-between flex-wrap gap-4">
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

          {/* Stats */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mt-5 pt-5 border-t border-gray-100">
            <div>
              <p className="text-xs text-gray-400 mb-0.5">วันที่เปิด PO</p>
              {editingDate ? (
                <div className="flex items-center gap-2">
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
                  <button
                    onClick={() => requireUnlock(() => setEditingDate(true))}
                    className="text-xs text-gray-400 hover:text-blue-500"
                  >
                    แก้ไข
                  </button>
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
    </div>
  )
}

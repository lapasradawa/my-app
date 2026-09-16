'use client'

import { useCallback, useEffect, useMemo, useState } from 'react'
import Link from 'next/link'
import { useRouter } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import { isUnlocked, tryUnlock } from '@/lib/auth'
import NavBar from '@/components/NavBar'

// ── Period helpers ────────────────────────────────────────────────────────────
function mKey(d: string | null): string | null {
  if (!d) return null
  return d.slice(0, 7)
}
function mLabel(k: string): string {
  const [y, m] = k.split('-')
  const names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
  return `${names[parseInt(m) - 1]} ${y}`
}
function generateMonthKeys(count = 18): string[] {
  const keys: string[] = []
  const d = new Date()
  for (let i = 0; i < count; i++) {
    keys.unshift(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`)
    d.setMonth(d.getMonth() - 1)
  }
  return keys
}

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

interface QCReport {
  id: string
  report_no: string
  supplier_company: string | null
  invoice_no: string | null
  issue_found_date: string | null
  status: string
  verification_accepted: boolean | null
  created_at: string
}


export default function QCPage() {
  const router = useRouter()
  const [reports, setReports] = useState<QCReport[]>([])
  const [loading, setLoading] = useState(true)
  const [creating, setCreating] = useState(false)
  const [showPW, setShowPW] = useState(false)

  // ── Period filter ─────────────────────────────────────────────────────
  const allMonthKeys = useMemo(() => generateMonthKeys(18), [])
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(() => {
    const now = new Date()
    return new Set([`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`])
  })
  const [periodOpen, setPeriodOpen] = useState(false)
  const toggleMonth = useCallback((k: string) => {
    setSelectedMonths(prev => { const n = new Set(prev); n.has(k) ? n.delete(k) : n.add(k); return n })
  }, [])

  const filteredReports = useMemo(() => {
    if (selectedMonths.size === 0) return reports
    return reports.filter(r => {
      const k = mKey(r.created_at)
      return k !== null && selectedMonths.has(k)
    })
  }, [reports, selectedMonths])

  function requireUnlock(action: () => void) {
    if (isUnlocked()) { action() }
    else { setShowPW(true); setPendingAction(() => action) }
  }
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null)

  useEffect(() => { loadReports() }, [])

  async function loadReports() {
    setLoading(true)
    const { data } = await supabase
      .from('qc_reports')
      .select('id, report_no, supplier_company, invoice_no, issue_found_date, status, verification_accepted, created_at')
      .order('created_at', { ascending: false })
    if (data) setReports(data as QCReport[])
    setLoading(false)
  }

  async function createNew() {
    setCreating(true)
    const now = new Date()
    const yy = now.getFullYear()
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const { data: existing } = await supabase
      .from('qc_reports')
      .select('report_no')
      .like('report_no', `RBSQC${yy}/${mm}-%`)
      .order('report_no', { ascending: false })
      .limit(1)
    const seq = existing?.[0]
      ? parseInt(existing[0].report_no.split('-').pop() || '0') + 1
      : 1
    const report_no = `RBSQC${yy}/${mm}-${String(seq).padStart(3, '0')}`
    const { data, error } = await supabase
      .from('qc_reports')
      .insert({ report_no, status: 'open', items: [], corrective_actions: [] })
      .select('id').single()
    if (!error && data) router.push(`/qc/${data.id}`)
    setCreating(false)
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <NavBar />

      <div className="max-w-5xl mx-auto px-6 py-8">
        <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">QC Report</h1>
            <p className="text-sm text-gray-500 mt-1">รายงานคุณภาพและการเคลม Supplier</p>
          </div>
          <button onClick={() => requireUnlock(createNew)} disabled={creating}
            className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold px-4 py-2 rounded-lg transition-colors disabled:opacity-50">
            {creating ? 'กำลังสร้าง...' : '+ สร้าง QC Report ใหม่'}
          </button>
        </div>

        {/* ── Period filter ──────────────────────────────────────────── */}
        <div className="bg-gray-800 rounded-xl px-5 py-3 mb-5 flex items-center gap-3 flex-wrap relative">
          <span className="text-xs font-bold text-gray-400 uppercase tracking-widest shrink-0">Period:</span>

          <div className="relative">
            <button onClick={() => setPeriodOpen(o => !o)}
              className="flex items-center gap-2 px-4 py-1.5 rounded-lg text-xs font-bold text-gray-200"
              style={{ background: '#1e3a4a', border: '1px solid #2e5060', minWidth: 170 }}>
              <span className="flex-1 text-left">
                {selectedMonths.size === 0
                  ? 'ทั้งหมด'
                  : selectedMonths.size === 1
                  ? mLabel([...selectedMonths][0])
                  : `${selectedMonths.size} เดือนที่เลือก`}
              </span>
              <span className="text-gray-500 text-xs">{periodOpen ? '▲' : '▼'}</span>
            </button>

            {periodOpen && (
              <div className="absolute top-full left-0 mt-1 z-50 rounded-xl p-3 shadow-2xl"
                style={{ background: '#1a2e3c', border: '1px solid #2e5060', minWidth: 280 }}>
                <div className="flex gap-2 mb-3 pb-2" style={{ borderBottom: '1px solid #2a4455' }}>
                  <button onClick={() => setSelectedMonths(new Set())}
                    className="flex-1 py-1 rounded-lg text-xs font-bold"
                    style={{ background: '#2a4455', color: '#8a9aaa' }}>ทั้งหมด</button>
                  <button onClick={() => setPeriodOpen(false)}
                    className="px-3 py-1 rounded-lg text-xs font-bold"
                    style={{ background: '#d4962a', color: '#fff' }}>Done</button>
                </div>
                <div className="grid grid-cols-3 gap-1.5">
                  {allMonthKeys.map(k => (
                    <button key={k} onClick={() => toggleMonth(k)}
                      className="py-1.5 rounded-lg text-xs font-bold transition-all"
                      style={selectedMonths.has(k)
                        ? { background: '#d4962a', color: '#1a2d3a', border: '1px solid #d4962a' }
                        : { background: 'transparent', color: '#8a9aaa', border: '1px solid #2a4455' }}>
                      {mLabel(k)}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>

          {[...selectedMonths].sort().map(k => (
            <span key={k} className="flex items-center gap-1 px-3 py-1 rounded-full text-xs font-bold"
              style={{ background: '#d4962a', color: '#fff' }}>
              {mLabel(k)}
              <button onClick={() => toggleMonth(k)} className="ml-1 opacity-75 hover:opacity-100 leading-none">×</button>
            </span>
          ))}

          {selectedMonths.size > 0 && (
            <span className="ml-auto text-xs text-gray-400">{filteredReports.length} QC Report</span>
          )}
        </div>

        {loading ? (
          <div className="text-center py-16 text-gray-400 text-sm">กำลังโหลด...</div>
        ) : filteredReports.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <div className="text-4xl mb-3">📋</div>
            <p className="text-sm">ไม่มี QC Report ในช่วงเวลาที่เลือก</p>
            <p className="text-xs mt-1">ลองเลือกเดือนอื่น หรือกด "สร้าง QC Report ใหม่" เพื่อเริ่มต้น</p>
          </div>
        ) : (
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-gray-50 border-b border-gray-200">
                  <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Report No.</th>
                  <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Supplier</th>
                  <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Invoice</th>
                  <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Issue Date</th>
                  <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">สถานะ</th>
                  <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Verification</th>
                </tr>
              </thead>
              <tbody>
                {filteredReports.map((r, i) => (
                  <tr key={r.id}
                    onClick={() => router.push(`/qc/${r.id}`)}
                    className={`border-b border-gray-100 hover:bg-blue-50 cursor-pointer transition-colors ${i % 2 === 0 ? '' : 'bg-gray-50/50'}`}>
                    <td className="px-4 py-3 font-mono font-semibold text-blue-700">{r.report_no}</td>
                    <td className="px-4 py-3 text-gray-700">{r.supplier_company || '—'}</td>
                    <td className="px-4 py-3 text-gray-600">{r.invoice_no || '—'}</td>
                    <td className="px-4 py-3 text-gray-600">{r.issue_found_date ? new Date(r.issue_found_date).toLocaleDateString('th-TH') : '—'}</td>
                    <td className="px-4 py-3">
                      <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold ${r.status === 'closed' ? 'bg-green-100 text-green-700' : 'bg-orange-100 text-orange-700'}`}>
                        {r.status === 'closed' ? '✓ ปิดจบแล้ว' : '● เปิดอยู่'}
                      </span>
                    </td>
                    <td className="px-4 py-3">
                      {r.verification_accepted === true && <span className="text-green-600 text-xs font-semibold">Accepted</span>}
                      {r.verification_accepted === false && <span className="text-red-500 text-xs font-semibold">Not Accepted</span>}
                      {r.verification_accepted === null && <span className="text-gray-400 text-xs">รอตรวจสอบ</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
      {showPW && pendingAction && (
        <PasswordGate onSuccess={() => { setShowPW(false); pendingAction() }} onCancel={() => setShowPW(false)} />
      )}
    </div>
  )
}

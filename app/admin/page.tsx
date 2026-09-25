'use client'

import { useEffect, useState } from 'react'
import { useRouter } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import { usePermissions, type PageKey } from '@/lib/permissions'
import NavBar from '@/components/NavBar'

const DEFAULT_EMAIL = '__default__'

interface PermRow {
  email: string
  is_admin: boolean
  allowed_pages: PageKey[] | null
}

const ALL_PAGES: { key: PageKey; label: string }[] = [
  { key: 'po-matching',  label: 'PO Matching' },
  { key: 'dashboard',    label: 'Dashboard' },
  { key: 'calendar',     label: 'Calendar' },
  { key: 'report',       label: 'Report' },
  { key: 'compare',      label: 'Cost Compare' },
  { key: 'po-builder',   label: 'PO Insights' },
  { key: 'order-plan',   label: 'Order Plan' },
  { key: 'summary',      label: 'Summary' },
  { key: 'qc',           label: 'QC Report' },
  { key: 'guide',        label: 'Guide' },
  { key: 'po-summary',  label: 'PO Summary' },
]

export default function AdminPage() {
  const { isAdmin } = usePermissions()
  const router = useRouter()

  const [rows, setRows] = useState<PermRow[]>([])
  const [defaultPages, setDefaultPages] = useState<PageKey[]>(['po-matching'])
  const [hubRequests, setHubRequests] = useState<{ id: string; invoice_id: string; invoice_no: string; container_name: string; hub: string; requested_by: string; created_at: string }[]>([])
  const [lineFailures, setLineFailures] = useState<{ id: string; context: string; invoice_id: string | null; invoice_no: string | null; container_name: string | null; status_code: number | null; response_body: string | null; error_message: string | null; created_at: string }[]>([])
  const [dismissingFailure, setDismissingFailure] = useState<string | null>(null)
  const [confirmedHubRequests, setConfirmedHubRequests] = useState<{ id: string; invoice_id: string; invoice_no: string; container_name: string; hub: string; confirmed_by: string | null; hub_arrival_date: string | null }[]>([])
  const [rejecting, setRejecting] = useState<string | null>(null) // container_name being rejected
  const [confirming, setConfirming] = useState<string | null>(null) // container_name being confirmed
  const [confirmDates, setConfirmDates] = useState<Record<string, string>>({})
  const [editDates, setEditDates] = useState<Record<string, string>>({}) // confirmed requests — editable ETA
  const [savingDate, setSavingDate] = useState<string | null>(null) // container_name being saved
  const [savingDefault, setSavingDefault] = useState(false)
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState<string | null>(null)
  const [newEmail, setNewEmail] = useState('')
  const [addError, setAddError] = useState('')

  useEffect(() => {
    if (!isAdmin) { router.replace('/'); return }
    load()
  }, [isAdmin])

  async function load() {
    setLoading(true)
    const { data } = await supabase
      .from('page_permissions')
      .select('email, is_admin, allowed_pages')
      .order('email')
    const all = (data ?? []) as PermRow[]
    const defRow = all.find(r => r.email === DEFAULT_EMAIL)
    if (defRow?.allowed_pages) setDefaultPages(defRow.allowed_pages)
    setRows(all.filter(r => r.email !== DEFAULT_EMAIL))
    const { data: hubs } = await supabase.from('container_hub_requests').select('id, invoice_id, invoice_no, container_name, hub, requested_by, created_at').eq('status', 'pending').order('created_at', { ascending: false })
    setHubRequests((hubs ?? []) as typeof hubRequests)
    const { data: confirmedHubs } = await supabase.from('container_hub_requests').select('id, invoice_id, invoice_no, container_name, hub, confirmed_by, hub_arrival_date').eq('status', 'confirmed').order('hub_arrival_date', { ascending: true })
    setConfirmedHubRequests((confirmedHubs ?? []) as typeof confirmedHubRequests)
    const { data: failures } = await supabase
      .from('line_notification_failures')
      .select('id, context, invoice_id, invoice_no, container_name, status_code, response_body, error_message, created_at')
      .eq('resolved', false)
      .order('created_at', { ascending: false })
    setLineFailures((failures ?? []) as typeof lineFailures)
    setLoading(false)
  }

  async function dismissLineFailure(id: string) {
    setDismissingFailure(id)
    await supabase.from('line_notification_failures').update({ resolved: true }).eq('id', id)
    setLineFailures(prev => prev.filter(f => f.id !== id))
    setDismissingFailure(null)
  }

  async function saveHubDate(r: typeof confirmedHubRequests[0]) {
    const key = `${r.invoice_id}:${r.container_name}`
    const newDate = editDates[key] ?? r.hub_arrival_date ?? ''
    if (!newDate) { alert('กรุณาใส่วันที่'); return }
    setSavingDate(r.container_name)
    const res = await fetch('/api/hub-request', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ invoice_id: r.invoice_id, container_name: r.container_name, action: 'update_date', hub_arrival_date: newDate }),
    })
    setSavingDate(null)
    if (!res.ok) {
      alert('บันทึกไม่สำเร็จ — ลองโหลดหน้าใหม่')
      return
    }
    setConfirmedHubRequests(prev => prev.map(x =>
      x.invoice_id === r.invoice_id && x.container_name === r.container_name ? { ...x, hub_arrival_date: newDate } : x
    ))
  }

  async function confirmHub(invoiceId: string, containerName: string) {
    const key = `${invoiceId}:${containerName}`
    const hubArrivalDate = confirmDates[key]
    if (!hubArrivalDate) { alert('กรุณาใส่วันที่เข้าคลังก่อนยืนยัน'); return }
    setConfirming(containerName)
    const { data: { user } } = await supabase.auth.getUser()
    const confirmedBy = user?.email ?? 'admin'
    const res = await fetch('/api/hub-request', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ invoice_id: invoiceId, container_name: containerName, confirmed_by: confirmedBy, hub_arrival_date: hubArrivalDate }),
    })
    setConfirming(null)
    if (res.status === 409) {
      alert('รายการนี้ถูกดำเนินการไปแล้วโดยแอดมินคนอื่น — รายการจะถูกโหลดใหม่')
      await load()
      return
    }
    setHubRequests(prev => prev.filter(r => !(r.invoice_id === invoiceId && r.container_name === containerName)))
  }

  async function rejectHub(r: typeof hubRequests[0]) {
    if (!confirm(`ปฏิเสธคำขอตู้ ${r.container_name} → Hub ${r.hub} ?\nตู้นี้จะกลับไปเป็น Warehouse มัยลาภตามเดิม`)) return
    setRejecting(r.container_name)
    const res = await fetch('/api/hub-request', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ invoice_id: r.invoice_id, container_name: r.container_name, action: 'reject', invoice_no: r.invoice_no, hub: r.hub, requested_by: r.requested_by }),
    })
    setRejecting(null)
    if (res.status === 409) {
      alert('รายการนี้ถูกดำเนินการไปแล้วโดยแอดมินคนอื่น — รายการจะถูกโหลดใหม่')
      await load()
      return
    }
    setHubRequests(prev => prev.filter(x => !(x.invoice_id === r.invoice_id && x.container_name === r.container_name)))
  }

  async function saveDefaultPages() {
    setSavingDefault(true)
    await supabase.from('page_permissions').upsert({
      email: DEFAULT_EMAIL,
      is_admin: false,
      allowed_pages: defaultPages,
      updated_at: new Date().toISOString(),
    })
    setSavingDefault(false)
  }

  function toggleDefaultPage(page: PageKey, checked: boolean) {
    setDefaultPages(prev =>
      checked ? [...new Set([...prev, page])] : prev.filter(p => p !== page)
    )
  }

  async function upsert(row: PermRow) {
    setSaving(row.email)
    await supabase.from('page_permissions').upsert({
      email: row.email,
      is_admin: row.is_admin,
      allowed_pages: row.allowed_pages,
      updated_at: new Date().toISOString(),
    })
    setSaving(null)
  }

  async function removeUser(email: string) {
    if (!confirm(`ลบ ${email} ออกจากรายการ? (จะเข้าได้เฉพาะหน้า Default เท่านั้น)`)) return
    await supabase.from('page_permissions').delete().eq('email', email)
    setRows(r => r.filter(x => x.email !== email))
  }

  async function addUser() {
    const email = newEmail.trim().toLowerCase()
    if (!email) return
    if (rows.find(r => r.email === email)) { setAddError('มีอยู่แล้ว'); return }
    setAddError('')
    const newRow: PermRow = { email, is_admin: false, allowed_pages: [] }
    const { error } = await supabase.from('page_permissions').insert({
      email,
      is_admin: false,
      allowed_pages: [],
    })
    if (error) { setAddError(error.message); return }
    setRows(r => [...r, newRow].sort((a, b) => a.email.localeCompare(b.email)))
    setNewEmail('')
  }

  function togglePage(email: string, page: PageKey, checked: boolean) {
    setRows(prev => prev.map(r => {
      if (r.email !== email) return r
      const current = r.allowed_pages ?? ALL_PAGES.map(p => p.key)
      const next: PageKey[] = checked
        ? [...new Set([...current, page])]
        : current.filter(p => p !== page)
      return { ...r, allowed_pages: next }
    }))
  }

  function setAllPages(email: string, all: boolean) {
    setRows(prev => prev.map(r =>
      r.email === email ? { ...r, allowed_pages: all ? null : [] } : r
    ))
  }

  function toggleAdmin(email: string, val: boolean) {
    setRows(prev => prev.map(r => r.email === email ? { ...r, is_admin: val } : r))
  }

  function isPageChecked(row: PermRow, page: PageKey): boolean {
    if (row.is_admin) return true
    if (row.allowed_pages === null) return true
    return row.allowed_pages.includes(page)
  }

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      <NavBar />

      <div className="max-w-5xl mx-auto w-full px-6 py-8">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900">Admin — จัดการสิทธิ์ผู้ใช้</h1>
        </div>

        {/* LINE notification failures — only rendered when something's actually
            wrong, so it adds zero clutter the rest of the time. */}
        {lineFailures.length > 0 && (
          <div className="bg-white border border-red-200 rounded-xl p-5 mb-8 shadow-sm">
            <h2 className="text-base font-bold text-gray-900 mb-1">🔔 แจ้งเตือน LINE ส่งไม่สำเร็จ ({lineFailures.length})</h2>
            <p className="text-xs text-gray-500 mb-3">อาจเกิดจากโควตาข้อความ LINE เต็มในเดือนนี้ หรือ token มีปัญหา — เช็คโควตาได้ที่ LINE Official Account Manager (manager.line.biz)</p>
            <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
              {lineFailures.map(f => (
                <div key={f.id} className="bg-red-50 rounded-lg border border-red-100 px-4 py-3 flex items-center justify-between gap-4 flex-wrap">
                  <div className="text-sm">
                    <span className="font-mono text-xs text-gray-500">{new Date(f.created_at).toLocaleString('th-TH')}</span>
                    <span className="text-gray-400 mx-2">·</span>
                    <span className="font-semibold text-gray-800">{f.context}</span>
                    {f.container_name && <><span className="text-gray-400 mx-2">·</span><span className="font-mono text-gray-700">{f.container_name}</span></>}
                    {f.invoice_no && <><span className="text-gray-400 mx-2">·</span><span className="text-gray-600">{f.invoice_no}</span></>}
                    <div className="text-xs text-red-600 mt-1">
                      {f.status_code ? `LINE ตอบกลับ ${f.status_code}: ${f.response_body?.slice(0, 200)}` : f.error_message}
                    </div>
                  </div>
                  <button
                    onClick={() => dismissLineFailure(f.id)}
                    disabled={dismissingFailure === f.id}
                    className="px-3 py-1 text-xs border border-gray-200 rounded-lg text-gray-600 hover:bg-white bg-white disabled:opacity-50">
                    {dismissingFailure === f.id ? '...' : 'ปิดการแจ้งเตือน'}
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Hub Management — pending hub-change requests + already-confirmed
            ones (with an editable ETA). Kept as one titled card with each
            list height-capped so it can't push the permissions section
            (the page's main purpose) far down the page. */}
        {(hubRequests.length > 0 || confirmedHubRequests.length > 0) && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 mb-8 shadow-sm">
            <h2 className="text-base font-bold text-gray-900 mb-4">📦 Hub Management</h2>

            {hubRequests.length > 0 && (
              <div className={confirmedHubRequests.length > 0 ? 'mb-5' : ''}>
                <h3 className="text-xs font-bold text-red-700 uppercase tracking-wide mb-2">รอการยืนยัน ({hubRequests.length})</h3>
                <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
                  {hubRequests.map(r => (
                    <div key={r.id} className="bg-red-50 rounded-lg border border-red-100 px-4 py-3 flex items-center justify-between gap-4 flex-wrap">
                      <div className="text-sm">
                        <span className="font-mono font-bold text-gray-800">{r.container_name}</span>
                        <span className="text-gray-400 mx-2">·</span>
                        <span className="text-gray-600">{r.invoice_no}</span>
                        <span className="text-gray-400 mx-2">→</span>
                        <span className="font-semibold text-orange-700">Hub {r.hub}</span>
                        <span className="text-xs text-gray-400 ml-2">โดย {r.requested_by}</span>
                      </div>
                      <div className="flex items-center gap-2 flex-wrap">
                        <a href={`/dashboard/${r.invoice_id}?container=${encodeURIComponent(r.container_name)}`} target="_blank" className="px-3 py-1 text-xs border border-gray-200 rounded-lg text-gray-600 hover:bg-white bg-white">ดู Invoice</a>
                        <div className="flex items-center gap-1">
                          <label className="text-xs text-gray-500 whitespace-nowrap">วันเข้าคลัง:</label>
                          <input
                            type="date"
                            value={confirmDates[`${r.invoice_id}:${r.container_name}`] ?? ''}
                            onChange={e => setConfirmDates(prev => ({ ...prev, [`${r.invoice_id}:${r.container_name}`]: e.target.value }))}
                            className="border border-gray-200 rounded px-2 py-0.5 text-xs outline-none focus:border-green-400 bg-white"
                          />
                        </div>
                        <button onClick={() => rejectHub(r)} disabled={rejecting === r.container_name || confirming === r.container_name} className="px-3 py-1 text-xs border border-red-200 text-red-500 rounded-lg hover:bg-red-100 bg-white font-medium disabled:opacity-50">✕ ปฏิเสธ</button>
                        <button onClick={() => confirmHub(r.invoice_id, r.container_name)} disabled={confirming === r.container_name || rejecting === r.container_name} className="px-3 py-1 text-xs bg-green-600 text-white rounded-lg hover:bg-green-700 font-medium disabled:opacity-50">{confirming === r.container_name ? '...' : '✓ ยืนยัน'}</button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {confirmedHubRequests.length > 0 && (
              <div>
                <h3 className="text-xs font-bold text-blue-700 uppercase tracking-wide mb-2">ยืนยันแล้ว ({confirmedHubRequests.length})</h3>
                <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
                  {confirmedHubRequests.map(r => {
                    const key = `${r.invoice_id}:${r.container_name}`
                    return (
                      <div key={r.id} className="bg-blue-50 rounded-lg border border-blue-100 px-4 py-3 flex items-center justify-between gap-4 flex-wrap">
                        <div className="text-sm">
                          <span className="font-mono font-bold text-gray-800">{r.container_name}</span>
                          <span className="text-gray-400 mx-2">·</span>
                          <span className="text-gray-600">{r.invoice_no}</span>
                          <span className="text-gray-400 mx-2">→</span>
                          <span className="font-semibold text-blue-700">Hub {r.hub}</span>
                          {r.confirmed_by && <span className="text-xs text-gray-400 ml-2">ยืนยันโดย {r.confirmed_by}</span>}
                        </div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <a href={`/dashboard/${r.invoice_id}?container=${encodeURIComponent(r.container_name)}`} target="_blank" className="px-3 py-1 text-xs border border-gray-200 rounded-lg text-gray-600 hover:bg-white bg-white">ดู Invoice</a>
                          <div className="flex items-center gap-1">
                            <label className="text-xs text-gray-500 whitespace-nowrap">วันเข้าคลัง:</label>
                            <input
                              type="date"
                              value={editDates[key] ?? r.hub_arrival_date ?? ''}
                              onChange={e => setEditDates(prev => ({ ...prev, [key]: e.target.value }))}
                              className="border border-gray-200 rounded px-2 py-0.5 text-xs outline-none focus:border-blue-400 bg-white"
                            />
                          </div>
                          <button
                            onClick={() => saveHubDate(r)}
                            disabled={savingDate === r.container_name || (editDates[key] ?? r.hub_arrival_date ?? '') === (r.hub_arrival_date ?? '')}
                            className="px-3 py-1 text-xs bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-medium disabled:opacity-50"
                          >
                            {savingDate === r.container_name ? '...' : 'บันทึก'}
                          </button>
                        </div>
                      </div>
                    )
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {/* Default pages */}
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-5 mb-8">
          <div className="flex items-center justify-between flex-wrap gap-3 mb-3">
            <div>
              <h2 className="text-sm font-bold text-amber-900">หน้าที่ทุกคนเข้าได้ (Default)</h2>
              <p className="text-xs text-amber-700 mt-0.5">บัญชีที่ไม่ได้อยู่ในรายการด้านล่างจะเข้าได้เฉพาะหน้าที่เลือกไว้นี้</p>
            </div>
            <button
              onClick={saveDefaultPages}
              disabled={savingDefault}
              className="px-4 py-1.5 bg-amber-600 text-white text-xs font-medium rounded-lg hover:bg-amber-700 disabled:opacity-50 transition-colors"
            >
              {savingDefault ? 'กำลังบันทึก...' : 'บันทึก Default'}
            </button>
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2">
            {ALL_PAGES.map(({ key, label }) => (
              <label key={key} className={`flex items-center gap-2 text-xs cursor-pointer select-none px-2 py-1.5 rounded-lg border transition-colors ${
                defaultPages.includes(key)
                  ? 'border-amber-400 bg-amber-100 text-amber-900'
                  : 'border-amber-200 bg-white text-gray-500 hover:border-amber-300'
              }`}>
                <input
                  type="checkbox"
                  checked={defaultPages.includes(key)}
                  onChange={e => toggleDefaultPage(key, e.target.checked)}
                  className="accent-amber-600"
                />
                {label}
              </label>
            ))}
          </div>
        </div>

        {/* Add user */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 mb-6 flex items-end gap-3 flex-wrap">
          <div className="flex-1 min-w-[220px]">
            <label className="block text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1">เพิ่มผู้ใช้ (สิทธิ์พิเศษ)</label>
            <input
              type="email"
              value={newEmail}
              onChange={e => { setNewEmail(e.target.value); setAddError('') }}
              onKeyDown={e => e.key === 'Enter' && addUser()}
              placeholder="email@rbs-groups.com"
              className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm outline-none focus:border-blue-400"
            />
            {addError && <p className="text-xs text-red-500 mt-1">{addError}</p>}
          </div>
          <button
            onClick={addUser}
            className="px-4 py-2 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700 transition-colors"
          >
            + เพิ่ม
          </button>
        </div>

        {loading ? (
          <p className="text-sm text-gray-400">กำลังโหลด...</p>
        ) : rows.length === 0 ? (
          <p className="text-sm text-gray-400">ยังไม่มีผู้ใช้ในระบบสิทธิ์พิเศษ</p>
        ) : (
          <div className="space-y-4">
            {rows.map(row => (
              <div key={row.email} className="bg-white rounded-xl border border-gray-200 shadow-sm p-5">
                <div className="flex items-center gap-4 flex-wrap mb-4">
                  <span className="font-medium text-gray-900 text-sm">{row.email}</span>

                  <label className="flex items-center gap-1.5 text-xs cursor-pointer select-none">
                    <input
                      type="checkbox"
                      checked={row.is_admin}
                      onChange={e => toggleAdmin(row.email, e.target.checked)}
                      className="accent-purple-600"
                    />
                    <span className="font-semibold text-purple-700">Admin</span>
                  </label>

                  <label className="flex items-center gap-1.5 text-xs cursor-pointer select-none ml-2">
                    <input
                      type="checkbox"
                      checked={row.allowed_pages === null}
                      onChange={e => setAllPages(row.email, e.target.checked)}
                      disabled={row.is_admin}
                      className="accent-blue-600"
                    />
                    <span className="text-gray-600">เข้าได้ทุกหน้า</span>
                  </label>

                  <div className="ml-auto flex items-center gap-2">
                    <button
                      onClick={() => upsert(row)}
                      disabled={saving === row.email}
                      className="px-3 py-1 bg-green-600 text-white text-xs font-medium rounded-lg hover:bg-green-700 disabled:opacity-50 transition-colors"
                    >
                      {saving === row.email ? 'กำลังบันทึก...' : 'บันทึก'}
                    </button>
                    <button
                      onClick={() => removeUser(row.email)}
                      className="px-3 py-1 border border-red-200 text-red-500 text-xs rounded-lg hover:bg-red-50 transition-colors"
                    >
                      ลบ
                    </button>
                  </div>
                </div>

                {!row.is_admin && (
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2">
                    {ALL_PAGES.map(({ key, label }) => (
                      <label key={key} className={`flex items-center gap-2 text-xs cursor-pointer select-none px-2 py-1.5 rounded-lg border transition-colors ${
                        isPageChecked(row, key)
                          ? 'border-blue-300 bg-blue-50 text-blue-800'
                          : 'border-gray-200 text-gray-500 hover:border-gray-300'
                      }`}>
                        <input
                          type="checkbox"
                          checked={isPageChecked(row, key)}
                          onChange={e => togglePage(row.email, key, e.target.checked)}
                          disabled={row.allowed_pages === null}
                          className="accent-blue-600"
                        />
                        {label}
                      </label>
                    ))}
                  </div>
                )}
                {row.is_admin && (
                  <p className="text-xs text-purple-600 bg-purple-50 rounded-lg px-3 py-2 border border-purple-100">
                    Admin มีสิทธิ์เข้าถึงทุกหน้าและจัดการสิทธิ์ผู้อื่นได้
                  </p>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}

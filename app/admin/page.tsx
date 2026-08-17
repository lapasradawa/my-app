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
    setLoading(false)
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

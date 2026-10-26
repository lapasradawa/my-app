'use client'

import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { IBM_Plex_Sans_Thai, Barlow_Condensed } from 'next/font/google'
import { supabase } from '@/lib/supabase'
import NavBar from '@/components/NavBar'
import { HUB_COLORS, hubColor } from '@/lib/hub-colors'

const bodyFont = IBM_Plex_Sans_Thai({ subsets: ['thai', 'latin'], weight: ['400', '500', '600', '700'] })
const numFont = Barlow_Condensed({ subsets: ['latin'], weight: ['500', '600', '700'] })

const DEFAULT_HUB = 'มัยลาภ'
const REAL_HUBS = Object.keys(HUB_COLORS).filter(h => h !== DEFAULT_HUB) // ['ขอนแก่น','พิษณุโลก','สุราษฎร์ธานี']
const CAPACITY_KEY = 'warehouse_daily_capacity'

// ── Date helpers (Monday-first) ─────────────────────────────────────────────
function ds(d: Date) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}` }
function pd(s: string) { return new Date(s + 'T00:00:00') }
function dow(d: Date) { return (d.getDay() + 6) % 7 } // Mon=0 … Sun=6
function isSunday(d: Date) { return d.getDay() === 0 }
function addDays(d: Date, n: number) { const r = new Date(d); r.setDate(r.getDate() + n); return r }
function sameDay(a: Date, b: Date) { return ds(a) === ds(b) }

const MONTHS_TH = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
function monthLabel(d: Date) { return `${MONTHS_TH[d.getMonth()]} ${d.getFullYear() + 543}` }

// ── Per-invoice color (deterministic hash, stable for everyone) ────────────
const INVOICE_PALETTE = [
  { bg: '#1d4ed8', text: '#ffffff' }, { bg: '#c2410c', text: '#ffffff' },
  { bg: '#15803d', text: '#ffffff' }, { bg: '#7e22ce', text: '#ffffff' },
  { bg: '#b91c1c', text: '#ffffff' }, { bg: '#0e7490', text: '#ffffff' },
  { bg: '#854d0e', text: '#ffffff' }, { bg: '#4338ca', text: '#ffffff' },
  { bg: '#be185d', text: '#ffffff' }, { bg: '#166534', text: '#ffffff' },
]
function invoiceColor(invoiceNo: string) {
  let h = 0
  for (let i = 0; i < invoiceNo.length; i++) h = (h * 31 + invoiceNo.charCodeAt(i)) >>> 0
  return INVOICE_PALETTE[h % INVOICE_PALETTE.length]
}

// ── Data shapes ──────────────────────────────────────────────────────────
interface InvoiceRow {
  id: string
  invoice_no: string
  supplier: string | null
  estimated_arrival: string | null
  estimated_arrival_end: string | null
  container_names: string[] | null
}
interface HubReqRow {
  invoice_id: string
  invoice_no: string
  container_name: string
  hub: string
  status: string
  hub_arrival_date: string | null
}
interface SizeRow {
  invoice_id: string
  container_name: string
  is_long: boolean
}
interface ContainerUnit {
  invoiceId: string
  invoiceNo: string
  supplier: string | null
  containerName: string
  label: string // what's printed on the block
  isLong: boolean
  scheduledDate: string | null // hub_arrival_date, only when hub === DEFAULT_HUB — drives capacity + ETA sync
  hubDate: string | null // hub_arrival_date regardless of hub — which day cell a hub-routed block renders under
  hub: string
  estimatedArrival: string | null
  estimatedArrivalEnd: string | null
}

function unitKey(u: { invoiceId: string; containerName: string }) { return `${u.invoiceId}::${u.containerName}` }

// ── Theme tokens (concrete-yard theme, light/dark) ──────────────────────────
const LIGHT = {
  pageBg: '#e4e4e0', panelBg: '#f2f2ef', cardBg: '#ffffff',
  text: '#1a2a3a', textMuted: '#5a6a7a', border: '#c8cdd2',
  yellow: '#f5c518', yellowDark: '#a88300',
  dashedBorder: '#9aa5ad', sundayBg: '#eceae4', todayRing: '#1a2a3a',
  overBg: '#fde2e2', overBorder: '#dc2626',
}
const DARK = {
  pageBg: '#15181b', panelBg: '#1d2227', cardBg: '#262b31',
  text: '#e7ebee', textMuted: '#9aa7b0', border: '#394049',
  yellow: '#f5c518', yellowDark: '#d4a60a',
  dashedBorder: '#4b545d', sundayBg: '#1a1e22', todayRing: '#f5c518',
  overBg: '#3a1f1f', overBorder: '#ef4444',
}

// Diagonal "hazard stripe" texture, used sparingly for accent elements only.
function stripeBg(color: string) {
  return `repeating-linear-gradient(45deg, ${color} 0 10px, transparent 10px 20px)`
}

export default function ContainerManagementPage() {
  const [invoices, setInvoices] = useState<InvoiceRow[]>([])
  const [hubReqs, setHubReqs] = useState<HubReqRow[]>([])
  const [sizes, setSizes] = useState<SizeRow[]>([])
  const [loading, setLoading] = useState(true)
  const [capacity, setCapacity] = useState(3)
  const [monthDate, setMonthDate] = useState(() => { const d = new Date(); d.setDate(1); return d })
  const [dark, setDark] = useState(false)
  const [heldKey, setHeldKey] = useState<string | null>(null)
  // Position captured from the clicked icon's own bounding box, and rendered
  // via a portal straight into document.body — the calendar's horizontal
  // scroll wrapper has overflowX set, which per the CSS spec silently forces
  // overflowY to clip too, so a popover nested/anchored inside it (instead
  // of escaping via a portal) was getting cut off/covered instead of
  // floating visibly above the grid.
  const [hubPopover, setHubPopover] = useState<{ day: string; hub: string; x: number; y: number } | null>(null)
  const [toasts, setToasts] = useState<{ id: number; text: string; warn?: boolean }[]>([])
  const [userEmail, setUserEmail] = useState<string | null>(null)
  const dragKeyRef = useRef<string | null>(null)
  const toastIdRef = useRef(0)

  useEffect(() => {
    const saved = typeof window !== 'undefined' ? window.localStorage.getItem('cm-dark') : null
    if (saved === '1') setDark(true)
    supabase.auth.getSession().then(({ data }) => setUserEmail(data.session?.user?.email ?? null))
  }, [])

  function toggleDark() {
    setDark(v => {
      const next = !v
      try { window.localStorage.setItem('cm-dark', next ? '1' : '0') } catch { /* ignore */ }
      return next
    })
  }

  function toast(text: string, warn = false) {
    const id = ++toastIdRef.current
    setToasts(prev => [...prev, { id, text, warn }])
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), 3500)
  }

  async function load() {
    const [{ data: inv }, { data: hubs }, { data: sz }, { data: settings }] = await Promise.all([
      supabase.from('invoices').select('id, invoice_no, supplier, estimated_arrival, estimated_arrival_end, container_names'),
      supabase.from('container_hub_requests').select('invoice_id, invoice_no, container_name, hub, status, hub_arrival_date').eq('status', 'confirmed'),
      supabase.from('container_sizes').select('invoice_id, container_name, is_long'),
      supabase.from('cost_settings').select('key, value').eq('key', CAPACITY_KEY),
    ])
    setInvoices((inv ?? []) as InvoiceRow[])
    setHubReqs((hubs ?? []) as HubReqRow[])
    setSizes((sz ?? []) as SizeRow[])
    const capSetting = settings?.[0]?.value
    if (capSetting) setCapacity(Math.max(1, parseInt(capSetting, 10) || 3))
    setLoading(false)
  }

  useEffect(() => { load() }, [])

  // ── Derived: one ContainerUnit per real (or synthetic) container ─────────
  const units = useMemo((): ContainerUnit[] => {
    const hubMap = new Map<string, HubReqRow>()
    for (const h of hubReqs) hubMap.set(`${h.invoice_id}::${h.container_name}`, h)
    const sizeMap = new Map<string, boolean>()
    for (const s of sizes) sizeMap.set(`${s.invoice_id}::${s.container_name}`, s.is_long)

    const list: ContainerUnit[] = []
    for (const inv of invoices) {
      const hasReal = (inv.container_names ?? []).length > 0
      const names = hasReal ? inv.container_names! : [inv.invoice_no]
      for (const name of names) {
        const key = `${inv.id}::${name}`
        const req = hubMap.get(key)
        list.push({
          invoiceId: inv.id,
          invoiceNo: inv.invoice_no,
          supplier: inv.supplier,
          containerName: name,
          label: hasReal ? name : inv.invoice_no,
          isLong: sizeMap.get(key) ?? false,
          scheduledDate: req && req.hub === DEFAULT_HUB ? req.hub_arrival_date : null,
          hubDate: req?.hub_arrival_date ?? null,
          hub: req?.hub ?? DEFAULT_HUB,
          estimatedArrival: inv.estimated_arrival,
          estimatedArrivalEnd: inv.estimated_arrival_end,
        })
      }
    }
    return list
  }, [invoices, hubReqs, sizes])

  const unitByKey = useMemo(() => {
    const m = new Map<string, ContainerUnit>()
    for (const u of units) m.set(unitKey(u), u)
    return m
  }, [units])

  const byInvoice = useMemo(() => {
    const m = new Map<string, ContainerUnit[]>()
    for (const u of units) {
      const arr = m.get(u.invoiceId) ?? []
      arr.push(u)
      m.set(u.invoiceId, arr)
    }
    return m
  }, [units])

  // Tray scope: only invoices relevant to the month currently being planned
  // (matching the calendar view), so an old invoice that already arrived
  // months ago doesn't sit around cluttering the tray forever. An invoice
  // with no ETA at all has nothing to match against, so it's always shown —
  // better surfaced than silently lost.
  const today = useMemo(() => { const d = new Date(); d.setHours(0, 0, 0, 0); return d }, [])
  const todayStr = useMemo(() => ds(today), [today])

  const trayInvoices = useMemo(() => {
    return Array.from(byInvoice.entries()).filter(([, list]) => {
      const eta = list[0]?.estimatedArrival
      if (!eta) return true
      // วันที่ประมาณการเข้าคลังผ่านไปแล้ว — ไม่ต้องเอามาใส่ถาดอีก ไม่ว่าจะ
      // ลากตู้ไหนไปวางแล้วกี่ตู้ก็ตาม (Weekly Inbound/Calendar ยัง fallback
      // ไปใช้วันนี้เหมือนเดิมสำหรับตู้ที่ยังไม่ถูกจัดอย่างชัดเจน)
      if (eta < todayStr) return false
      const d = pd(eta)
      return d.getFullYear() === monthDate.getFullYear() && d.getMonth() === monthDate.getMonth()
    })
  }, [byInvoice, monthDate, todayStr])

  const byDay = useMemo(() => {
    const m = new Map<string, ContainerUnit[]>()
    for (const u of units) {
      if (u.hub === DEFAULT_HUB && u.scheduledDate) {
        const arr = m.get(u.scheduledDate) ?? []
        arr.push(u)
        m.set(u.scheduledDate, arr)
      }
    }
    return m
  }, [units])

  // Hub-routed units keyed by `${date}::${hub}`, so each calendar day cell can
  // show its own small hub icons instead of one shared row for the whole
  // month (a hub-routed request made before this page existed, with no
  // hub_arrival_date yet, has no key here — it stays visible/draggable in
  // the tray below until given an actual day).
  const byHubDay = useMemo(() => {
    const m = new Map<string, ContainerUnit[]>()
    for (const u of units) {
      if (u.hub !== DEFAULT_HUB && u.hubDate) {
        const key = `${u.hubDate}::${u.hub}`
        const arr = m.get(key) ?? []
        arr.push(u)
        m.set(key, arr)
      }
    }
    return m
  }, [units])

  const isPlaced = useCallback((u: ContainerUnit) =>
    (u.hub === DEFAULT_HUB && !!u.scheduledDate) || (u.hub !== DEFAULT_HUB && !!u.hubDate)
  , [])

  // ── Calendar grid for the current month (Monday-first, full weeks) ──────
  const monthCells = useMemo(() => {
    const first = new Date(monthDate)
    const firstMon = addDays(first, -dow(first))
    const lastDay = new Date(first.getFullYear(), first.getMonth() + 1, 0)
    const lastSun = addDays(lastDay, 6 - dow(lastDay))
    const cells: Date[] = []
    for (let d = new Date(firstMon); d <= lastSun; d = addDays(d, 1)) cells.push(new Date(d))
    return cells
  }, [monthDate])

  // Days in the visible month that are over capacity (Sundays excluded, they
  // have no receiving capacity at all so "over capacity" doesn't apply there).
  const overCapacityDays = useMemo(() => {
    let n = 0
    for (const d of monthCells) {
      if (d.getMonth() !== monthDate.getMonth()) continue
      if (isSunday(d)) continue
      if ((byDay.get(ds(d))?.length ?? 0) > capacity) n++
    }
    return n
  }, [monthCells, monthDate, byDay, capacity])

  // Scoped to the same invoices shown in the tray (this month's plan), not
  // every container in the whole database.
  const kpi = useMemo(() => {
    const scoped = trayInvoices.flatMap(([, list]) => list)
    return {
      total: scoped.length,
      placed: scoped.filter(isPlaced).length,
      pending: scoped.filter(u => !isPlaced(u)).length,
    }
  }, [trayInvoices, isPlaced])

  // ── Writes ────────────────────────────────────────────────────────────
  async function syncInvoiceEta(invoiceId: string) {
    const { data } = await supabase
      .from('container_hub_requests')
      .select('hub_arrival_date')
      .eq('invoice_id', invoiceId).eq('status', 'confirmed').eq('hub', DEFAULT_HUB)
      .not('hub_arrival_date', 'is', null)
    const dates = (data ?? []).map(r => r.hub_arrival_date as string).filter(Boolean).sort()
    // Only ever widen/update when there's something scheduled — never clear an
    // ETA that was set independently on Dashboard just because the tray is
    // currently empty of explicit placements.
    if (dates.length === 0) return
    await supabase.from('invoices').update({
      estimated_arrival: dates[0], estimated_arrival_end: dates[dates.length - 1],
    }).eq('id', invoiceId)
  }

  async function scheduleUnit(u: ContainerUnit, date: string | null, hub: string) {
    const { error } = await supabase.from('container_hub_requests').upsert({
      invoice_id: u.invoiceId, invoice_no: u.invoiceNo, container_name: u.containerName,
      hub, status: 'confirmed', hub_arrival_date: date,
      requested_by: userEmail ?? 'container-management', confirmed_by: userEmail ?? 'container-management',
      confirmed_at: new Date().toISOString(),
    }, { onConflict: 'invoice_id,container_name' })
    if (error) { toast(`บันทึกไม่สำเร็จ: ${error.message}`, true); return }
    if (hub === DEFAULT_HUB) await syncInvoiceEta(u.invoiceId)
    await load()
    // A drag-out from inside the hub popover (moving one of its own
    // containers elsewhere) never otherwise closes it — its full-screen
    // invisible backdrop would then keep silently swallowing every
    // subsequent drag/drop on the page, since it sits above everything else.
    setHubPopover(null)
  }

  async function unscheduleUnit(u: ContainerUnit) {
    const { error } = await supabase.from('container_hub_requests')
      .delete().eq('invoice_id', u.invoiceId).eq('container_name', u.containerName)
    if (error) { toast(`ลบไม่สำเร็จ: ${error.message}`, true); return }
    await syncInvoiceEta(u.invoiceId)
    await load()
    setHubPopover(null)
  }

  async function toggleSize(u: ContainerUnit) {
    await supabase.from('container_sizes').upsert({
      invoice_id: u.invoiceId, container_name: u.containerName, is_long: !u.isLong,
    }, { onConflict: 'invoice_id,container_name' })
    await load()
  }

  const [capacityInput, setCapacityInput] = useState('3')
  useEffect(() => { setCapacityInput(String(capacity)) }, [capacity])
  async function saveCapacity() {
    const v = Math.max(1, parseInt(capacityInput, 10) || 1)
    setCapacity(v)
    await supabase.from('cost_settings').upsert({ key: CAPACITY_KEY, value: String(v) })
  }

  // ── Place/move/remove a unit, with toasts + capacity warning ───────────
  async function placeOnDay(u: ContainerUnit, day: Date) {
    // Sunday is normally closed — still allowed, but only as a deliberate
    // emergency override (confirm dialog), never a plain accidental drop.
    if (isSunday(day)) {
      const ok = window.confirm(`วันที่ ${day.getDate()} ${MONTHS_TH[day.getMonth()]} เป็นวันอาทิตย์ (คลังปิด)\nยืนยันว่าเป็นกรณีฉุกเฉินและต้องการวางตู้วันนี้จริงหรือไม่?`)
      if (!ok) return
    }
    const dayStr = ds(day)
    const existing = byDay.get(dayStr)?.filter(x => unitKey(x) !== unitKey(u)) ?? []
    await scheduleUnit(u, dayStr, DEFAULT_HUB)
    toast(`${u.invoiceNo} ตู้ ${u.label} → ${day.getDate()} ${MONTHS_TH[day.getMonth()]}${isSunday(day) ? ' (ฉุกเฉิน — วันคลังปิด)' : ''}`)
    if (!isSunday(day) && existing.length + 1 > capacity) toast(`⚠ วันที่ ${day.getDate()} ${MONTHS_TH[day.getMonth()]} เกินกำลังรับแล้ว (${existing.length + 1}/${capacity})`, true)
  }
  async function placeOnHub(u: ContainerUnit, hub: string, day: Date) {
    await scheduleUnit(u, ds(day), hub)
    toast(`${u.invoiceNo} ตู้ ${u.label} → Hub ${hub} (${day.getDate()} ${MONTHS_TH[day.getMonth()]})`)
  }
  async function returnToTray(u: ContainerUnit) {
    await unscheduleUnit(u)
    toast(`${u.invoiceNo} ตู้ ${u.label} → กลับเข้าถาด`)
  }

  // ── Auto-schedule a whole invoice's unplaced units onto their own ETA ───
  async function autoSchedule(invoiceId: string) {
    const list = byInvoice.get(invoiceId) ?? []
    const unplaced = list.filter(u => !isPlaced(u))
    if (unplaced.length === 0) return
    // Local working copy of day loads so consecutive placements in this same
    // run see each other's capacity usage immediately.
    const dayLoad = new Map<string, number>()
    for (const [day, arr] of byDay) dayLoad.set(day, arr.length)
    for (const u of unplaced) {
      let d = u.estimatedArrival ? pd(u.estimatedArrival) : new Date()
      while (true) {
        if (!isSunday(d) && (dayLoad.get(ds(d)) ?? 0) < capacity) break
        d = addDays(d, 1)
      }
      const dayStr = ds(d)
      dayLoad.set(dayStr, (dayLoad.get(dayStr) ?? 0) + 1)
      await scheduleUnit(u, dayStr, DEFAULT_HUB)
    }
    toast(`จัดอัตโนมัติ ${unplaced.length} ตู้ ของ ${list[0]?.invoiceNo} เรียบร้อย`)
  }

  // ── Drag & drop (desktop) ───────────────────────────────────────────────
  function onDragStart(key: string) { dragKeyRef.current = key }
  // If a drag is dropped somewhere with no drop handler (abandoned), nothing
  // else clears the ref — a later tap-to-hold placement would otherwise read
  // this stale key instead of the one actually held.
  function onDragEnd() { dragKeyRef.current = null }
  function onDropDay(day: Date) {
    const key = dragKeyRef.current ?? heldKey
    dragKeyRef.current = null
    if (!key) return
    const u = unitByKey.get(key)
    if (u) placeOnDay(u, day)
    setHeldKey(null)
  }
  function onDropHub(hub: string, day: Date) {
    const key = dragKeyRef.current ?? heldKey
    dragKeyRef.current = null
    if (!key) return
    const u = unitByKey.get(key)
    if (u) placeOnHub(u, hub, day)
    setHeldKey(null)
  }
  function onDropTray() {
    const key = dragKeyRef.current ?? heldKey
    dragKeyRef.current = null
    if (!key) return
    const u = unitByKey.get(key)
    if (u) returnToTray(u)
    setHeldKey(null)
  }

  // ── Tap-to-hold (mobile + keyboard) ─────────────────────────────────────
  function onBlockActivate(key: string) {
    setHeldKey(prev => prev === key ? null : key)
  }
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (e.key === 'Escape') {
        if (heldKey) setHeldKey(null)
        if (hubPopover) setHubPopover(null)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [heldKey, hubPopover])

  const heldUnit = heldKey ? unitByKey.get(heldKey) ?? null : null

  // ── CSV export ───────────────────────────────────────────────────────────
  function exportCsv() {
    const placedSorted = units.filter(isPlaced).sort((a, b) => (a.scheduledDate ?? '9999').localeCompare(b.scheduledDate ?? '9999'))
    const unplaced = units.filter(u => !isPlaced(u))
    const header = ['วันที่เข้าคลัง', 'Invoice', 'Supplier', 'ตู้ที่', 'จำนวนตู้ใน Invoice', 'ขนาด', 'ETA']
    const rows: string[][] = []
    for (const u of placedSorted) {
      const total = byInvoice.get(u.invoiceId)?.length ?? 1
      rows.push([
        u.hub === DEFAULT_HUB ? (u.scheduledDate ?? '') : `Hub ${u.hub}`,
        u.invoiceNo, u.supplier ?? '', u.label, String(total), u.isLong ? '40ft' : '20ft',
        u.estimatedArrival ?? '',
      ])
    }
    for (const u of unplaced) {
      const total = byInvoice.get(u.invoiceId)?.length ?? 1
      rows.push(['ยังไม่จัด', u.invoiceNo, u.supplier ?? '', u.label, String(total), u.isLong ? '40ft' : '20ft', u.estimatedArrival ?? ''])
    }
    const csv = [header, ...rows].map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\r\n')
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url; a.download = `container-plan-${ds(new Date())}.csv`; a.click()
    URL.revokeObjectURL(url)
  }

  const t = dark ? DARK : LIGHT

  return (
    <div className={bodyFont.className} style={{ background: t.pageBg, minHeight: '100vh', color: t.text }}>
      <NavBar />

      {/* ── Header ── */}
      <div style={{ background: t.panelBg, borderBottom: `4px solid transparent`, backgroundImage: `linear-gradient(${t.panelBg},${t.panelBg}), ${stripeBg(t.yellow)}`, backgroundOrigin: 'border-box', backgroundClip: 'padding-box, border-box', padding: '16px 20px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap', maxWidth: 1600, margin: '0 auto' }}>
          <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>แผนรับตู้เข้าคลัง</h1>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <button onClick={() => setMonthDate(d => { const n = new Date(d); n.setMonth(n.getMonth() - 1); return n })}
              style={{ width: 28, height: 28, border: `1px solid ${t.border}`, borderRadius: 6, background: t.cardBg, color: t.text, cursor: 'pointer' }}>‹</button>
            <span className={numFont.className} style={{ fontSize: 17, fontWeight: 600, minWidth: 150, textAlign: 'center' }}>{monthLabel(monthDate)}</span>
            <button onClick={() => setMonthDate(d => { const n = new Date(d); n.setMonth(n.getMonth() + 1); return n })}
              style={{ width: 28, height: 28, border: `1px solid ${t.border}`, borderRadius: 6, background: t.cardBg, color: t.text, cursor: 'pointer' }}>›</button>
            <button onClick={() => setMonthDate(() => { const d = new Date(); d.setDate(1); return d })} title="กลับเดือนนี้"
              style={{ width: 28, height: 28, border: `1px solid ${t.border}`, borderRadius: '50%', background: t.yellow, color: '#1a1a1a', cursor: 'pointer', fontSize: 10 }}>●</button>
          </div>
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13, color: t.textMuted }}>
            คลังรับได้
            <input value={capacityInput} onChange={e => setCapacityInput(e.target.value.replace(/\D/g, ''))} onBlur={saveCapacity}
              className={numFont.className}
              style={{ width: 44, textAlign: 'center', border: `1px solid ${t.border}`, borderRadius: 6, padding: '3px 4px', background: t.cardBg, color: t.text, fontSize: 15 }} />
            ตู้/วัน
          </label>
          <button onClick={toggleDark} title="สลับโหมดมืด/สว่าง"
            style={{ width: 28, height: 28, border: `1px solid ${t.border}`, borderRadius: 6, background: t.cardBg, color: t.text, cursor: 'pointer' }}>
            {dark ? '☀' : '☾'}
          </button>
          <div style={{ marginLeft: 'auto', display: 'flex', gap: 18 }}>
            <Stat label="ตู้ทั้งหมด" value={kpi.total} t={t} />
            <Stat label="จัดแล้ว" value={kpi.placed} t={t} />
            <Stat label="รอจัด" value={kpi.pending} t={t} />
            <Stat label="วันที่เกินกำลัง" value={overCapacityDays} t={t} danger={overCapacityDays > 0} />
          </div>
          <button onClick={exportCsv}
            style={{ padding: '8px 14px', borderRadius: 8, border: 'none', background: t.text, color: t.panelBg, fontWeight: 600, fontSize: 12, cursor: 'pointer' }}>
            ↓ ดาวน์โหลดตาราง (CSV)
          </button>
        </div>

        {/* Hub legend — maps the tiny 🏭 icons inside each day cell to a hub
            name/color, since the icons alone are too small to label. */}
        <div style={{ maxWidth: 1600, margin: '8px auto 0', display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap' }}>
          <span style={{ fontSize: 10, color: t.textMuted, fontWeight: 700 }}>🏭 Hub ในปฏิทิน:</span>
          {REAL_HUBS.map(hub => {
            const hc = hubColor(hub)
            return (
              <span key={hub} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 11, fontWeight: 600, color: t.textMuted }}>
                <span style={{ width: 16, height: 16, borderRadius: '50%', background: hc.dot, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 9 }}>🏭</span>
                {hub}
              </span>
            )
          })}
        </div>
      </div>

      {/* ── Held-unit bar (mobile/keyboard) ── */}
      {heldUnit && (
        <div style={{ position: 'sticky', top: 0, zIndex: 40, background: t.yellow, color: '#1a1a1a', padding: '8px 16px', display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', backgroundImage: stripeBg('rgba(0,0,0,0.06)') }}>
          <span style={{ fontWeight: 700, fontSize: 13 }}>📦 กำลังถือ: {heldUnit.invoiceNo} ตู้ {heldUnit.label}</span>
          <span style={{ fontSize: 12, opacity: 0.8 }}>แตะวันที่ปลายทาง หรือ icon Hub เพื่อวาง</span>
          <div style={{ marginLeft: 'auto', display: 'flex', gap: 8 }}>
            <button onClick={() => returnToTray(heldUnit)} style={{ padding: '4px 10px', borderRadius: 6, border: '1px solid #1a1a1a', background: 'transparent', color: '#1a1a1a', fontSize: 12, cursor: 'pointer' }}>คืนเข้าถาด</button>
            <button onClick={() => setHeldKey(null)} style={{ padding: '4px 10px', borderRadius: 6, border: 'none', background: '#1a1a1a', color: t.yellow, fontSize: 12, cursor: 'pointer' }}>ยกเลิก (Esc)</button>
          </div>
        </div>
      )}

      <div style={{ maxWidth: 1600, margin: '0 auto', padding: 16, display: 'flex', gap: 16, flexDirection: 'column' }}>
        <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', alignItems: 'flex-start' }}>
          {/* ── Tray ── */}
          {/* Sticky + its own scroll region: scrolling through a long invoice
              list shouldn't carry the calendar out of view along with it. */}
          <div
            onDragOver={e => e.preventDefault()}
            onDrop={onDropTray}
            style={{
              flex: '1 1 320px', minWidth: 300, maxWidth: 380, display: 'flex', flexDirection: 'column', gap: 12,
              position: 'sticky', top: 12, alignSelf: 'flex-start',
              maxHeight: 'calc(100vh - 24px)', overflowY: 'auto', paddingRight: 4,
            }}
          >
            <div style={{ background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 10, padding: 12 }}>
              <p style={{ fontWeight: 700, fontSize: 14, margin: 0 }}>Invoice รอจัด</p>
              <p style={{ fontSize: 11, color: t.textMuted, margin: '4px 0 0' }}>ลากตู้ไปวางในวันที่ต้องการ หรือแตะตู้หนึ่งครั้งแล้วแตะวันที่ ลากตู้กลับมาที่ถาดนี้เพื่อเอาออก</p>
            </div>
            {loading ? (
              <p style={{ color: t.textMuted, fontSize: 13 }}>กำลังโหลด...</p>
            ) : (
              trayInvoices.map(([invId, list]) => {
                const first = list[0]
                const color = invoiceColor(first.invoiceNo)
                const placedCount = list.filter(isPlaced).length
                const allPlaced = placedCount === list.length
                return (
                  <div key={invId} style={{ background: t.cardBg, border: `1px solid ${t.border}`, borderLeft: `5px solid ${color.bg}`, borderRadius: 10, padding: 12 }}>
                    <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, flexWrap: 'wrap' }}>
                      <span style={{ fontWeight: 700, fontSize: 14 }}>{first.invoiceNo}</span>
                      <span style={{ fontSize: 11, color: t.textMuted }}>{first.supplier ?? '—'}</span>
                    </div>
                    <div className={numFont.className} style={{ fontSize: 12, color: t.textMuted, marginTop: 2 }}>
                      {list.length} ตู้ · ETA {first.estimatedArrival ? pd(first.estimatedArrival).toLocaleDateString('th-TH', { day: 'numeric', month: 'short' }) : '—'} · จัดแล้ว {placedCount}/{list.length}
                    </div>
                    <div style={{ height: 5, borderRadius: 3, background: t.border, marginTop: 6, overflow: 'hidden' }}>
                      <div style={{ height: '100%', width: `${list.length ? (placedCount / list.length) * 100 : 0}%`, background: color.bg, transition: 'width 0.25s' }} />
                    </div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 10 }}>
                      {list.map(u => (
                        <Block key={unitKey(u)} u={u} color={color} solid={!isPlaced(u)} held={heldKey === unitKey(u)}
                          numFontClass={numFont.className}
                          onDragStart={() => onDragStart(unitKey(u))}
                          onDragEnd={onDragEnd}
                          onClick={() => onBlockActivate(unitKey(u))}
                          onContextMenu={e => { e.preventDefault(); toggleSize(u) }}
                        />
                      ))}
                    </div>
                    {allPlaced ? (
                      <p style={{ color: '#16a34a', fontSize: 12, fontWeight: 700, marginTop: 8 }}>✓ จัดครบแล้ว</p>
                    ) : (
                      <button onClick={() => autoSchedule(invId)}
                        style={{ marginTop: 10, padding: '6px 10px', borderRadius: 7, border: `1px solid ${t.border}`, background: t.panelBg, color: t.text, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                        จัดอัตโนมัติ
                      </button>
                    )}
                  </div>
                )
              })
            )}
          </div>

          {/* ── Calendar ── */}
          <div style={{ flex: '3 1 640px', minWidth: 300 }}>
            <div style={{ overflowX: 'auto' }}>
              <div onClick={() => setHubPopover(null)} style={{ display: 'grid', gridTemplateColumns: 'repeat(7, minmax(110px, 1fr))', gap: 6, minWidth: 780 }}>
                {['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสฯ', 'ศุกร์', 'เสาร์', 'อาทิตย์'].map(d => (
                  <div key={d} style={{ fontSize: 11, fontWeight: 700, color: t.textMuted, textAlign: 'center', padding: '4px 0' }}>{d}</div>
                ))}
                {monthCells.map(day => {
                  const dayStr = ds(day)
                  const inMonth = day.getMonth() === monthDate.getMonth()
                  const sunday = isSunday(day)
                  const load = byDay.get(dayStr)?.length ?? 0
                  const over = !sunday && load > capacity
                  const isToday = sameDay(day, today)
                  const etaInvoices = invoices.filter(inv => inv.estimated_arrival === dayStr)
                  return (
                    <div
                      key={dayStr}
                      onDragOver={e => e.preventDefault()}
                      onDrop={() => onDropDay(day)}
                      onClick={() => { if (heldKey) { const u = unitByKey.get(heldKey); if (u) placeOnDay(u, day) } }}
                      style={{
                        position: 'relative',
                        background: sunday ? t.sundayBg : over ? t.overBg : t.cardBg,
                        border: `1px solid ${over ? t.overBorder : t.border}`,
                        borderRadius: 8, padding: 6, minHeight: 100, opacity: inMonth ? 1 : 0.45,
                        boxShadow: isToday ? `0 0 0 2px ${t.todayRing}` : 'none',
                        cursor: heldKey ? 'pointer' : 'default',
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
                        <span className={numFont.className} style={{ fontWeight: isToday ? 800 : 600, fontSize: 15 }}>{day.getDate()}</span>
                        {sunday ? (
                          <span style={{ fontSize: 9, color: t.textMuted }}>คลังปิด</span>
                        ) : (
                          <span className={numFont.className} style={{ fontSize: 10, color: over ? t.overBorder : t.textMuted, fontWeight: 600 }}>{load}/{capacity}</span>
                        )}
                      </div>
                      {!sunday && (
                        <div style={{ display: 'flex', gap: 2, marginTop: 3 }}>
                          {Array.from({ length: Math.min(capacity, 8) }, (_, i) => (
                            <div key={i} style={{ flex: 1, height: 4, borderRadius: 2, background: i < load ? (over ? t.overBorder : t.yellow) : t.border }} />
                          ))}
                        </div>
                      )}
                      {etaInvoices.map(inv => (
                        <div key={inv.id} className={numFont.className} style={{ marginTop: 4, fontSize: 9, fontWeight: 700, color: '#1a1a1a', background: t.yellow, borderRadius: 4, padding: '1px 5px', display: 'inline-block' }}>
                          ETA {inv.invoice_no}
                        </div>
                      ))}
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 3, marginTop: 4 }}>
                        {(byDay.get(dayStr) ?? []).map(u => (
                          <Block key={unitKey(u)} u={u} color={invoiceColor(u.invoiceNo)} solid held={heldKey === unitKey(u)} small
                            numFontClass={numFont.className}
                            onDragStart={() => onDragStart(unitKey(u))}
                          onDragEnd={onDragEnd}
                            onClick={e => { e.stopPropagation(); onBlockActivate(unitKey(u)) }}
                            onContextMenu={e => { e.preventDefault(); e.stopPropagation(); toggleSize(u) }}
                            warn={!!u.scheduledDate && !!u.estimatedArrival && !!u.estimatedArrivalEnd && (u.scheduledDate < u.estimatedArrival || u.scheduledDate > u.estimatedArrivalEnd)}
                          />
                        ))}
                      </div>
                      {/* Per-day hub icons — dropping here sets hub + this exact
                          day in one action, so it's always clear which day a
                          hub-routed container is going on. */}
                      <div style={{ display: 'flex', gap: 4, marginTop: 5, justifyContent: 'flex-end' }}>
                        {REAL_HUBS.map(hub => {
                          const hc = hubColor(hub)
                          const list = byHubDay.get(`${dayStr}::${hub}`) ?? []
                          return (
                            <div key={hub}
                              onDragOver={e => { e.preventDefault(); e.stopPropagation() }}
                              onDrop={e => { e.stopPropagation(); onDropHub(hub, day) }}
                              onClick={e => {
                                e.stopPropagation()
                                if (heldKey) { const u = unitByKey.get(heldKey); if (u) placeOnHub(u, hub, day); return }
                                if (list.length > 0) {
                                  const r = e.currentTarget.getBoundingClientRect()
                                  setHubPopover({ day: dayStr, hub, x: r.left, y: r.bottom })
                                }
                              }}
                              title={list.length > 0 ? `คลิกเพื่อดูตู้ที่ส่งไป Hub ${hub} วันที่ ${day.getDate()} ${MONTHS_TH[day.getMonth()]}` : `ส่งไป Hub ${hub} วันที่ ${day.getDate()} ${MONTHS_TH[day.getMonth()]}`}
                              style={{
                                position: 'relative', width: 18, height: 18, borderRadius: '50%',
                                background: list.length > 0 ? hc.dot : 'transparent', border: `1px solid ${hc.dot}`,
                                display: 'flex', alignItems: 'center', justifyContent: 'center',
                                cursor: heldKey || list.length > 0 ? 'pointer' : 'default', flexShrink: 0,
                                opacity: list.length > 0 ? 1 : 0.5,
                              }}
                            >
                              {list.length > 1 ? (
                                <span className={numFont.className} style={{ fontSize: 9, fontWeight: 700, color: '#fff' }}>{list.length}</span>
                              ) : list.length === 1 ? (
                                <span style={{ fontSize: 10, lineHeight: 1 }}>🏭</span>
                              ) : null}
                            </div>
                          )
                        })}
                      </div>

                    </div>
                  )
                })}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* ── Hub day popover: view + drag out what's inside ── portal straight
          into document.body, positioned from the clicked icon's own bounding
          box, so it can't be clipped by the calendar's scroll wrapper. ── */}
      {hubPopover && typeof document !== 'undefined' && createPortal((() => {
        const popList = byHubDay.get(`${hubPopover.day}::${hubPopover.hub}`) ?? []
        const phc = hubColor(hubPopover.hub)
        return (
          <>
            <div onClick={() => setHubPopover(null)} style={{ position: 'fixed', inset: 0, zIndex: 999 }} />
            <div
              onClick={e => e.stopPropagation()}
              onDragOver={e => { e.preventDefault(); e.stopPropagation() }}
              onDrop={e => { e.stopPropagation(); onDropTray() }}
              className={bodyFont.className}
              style={{
                position: 'fixed', top: hubPopover.y + 4, left: Math.min(hubPopover.x, window.innerWidth - 190), zIndex: 1000,
                background: t.cardBg, color: t.text, borderRadius: 10, padding: 10, minWidth: 170, maxWidth: 220,
                border: `1px solid ${t.border}`, boxShadow: '0 8px 24px rgba(0,0,0,0.35)',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 8 }}>
                <span style={{ width: 8, height: 8, borderRadius: '50%', background: phc.dot, flexShrink: 0 }} />
                <span style={{ fontWeight: 700, fontSize: 12 }}>Hub {hubPopover.hub}</span>
                <button onClick={() => setHubPopover(null)} style={{ marginLeft: 'auto', background: 'none', border: 'none', color: t.textMuted, fontSize: 14, cursor: 'pointer', lineHeight: 1 }}>×</button>
              </div>
              {popList.length === 0 ? (
                <p style={{ fontSize: 11, color: t.textMuted, margin: 0 }}>ไม่มีตู้แล้ว</p>
              ) : (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                  {popList.map(u => (
                    <div key={unitKey(u)} style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
                      <Block u={u} color={invoiceColor(u.invoiceNo)} solid held={heldKey === unitKey(u)} small
                        numFontClass={numFont.className}
                        onDragStart={() => onDragStart(unitKey(u))}
                        onDragEnd={onDragEnd}
                        onClick={() => onBlockActivate(unitKey(u))}
                        onContextMenu={e => { e.preventDefault(); toggleSize(u) }}
                      />
                      <span style={{ fontSize: 8, color: t.textMuted }}>{u.invoiceNo}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </>
        )
      })(), document.body)}

      {/* ── Toasts ── */}
      <div style={{ position: 'fixed', bottom: 16, right: 16, display: 'flex', flexDirection: 'column', gap: 8, zIndex: 50 }}>
        {toasts.map(tt => (
          <div key={tt.id} style={{
            background: tt.warn ? t.overBg : t.cardBg, color: tt.warn ? t.overBorder : t.text,
            border: `1px solid ${tt.warn ? t.overBorder : t.border}`, borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 600, boxShadow: '0 4px 12px rgba(0,0,0,0.15)', maxWidth: 320,
          }}>
            {tt.text}
          </div>
        ))}
      </div>
    </div>
  )
}

function Stat({ label, value, t, danger }: { label: string; value: number; t: typeof LIGHT; danger?: boolean }) {
  return (
    <div style={{ textAlign: 'center' }}>
      <div className={numFont.className} style={{ fontSize: 22, fontWeight: 700, color: danger ? '#dc2626' : t.text, lineHeight: 1 }}>{value}</div>
      <div style={{ fontSize: 10, color: t.textMuted, marginTop: 2, whiteSpace: 'nowrap' }}>{label}</div>
    </div>
  )
}

function Block({ u, color, solid, held, small, warn, numFontClass, onDragStart, onDragEnd, onClick, onContextMenu }: {
  u: ContainerUnit
  color: { bg: string; text: string }
  // Solid colored block = the "real" draggable container. Dashed outline =
  // a ghost placeholder for one that's already placed elsewhere (shown only
  // in the tray, so the card still lists every container it has).
  solid: boolean
  held: boolean
  small?: boolean
  warn?: boolean
  numFontClass: string
  onDragStart: () => void
  onDragEnd: () => void
  onClick: (e: React.MouseEvent) => void
  onContextMenu: (e: React.MouseEvent) => void
}) {
  const w = small ? (u.isLong ? 52 : 30) : (u.isLong ? 86 : 46)
  const h = small ? 20 : 30
  return (
    <div
      draggable
      onDragStart={onDragStart}
      onDragEnd={onDragEnd}
      onClick={onClick}
      onContextMenu={onContextMenu}
      tabIndex={0}
      role="button"
      onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onClick(e as unknown as React.MouseEvent) } }}
      title={`${u.invoiceNo} ตู้ ${u.label}${u.isLong ? ' (40ft)' : ' (20ft)'} — คลิกขวาเพื่อสลับขนาดตู้`}
      className={numFontClass}
      style={{
        width: w, height: h, borderRadius: 4, cursor: 'grab', userSelect: 'none',
        background: solid ? `repeating-linear-gradient(180deg, ${color.bg} 0 3px, ${shade(color.bg)} 3px 6px)` : 'transparent',
        border: solid ? `1px solid ${shade(color.bg)}` : `1.5px dashed ${color.bg}`,
        color: color.text, fontSize: small ? 8 : 10, fontWeight: 700,
        display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', lineHeight: 1,
        boxShadow: held ? `0 0 0 2px #f5c518` : 'none',
        animation: solid ? 'cm-drop 0.18s ease-out' : undefined,
        position: 'relative', overflow: 'hidden', padding: '0 2px',
      }}
    >
      <span style={{ color: solid ? color.text : color.bg }}>{u.label}</span>
      {warn && <span style={{ position: 'absolute', top: -2, right: -2, fontSize: 9 }}>⚠</span>}
      <style jsx>{`
        @keyframes cm-drop {
          0% { transform: translateY(-6px) scale(0.92); opacity: 0.4; }
          100% { transform: translateY(0) scale(1); opacity: 1; }
        }
      `}</style>
    </div>
  )
}

function shade(hex: string) {
  const n = parseInt(hex.slice(1), 16)
  const r = Math.max(0, (n >> 16) - 24), g = Math.max(0, ((n >> 8) & 0xff) - 24), b = Math.max(0, (n & 0xff) - 24)
  return `rgb(${r},${g},${b})`
}

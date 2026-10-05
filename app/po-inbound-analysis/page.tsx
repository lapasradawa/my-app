'use client'

import { useEffect, useMemo, useState, type ReactNode } from 'react'
import Link from 'next/link'
import { supabase } from '@/lib/supabase'
import NavBar from '@/components/NavBar'

// ── Types ────────────────────────────────────────────────────────────────────
interface POUploadRow {
  id: string
  supplier: string
  project: string
  currency: string
  total_amount: number | null
  exchange_rate: number | null
  po_rbs_ch_no: string | null
  po_rbs_th_no: string | null
  po_date: string | null
}

interface ExchangeRateEntry { amount: number; rate: number }

interface InvoiceRow {
  id: string
  invoice_no: string
  supplier: string | null
  estimated_arrival: string | null
  payment_date: string | null
  total_amount: number | null
  currency: string | null
  exchange_rate: number | null
  exchange_rates: ExchangeRateEntry[] | null
}

// ── Period helpers (same convention as /summary, /po-summary) ───────────────
function mKey(d: string | null): string | null {
  if (!d) return null
  const dt = new Date(d + 'T00:00:00')
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`
}
function mLabel(k: string): string {
  const [y, m] = k.split('-')
  const names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
  return `${names[parseInt(m) - 1]} ${y}`
}
function generateMonthKeys(count = 12): string[] {
  const keys: string[] = []
  const d = new Date()
  for (let i = 0; i < count; i++) {
    keys.unshift(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`)
    d.setMonth(d.getMonth() - 1)
  }
  return keys
}

function fmt(n: number, dec = 2) {
  return n.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec })
}
function fmtThb(n: number) {
  return `฿${Math.round(n).toLocaleString()}`
}
function sumByCurrency<T>(items: T[], getCurrency: (t: T) => string | null, getAmount: (t: T) => number | null): Map<string, number> {
  const map = new Map<string, number>()
  for (const item of items) {
    const ccy = getCurrency(item)
    const amt = getAmount(item)
    if (!ccy || amt == null) continue
    map.set(ccy, (map.get(ccy) ?? 0) + amt)
  }
  return map
}
function fmtByCurrency(map: Map<string, number>): string {
  return Array.from(map.entries())
    .filter(([, v]) => v > 0)
    .sort((a, b) => b[1] - a[1])
    .map(([ccy, amt]) => `${ccy} ${fmt(amt, 2)}`)
    .join(' · ')
}
function fmtMillions(n: number) {
  return `${(n / 1e6).toFixed(2)} M`
}
function fmtCompact(n: number) {
  if (n >= 1e6) return `${(n / 1e6).toFixed(1)}M`
  if (n >= 1e3) return `${(n / 1e3).toFixed(0)}K`
  return n.toFixed(0)
}
function fmtDate(d: string | null) {
  if (!d) return '—'
  return new Date(d + 'T00:00:00').toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
}

// Estimated rate: the document's own rate if already set, else the shared
// settings estimate — same fallback /po-summary uses for its "planned" totals.
function estimateRate(currency: string | null, exchangeRate: number | null, cnyRate: number, usdRate: number): number {
  if (exchangeRate != null) return exchangeRate
  if (currency === 'THB') return 1
  return currency === 'USD' ? usdRate : cnyRate
}

function poToThb(u: POUploadRow, cnyRate: number, usdRate: number): number {
  return (u.total_amount ?? 0) * estimateRate(u.currency, u.exchange_rate, cnyRate, usdRate)
}
function invoiceEstimateThb(inv: InvoiceRow, cnyRate: number, usdRate: number): number {
  return (inv.total_amount ?? 0) * estimateRate(inv.currency, inv.exchange_rate, cnyRate, usdRate)
}
// Actual THB — only the real rate (multi-tranche exchange_rates, else a set
// exchange_rate), null when neither is recorded yet. Exact same formula as
// Report page's "Actual FOB THB (Finance)" — intentionally NOT an estimate.
function invoiceActualThb(inv: InvoiceRow): number | null {
  if (inv.exchange_rates && inv.exchange_rates.length > 0) {
    return inv.exchange_rates.reduce((s, e) => s + e.amount * e.rate, 0)
  }
  if (inv.total_amount != null && inv.exchange_rate != null) {
    return inv.total_amount * inv.exchange_rate
  }
  return null
}

const COLOR_PO = '#6b5ea8'
const COLOR_INBOUND = '#2a7c9a'
const COLOR_PAYMENT = '#3d8b82'

export default function POInboundAnalysisPage() {
  const [poUploads, setPoUploads] = useState<POUploadRow[]>([])
  const [invoices, setInvoices] = useState<InvoiceRow[]>([])
  const [cnyRate, setCnyRate] = useState(4.85)
  const [usdRate, setUsdRate] = useState(33.0)
  const [loading, setLoading] = useState(true)
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(() => {
    const now = new Date()
    return new Set([`${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`])
  })
  const [periodOpen, setPeriodOpen] = useState(false)

  useEffect(() => {
    async function load() {
      const [{ data: uploads }, { data: invs }, { data: settings }] = await Promise.all([
        supabase.from('po_uploads').select('id, supplier, project, currency, total_amount, exchange_rate, po_rbs_ch_no, po_rbs_th_no, po_date'),
        supabase.from('invoices').select('id, invoice_no, supplier, estimated_arrival, payment_date, total_amount, currency, exchange_rate, exchange_rates'),
        supabase.from('cost_settings').select('key, value'),
      ])
      setPoUploads((uploads ?? []) as POUploadRow[])
      setInvoices((invs ?? []) as InvoiceRow[])
      if (settings) {
        const m = Object.fromEntries((settings as { key: string; value: string }[]).map(r => [r.key, r.value]))
        if (m.cny_rate) setCnyRate(parseFloat(m.cny_rate))
        if (m.usd_rate) setUsdRate(parseFloat(m.usd_rate))
      }
      setLoading(false)
    }
    load()
  }, [])

  const months = useMemo(() => generateMonthKeys(18), [])
  const toggleMonth = (k: string) => {
    setSelectedMonths(prev => { const n = new Set(prev); n.has(k) ? n.delete(k) : n.add(k); return n })
  }

  // The chart's own window — fixed starting at CHART_ANCHOR (the month real
  // data actually begins, Apr 2026) instead of always trailing 12 months
  // behind today, so it doesn't open mostly empty. It only starts rolling
  // forward like a normal trailing window once enough time has passed that
  // "today minus 11 months" would land AFTER the anchor on its own — it
  // follows the Period picker's selection, which is independent of this.
  const CHART_ANCHOR = '2026-04'
  const chartMonths = useMemo(() => {
    const now = new Date()
    const d = new Date(now.getFullYear(), now.getMonth() - 11, 1)
    const rollingStart = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
    const start = rollingStart > CHART_ANCHOR ? rollingStart : CHART_ANCHOR
    const sd = new Date(start + '-01T00:00:00')
    const keys: string[] = []
    for (let i = 0; i < 12; i++) {
      keys.push(`${sd.getFullYear()}-${String(sd.getMonth() + 1).padStart(2, '0')}`)
      sd.setMonth(sd.getMonth() + 1)
    }
    return keys
  }, [])

  // Per-month aggregates, computed once for the union of every month key
  // either the Period picker (months) or the chart (chartMonths) might need.
  const allMonthly = useMemo(() => {
    const keys = new Set([...months, ...chartMonths])
    const map = new Map<string, {
      month: string
      poList: POUploadRow[]
      inboundList: InvoiceRow[]
      paymentList: InvoiceRow[]
      poTotal: number
      inboundTotal: number
      paymentTotal: number
    }>()
    for (const month of keys) {
      const poList = poUploads.filter(u => mKey(u.po_date) === month)
      const inboundList = invoices.filter(i => mKey(i.estimated_arrival) === month)
      const paymentList = invoices.filter(i => mKey(i.payment_date) === month)
      map.set(month, {
        month,
        poList,
        inboundList,
        paymentList,
        poTotal: poList.reduce((s, u) => s + poToThb(u, cnyRate, usdRate), 0),
        inboundTotal: inboundList.reduce((s, i) => s + invoiceEstimateThb(i, cnyRate, usdRate), 0),
        paymentTotal: paymentList.reduce((s, i) => s + (invoiceActualThb(i) ?? 0), 0),
      })
    }
    return map
  }, [months, chartMonths, poUploads, invoices, cnyRate, usdRate])

  const monthly = useMemo(() => months.map(k => allMonthly.get(k)!), [months, allMonthly])
  const chartMonthly = useMemo(() => chartMonths.map(k => allMonthly.get(k)!), [chartMonths, allMonthly])

  // Empty selection (like Invoice Summary) means "all months" — everything
  // below (KPI cards + detail tables) is filtered/summed over this set.
  const inPeriod = useMemo(
    () => selectedMonths.size === 0 ? monthly : monthly.filter(m => selectedMonths.has(m.month)),
    [monthly, selectedMonths]
  )

  const selectedPoList = useMemo(() => inPeriod.flatMap(m => m.poList), [inPeriod])
  const selectedInboundList = useMemo(() => inPeriod.flatMap(m => m.inboundList), [inPeriod])
  const selectedPaymentList = useMemo(() => inPeriod.flatMap(m => m.paymentList), [inPeriod])

  // Original-currency totals (e.g. "CNY 123,456.00 · USD 7,890.00") shown
  // alongside each table's THB total, since FOB is always in native currency.
  const poCurrencyBreakdown = useMemo(
    () => fmtByCurrency(sumByCurrency(selectedPoList, u => u.currency, u => u.total_amount)),
    [selectedPoList]
  )
  const inboundCurrencyBreakdown = useMemo(
    () => fmtByCurrency(sumByCurrency(selectedInboundList, i => i.currency, i => i.total_amount)),
    [selectedInboundList]
  )
  const paymentCurrencyBreakdown = useMemo(
    () => fmtByCurrency(sumByCurrency(selectedPaymentList, i => i.currency, i => i.total_amount)),
    [selectedPaymentList]
  )

  const kpi = useMemo(() => ({
    po: inPeriod.reduce((s, m) => s + m.poTotal, 0),
    inbound: inPeriod.reduce((s, m) => s + m.inboundTotal, 0),
    payment: inPeriod.reduce((s, m) => s + m.paymentTotal, 0),
  }), [inPeriod])

  const periodLabel = selectedMonths.size === 0
    ? 'All months'
    : selectedMonths.size === months.length
    ? 'All months'
    : selectedMonths.size === 1
    ? mLabel([...selectedMonths][0])
    : `${selectedMonths.size} months selected`

  return (
    <div style={{ background: '#ede5d4', minHeight: '100vh', fontFamily: 'system-ui,-apple-system,sans-serif' }}>
      <NavBar />

      {/* ── Hero header ── */}
      <div style={{ background: '#1e3340', padding: '24px 32px' }}>
        <div style={{ maxWidth: 1400, margin: '0 auto', display: 'flex', alignItems: 'center', gap: 24, flexWrap: 'wrap' }}>
          <div style={{ flex: 1, minWidth: 200 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: '#3d8b82', textTransform: 'uppercase', letterSpacing: '0.12em', marginBottom: 4 }}>Import PO</div>
            <div style={{ fontSize: 30, fontWeight: 900, color: '#d4962a', lineHeight: 1.1 }}>PO & INBOUND ANALYSIS</div>
            <div style={{ fontSize: 13, color: '#7a9aaa', marginTop: 4 }}>เปรียบเทียบ PO ที่เปิด / สินค้าเข้าคลัง (ประมาณการณ์) / ยอดจ่ายจริง รายเดือน</div>
          </div>
          {[
            { icon: '📝', value: fmtThb(kpi.po), label: 'PO เปิด', bg: COLOR_PO },
            { icon: '🚢', value: fmtThb(kpi.inbound), label: 'Inbound', bg: COLOR_INBOUND },
            { icon: '💰', value: fmtThb(kpi.payment), label: 'Payment', bg: COLOR_PAYMENT },
          ].map(card => (
            <div key={card.label} style={{ background: card.bg, borderRadius: 14, padding: '14px 20px', minWidth: 160, display: 'flex', alignItems: 'center', gap: 12 }}>
              <span style={{ fontSize: 26 }}>{card.icon}</span>
              <div>
                <div style={{ fontSize: 20, fontWeight: 900, color: '#fff', lineHeight: 1.1 }}>{card.value}</div>
                <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.75)', fontWeight: 600, marginTop: 2 }}>{card.label}</div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* ── Period dropdown bar ── */}
      <div style={{ background: '#18303c', padding: '10px 32px', position: 'relative', zIndex: 30 }}>
        <div style={{ maxWidth: 1400, margin: '0 auto', display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 10, fontWeight: 800, color: '#5a8a9a', textTransform: 'uppercase', letterSpacing: '0.12em' }}>Period:</span>

          <div style={{ position: 'relative' }}>
            <button onClick={() => setPeriodOpen(o => !o)} style={{
              display: 'flex', alignItems: 'center', gap: 8,
              padding: '6px 14px', borderRadius: 8, cursor: 'pointer',
              background: '#1e3a4a', border: '1px solid #2e5060',
              color: '#d4c8a8', fontSize: 11, fontWeight: 700,
              minWidth: 200,
            }}>
              <span style={{ flex: 1, textAlign: 'left' }}>
                {selectedMonths.size === 0 ? 'Select period…' : periodLabel}
              </span>
              <span style={{ fontSize: 9, color: '#5a8a9a' }}>{periodOpen ? '▲' : '▼'}</span>
            </button>

            {periodOpen && (
              <div style={{
                position: 'absolute', top: 'calc(100% + 6px)', left: 0,
                background: '#1a2e3c', border: '1px solid #2e5060', borderRadius: 12,
                boxShadow: '0 8px 32px rgba(0,0,0,0.4)', padding: 12, minWidth: 280, zIndex: 100,
              }}>
                <div style={{ display: 'flex', gap: 8, marginBottom: 10, paddingBottom: 8, borderBottom: '1px solid #2a4455' }}>
                  <button onClick={() => setSelectedMonths(new Set(months))}
                    style={{ flex: 1, padding: '4px 0', borderRadius: 6, border: 'none', background: '#3d8b82', color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    All
                  </button>
                  <button onClick={() => setSelectedMonths(new Set())}
                    style={{ flex: 1, padding: '4px 0', borderRadius: 6, border: 'none', background: '#2a4455', color: '#8a9aaa', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    Clear
                  </button>
                  <button onClick={() => setPeriodOpen(false)}
                    style={{ padding: '4px 10px', borderRadius: 6, border: 'none', background: '#d4962a', color: '#fff', fontSize: 10, fontWeight: 700, cursor: 'pointer' }}>
                    Done
                  </button>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 5 }}>
                  {months.map(k => (
                    <button key={k} onClick={() => toggleMonth(k)} style={{
                      padding: '5px 4px', borderRadius: 6, cursor: 'pointer', fontSize: 10, fontWeight: 700,
                      background: selectedMonths.has(k) ? '#d4962a' : 'transparent',
                      color: selectedMonths.has(k) ? '#1a2d3a' : '#8a9aaa',
                      border: selectedMonths.has(k) ? '1px solid #d4962a' : '1px solid #2a4455',
                      transition: 'all 0.12s',
                    }}>{mLabel(k)}</button>
                  ))}
                </div>
              </div>
            )}
          </div>

          {selectedMonths.size > 0 && selectedMonths.size < months.length && [...selectedMonths].sort().map(k => (
            <div key={k} style={{ display: 'flex', alignItems: 'center', gap: 4, background: '#d4962a', borderRadius: 6, padding: '3px 8px' }}>
              <span style={{ fontSize: 10, fontWeight: 700, color: '#1a2d3a' }}>{mLabel(k)}</span>
              <button onClick={() => toggleMonth(k)} style={{ background: 'none', border: 'none', color: '#1a2d3a', fontSize: 11, cursor: 'pointer', padding: 0, lineHeight: 1, opacity: 0.7 }}>×</button>
            </div>
          ))}
        </div>
      </div>

      <div style={{ maxWidth: 1400, margin: '0 auto', padding: '24px 32px' }}>
        {loading ? (
          <div style={{ textAlign: 'center', padding: '60px 0', color: '#bbb', fontSize: 13 }}>กำลังโหลด...</div>
        ) : (
          <>
            {/* ── Chart ── */}
            <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, padding: 20, marginBottom: 20 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginBottom: 10 }}>
                <span style={{ fontSize: 12, fontWeight: 800, color: '#3a2a1a' }}>PO เปิด vs Inbound vs Payment รายเดือน</span>
                <Legend color={COLOR_PO} label="PO เปิด (ประมาณการณ์)" />
                <Legend color={COLOR_INBOUND} label="Inbound (ประมาณการณ์)" />
                <Legend color={COLOR_PAYMENT} label="จ่ายจริง" />
                <span style={{ fontSize: 10, color: '#bbb', marginLeft: 'auto' }}>คลิกเดือนเพื่อเลือก/ยกเลิกช่วงเวลา</span>
              </div>
              <GroupedBarChart
                months={chartMonths}
                poValues={chartMonthly.map(m => m.poTotal)}
                inboundValues={chartMonthly.map(m => m.inboundTotal)}
                paymentValues={chartMonthly.map(m => m.paymentTotal)}
                selectedMonths={selectedMonths}
                onSelect={toggleMonth}
              />
            </div>

            {/* ── Detail panel ── */}
            <div style={{ marginBottom: 12 }}>
              <span style={{ fontSize: 15, fontWeight: 900, color: '#3a2a1a' }}>{periodLabel}</span>
              <span style={{ fontSize: 11, color: '#9a8a7a', marginLeft: 8 }}>รายละเอียด PO / Inbound / Payment ของช่วงเวลาที่เลือก</span>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 16 }}>
              {/* PO table */}
              <DetailTable
                title="PO ที่เปิดเดือนนี้"
                color={COLOR_PO}
                total={kpi.po}
                currencyBreakdown={poCurrencyBreakdown}
                emptyLabel="ไม่มี PO ที่เปิดในเดือนนี้"
                headers={['PO No.', 'Supplier', 'Project', 'วันที่เปิด PO', 'FOB (Original Currency)']}
              >
                {selectedPoList.map(u => (
                  <tr key={u.id} style={{ borderBottom: '1px solid #f5efe8' }}>
                    <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                      <Link href={`/po-builder/${u.id}`} style={{ color: COLOR_PO, textDecoration: 'none', fontWeight: 700 }}>
                        {u.po_rbs_ch_no || u.po_rbs_th_no || '—'}
                      </Link>
                    </td>
                    <td style={{ padding: '7px 10px', color: '#5a4a3a', whiteSpace: 'nowrap' }}>{u.supplier}</td>
                    <td style={{ padding: '7px 10px', color: '#6a5a4a', whiteSpace: 'nowrap' }}>{u.project}</td>
                    <td style={{ padding: '7px 10px', color: '#6a5a4a', whiteSpace: 'nowrap' }}>{fmtDate(u.po_date)}</td>
                    <td style={{ padding: '7px 10px', textAlign: 'right', fontWeight: 700, color: '#2a2a1a', whiteSpace: 'nowrap' }}>
                      {u.total_amount != null ? `${fmt(u.total_amount)} ${u.currency}` : '—'}
                    </td>
                  </tr>
                ))}
              </DetailTable>

              {/* Inbound table */}
              <DetailTable
                title="Invoice ที่ประมาณการณ์เข้าคลังเดือนนี้"
                color={COLOR_INBOUND}
                total={kpi.inbound}
                currencyBreakdown={inboundCurrencyBreakdown}
                emptyLabel="ไม่มี Invoice ที่ประมาณการณ์เข้าคลังในเดือนนี้"
                headers={['Invoice No.', 'Supplier', 'ประมาณการณ์เข้าคลัง', 'FOB (Original Currency)']}
              >
                {selectedInboundList.map(i => (
                  <tr key={i.id} style={{ borderBottom: '1px solid #f5efe8' }}>
                    <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                      <Link href={`/dashboard/${i.id}`} style={{ color: COLOR_INBOUND, textDecoration: 'none', fontWeight: 700 }}>{i.invoice_no}</Link>
                    </td>
                    <td style={{ padding: '7px 10px', color: '#5a4a3a', whiteSpace: 'nowrap' }}>{i.supplier || '—'}</td>
                    <td style={{ padding: '7px 10px', color: '#6a5a4a', whiteSpace: 'nowrap' }}>{fmtDate(i.estimated_arrival)}</td>
                    <td style={{ padding: '7px 10px', textAlign: 'right', fontWeight: 700, color: '#2a2a1a', whiteSpace: 'nowrap' }}>
                      {i.total_amount != null ? `${fmt(i.total_amount)} ${i.currency}` : '—'}
                    </td>
                  </tr>
                ))}
              </DetailTable>

              {/* Payment table */}
              <DetailTable
                title="ยอดจ่ายจริงเดือนนี้"
                color={COLOR_PAYMENT}
                total={kpi.payment}
                totalLabel="Actual THB"
                currencyBreakdown={paymentCurrencyBreakdown}
                emptyLabel="ไม่มีรายการจ่ายเงินในเดือนนี้"
                headers={['Invoice No.', 'Supplier', 'วันที่จ่าย', 'ยอดจ่ายจริง (THB)']}
              >
                {selectedPaymentList.map(i => {
                  const actual = invoiceActualThb(i)
                  return (
                    <tr key={i.id} style={{ borderBottom: '1px solid #f5efe8' }}>
                      <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>
                        <Link href={`/dashboard/${i.id}`} style={{ color: COLOR_PAYMENT, textDecoration: 'none', fontWeight: 700 }}>{i.invoice_no}</Link>
                      </td>
                      <td style={{ padding: '7px 10px', color: '#5a4a3a', whiteSpace: 'nowrap' }}>{i.supplier || '—'}</td>
                      <td style={{ padding: '7px 10px', color: '#6a5a4a', whiteSpace: 'nowrap' }}>{fmtDate(i.payment_date)}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right', fontWeight: 700, color: actual != null ? '#2a2a1a' : '#c85a3a', whiteSpace: 'nowrap' }}>
                        {actual != null ? fmtThb(actual) : 'ไม่มี exchange rate จริง'}
                      </td>
                    </tr>
                  )
                })}
              </DetailTable>
            </div>
          </>
        )}
      </div>
    </div>
  )
}

function Legend({ color, label }: { color: string; label: string }) {
  return (
    <span style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 10, color: '#8a7a6a', fontWeight: 600 }}>
      <span style={{ width: 9, height: 9, borderRadius: 3, background: color, display: 'inline-block' }} />
      {label}
    </span>
  )
}

function DetailTable({ title, color, total, totalLabel = 'est. THB', currencyBreakdown, emptyLabel, headers, children }: {
  title: string
  color: string
  total: number
  totalLabel?: string
  currencyBreakdown?: string
  emptyLabel: string
  headers: string[]
  children: ReactNode
}) {
  const hasRows = Array.isArray(children) ? children.length > 0 : !!children
  return (
    <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, overflow: 'hidden' }}>
      <div style={{ padding: '10px 16px', borderBottom: '1px solid #e2d8c8', background: '#f5efe4', display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
        <span style={{ width: 8, height: 8, borderRadius: 2, background: color, display: 'inline-block' }} />
        <span style={{ fontSize: 12, fontWeight: 800, color: '#3a2a1a' }}>{title}</span>
        <span style={{ fontSize: 12, fontWeight: 900, color, marginLeft: 'auto' }}>{fmtThb(total)} ({totalLabel})</span>
        {currencyBreakdown && <span style={{ fontSize: 10, color: '#9a8a7a', fontWeight: 600 }}>{currencyBreakdown}</span>}
      </div>
      {!hasRows ? (
        <div style={{ textAlign: 'center', padding: '20px 0', color: '#bbb', fontSize: 12 }}>{emptyLabel}</div>
      ) : (
        <div style={{ overflowX: 'auto', maxHeight: 280, overflowY: 'auto' }}>
          <table style={{ fontSize: 11, width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr style={{ background: '#f0ebe0', position: 'sticky', top: 0, zIndex: 1 }}>
                {headers.map((h, i) => (
                  <th key={h} style={{ padding: '7px 10px', textAlign: i === headers.length - 1 ? 'right' : 'left', fontSize: 10, fontWeight: 800, color: '#8a7a6a', whiteSpace: 'nowrap', borderBottom: '1px solid #e2d8c8' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>{children}</tbody>
          </table>
        </div>
      )}
    </div>
  )
}

function GroupedBarChart({ months, poValues, inboundValues, paymentValues, selectedMonths, onSelect }: {
  months: string[]
  poValues: number[]
  inboundValues: number[]
  paymentValues: number[]
  selectedMonths: Set<string>
  onSelect: (month: string) => void
}) {
  // Extra top padding (padT) leaves room for each bar's rotated value label.
  const W = 1100, H = 230, padL = 46, padR = 10, padT = 36, padB = 28
  const cW = W - padL - padR, cH = H - padT - padB
  const maxVal = Math.max(...poValues, ...inboundValues, ...paymentValues, 1)
  const n = Math.max(months.length, 1)
  const groupW = cW / n
  const barW = groupW / 4.2
  const yPos = (v: number) => padT + cH - (v / maxVal) * cH
  const barH = (v: number) => (v / maxVal) * cH

  function ValueLabel({ x, v }: { x: number; v: number }) {
    if (v <= 0) return null
    const y = yPos(v) - 4
    return (
      <text x={x} y={y} textAnchor="start" fontSize={7} fill="#1a1a1a" fontWeight={700}
        transform={`rotate(-90 ${x} ${y})`}>
        {fmtMillions(v)}
      </text>
    )
  }

  return (
    <svg viewBox={`0 0 ${W} ${H}`} className="w-full" style={{ maxHeight: 260 }}>
      {[0, 0.25, 0.5, 0.75, 1].map(frac => {
        const y = padT + cH - frac * cH
        return (
          <g key={frac}>
            <line x1={padL} x2={W - padR} y1={y} y2={y} stroke="#e2d8c8" strokeWidth={1} />
            <text x={padL - 6} y={y + 3} textAnchor="end" fontSize={9} fill="#1a1a1a">{fmtCompact(maxVal * frac)}</text>
          </g>
        )
      })}
      {months.map((m, i) => {
        const gx = padL + i * groupW
        const isSel = selectedMonths.size > 0 && selectedMonths.size < months.length && selectedMonths.has(m)
        return (
          <g key={m} onClick={() => onSelect(m)} style={{ cursor: 'pointer' }}>
            {isSel && <rect x={gx} y={padT} width={groupW} height={cH} fill="#d4962a" opacity={0.1} />}
            <rect x={gx + barW * 0.3} y={yPos(poValues[i])} width={barW} height={barH(poValues[i])} fill={COLOR_PO} rx={2} />
            <rect x={gx + barW * 1.5} y={yPos(inboundValues[i])} width={barW} height={barH(inboundValues[i])} fill={COLOR_INBOUND} rx={2} />
            <rect x={gx + barW * 2.7} y={yPos(paymentValues[i])} width={barW} height={barH(paymentValues[i])} fill={COLOR_PAYMENT} rx={2} />
            <ValueLabel x={gx + barW * 0.3 + barW / 2} v={poValues[i]} />
            <ValueLabel x={gx + barW * 1.5 + barW / 2} v={inboundValues[i]} />
            <ValueLabel x={gx + barW * 2.7 + barW / 2} v={paymentValues[i]} />
            <text x={gx + groupW / 2} y={H - 8} textAnchor="middle" fontSize={9} fill="#1a1a1a" fontWeight={isSel ? 800 : 500}>
              {mLabel(m).slice(0, 3)}
            </text>
          </g>
        )
      })}
    </svg>
  )
}

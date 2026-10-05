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
  const [selectedMonth, setSelectedMonth] = useState<string>(() => generateMonthKeys(1)[0])

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

  const months = useMemo(() => generateMonthKeys(12), [])

  const monthly = useMemo(() => months.map(month => {
    const poList = poUploads.filter(u => mKey(u.po_date) === month)
    const inboundList = invoices.filter(i => mKey(i.estimated_arrival) === month)
    const paymentList = invoices.filter(i => mKey(i.payment_date) === month)
    return {
      month,
      poList,
      inboundList,
      paymentList,
      poTotal: poList.reduce((s, u) => s + poToThb(u, cnyRate, usdRate), 0),
      inboundTotal: inboundList.reduce((s, i) => s + invoiceEstimateThb(i, cnyRate, usdRate), 0),
      paymentTotal: paymentList.reduce((s, i) => s + (invoiceActualThb(i) ?? 0), 0),
    }
  }), [months, poUploads, invoices, cnyRate, usdRate])

  const selected = monthly.find(m => m.month === selectedMonth) ?? monthly[monthly.length - 1]

  const kpi = useMemo(() => ({
    po: monthly.reduce((s, m) => s + m.poTotal, 0),
    inbound: monthly.reduce((s, m) => s + m.inboundTotal, 0),
    payment: monthly.reduce((s, m) => s + m.paymentTotal, 0),
  }), [monthly])

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
            { icon: '📝', value: fmtThb(kpi.po), label: 'PO เปิด (12 เดือน, ประมาณการณ์)', bg: COLOR_PO },
            { icon: '🚢', value: fmtThb(kpi.inbound), label: 'Inbound (12 เดือน, ประมาณการณ์)', bg: COLOR_INBOUND },
            { icon: '💰', value: fmtThb(kpi.payment), label: 'จ่ายจริง (12 เดือน)', bg: COLOR_PAYMENT },
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
                <span style={{ fontSize: 10, color: '#bbb', marginLeft: 'auto' }}>คลิกเดือนเพื่อดูรายละเอียด</span>
              </div>
              <GroupedBarChart
                months={months}
                poValues={monthly.map(m => m.poTotal)}
                inboundValues={monthly.map(m => m.inboundTotal)}
                paymentValues={monthly.map(m => m.paymentTotal)}
                selectedMonth={selectedMonth}
                onSelect={setSelectedMonth}
              />
            </div>

            {/* ── Detail panel ── */}
            <div style={{ marginBottom: 12 }}>
              <span style={{ fontSize: 15, fontWeight: 900, color: '#3a2a1a' }}>{mLabel(selected.month)}</span>
              <span style={{ fontSize: 11, color: '#9a8a7a', marginLeft: 8 }}>รายละเอียด PO / Inbound / Payment ของเดือนนี้</span>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 16 }}>
              {/* PO table */}
              <DetailTable
                title="PO ที่เปิดเดือนนี้"
                color={COLOR_PO}
                total={selected.poTotal}
                emptyLabel="ไม่มี PO ที่เปิดในเดือนนี้"
                headers={['PO No.', 'Supplier', 'Project', 'วันที่เปิด PO', 'FOB (Original Currency)']}
              >
                {selected.poList.map(u => (
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
                total={selected.inboundTotal}
                emptyLabel="ไม่มี Invoice ที่ประมาณการณ์เข้าคลังในเดือนนี้"
                headers={['Invoice No.', 'Supplier', 'ประมาณการณ์เข้าคลัง', 'FOB (Original Currency)']}
              >
                {selected.inboundList.map(i => (
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
                total={selected.paymentTotal}
                emptyLabel="ไม่มีรายการจ่ายเงินในเดือนนี้"
                headers={['Invoice No.', 'Supplier', 'วันที่จ่าย', 'ยอดจ่ายจริง (THB)']}
              >
                {selected.paymentList.map(i => {
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

function DetailTable({ title, color, total, emptyLabel, headers, children }: {
  title: string
  color: string
  total: number
  emptyLabel: string
  headers: string[]
  children: ReactNode
}) {
  const hasRows = Array.isArray(children) ? children.length > 0 : !!children
  return (
    <div style={{ background: '#faf5ee', border: '1px solid #e2d8c8', borderRadius: 16, overflow: 'hidden' }}>
      <div style={{ padding: '10px 16px', borderBottom: '1px solid #e2d8c8', background: '#f5efe4', display: 'flex', alignItems: 'center', gap: 8 }}>
        <span style={{ width: 8, height: 8, borderRadius: 2, background: color, display: 'inline-block' }} />
        <span style={{ fontSize: 12, fontWeight: 800, color: '#3a2a1a' }}>{title}</span>
        <span style={{ fontSize: 12, fontWeight: 900, color, marginLeft: 'auto' }}>{fmtThb(total)} (est. THB)</span>
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

function GroupedBarChart({ months, poValues, inboundValues, paymentValues, selectedMonth, onSelect }: {
  months: string[]
  poValues: number[]
  inboundValues: number[]
  paymentValues: number[]
  selectedMonth: string
  onSelect: (month: string) => void
}) {
  const W = 1100, H = 230, padL = 46, padR = 10, padT = 10, padB = 28
  const cW = W - padL - padR, cH = H - padT - padB
  const maxVal = Math.max(...poValues, ...inboundValues, ...paymentValues, 1)
  const n = Math.max(months.length, 1)
  const groupW = cW / n
  const barW = groupW / 4.2
  const yPos = (v: number) => padT + cH - (v / maxVal) * cH
  const barH = (v: number) => (v / maxVal) * cH

  return (
    <svg viewBox={`0 0 ${W} ${H}`} className="w-full" style={{ maxHeight: 260 }}>
      {[0, 0.25, 0.5, 0.75, 1].map(frac => {
        const y = padT + cH - frac * cH
        return (
          <g key={frac}>
            <line x1={padL} x2={W - padR} y1={y} y2={y} stroke="#e2d8c8" strokeWidth={1} />
            <text x={padL - 6} y={y + 3} textAnchor="end" fontSize={9} fill="#bbb">{fmtCompact(maxVal * frac)}</text>
          </g>
        )
      })}
      {months.map((m, i) => {
        const gx = padL + i * groupW
        const isSel = m === selectedMonth
        return (
          <g key={m} onClick={() => onSelect(m)} style={{ cursor: 'pointer' }}>
            {isSel && <rect x={gx} y={padT} width={groupW} height={cH} fill="#d4962a" opacity={0.1} />}
            <rect x={gx + barW * 0.3} y={yPos(poValues[i])} width={barW} height={barH(poValues[i])} fill={COLOR_PO} rx={2} />
            <rect x={gx + barW * 1.5} y={yPos(inboundValues[i])} width={barW} height={barH(inboundValues[i])} fill={COLOR_INBOUND} rx={2} />
            <rect x={gx + barW * 2.7} y={yPos(paymentValues[i])} width={barW} height={barH(paymentValues[i])} fill={COLOR_PAYMENT} rx={2} />
            <text x={gx + groupW / 2} y={H - 8} textAnchor="middle" fontSize={9} fill={isSel ? '#d4962a' : '#bbb'} fontWeight={isSel ? 800 : 400}>
              {mLabel(m).slice(0, 3)}
            </text>
          </g>
        )
      })}
    </svg>
  )
}

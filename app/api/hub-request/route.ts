import { NextRequest, NextResponse } from 'next/server'
import { createClient } from '@supabase/supabase-js'

// Lazily constructed so a missing env var fails loudly on the first real
// request (with a clear message in the logs) instead of at build time, and
// instead of silently downgrading to the low-privilege anon key.
function getSupabase() {
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY
  if (!key) {
    throw new Error('SUPABASE_SERVICE_ROLE_KEY is not set — required for /api/hub-request to write with elevated privileges')
  }
  return createClient(process.env.NEXT_PUBLIC_SUPABASE_URL!, key)
}

const HUBS = ['มัยลาภ', 'ขอนแก่น', 'พิษณุโลก', 'สุราษฎร์ธานี'] as const
const APP_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://import-project-zeta.vercel.app'

// Sends a LINE push and — since fetch() only rejects on network failure, never
// on a non-2xx response — records both response errors and thrown ones to
// line_notification_failures so a bad token or an exhausted monthly quota
// shows up in the admin UI instead of failing completely silently.
async function sendLineNotification(
  supabase: ReturnType<typeof getSupabase>,
  lineToken: string,
  lineGroupId: string,
  text: string,
  ctx: { context: string; invoice_id: string; invoice_no?: string | null; container_name: string },
) {
  try {
    const res = await fetch('https://api.line.me/v2/bot/message/push', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${lineToken}` },
      body: JSON.stringify({ to: lineGroupId, messages: [{ type: 'text', text }] }),
    })
    if (!res.ok) {
      const responseBody = await res.text().catch(() => '')
      console.error('LINE notification rejected', res.status, responseBody)
      await supabase.from('line_notification_failures').insert({ ...ctx, status_code: res.status, response_body: responseBody.slice(0, 2000) })
    }
  } catch (e) {
    console.error('LINE notification failed', e)
    await supabase.from('line_notification_failures').insert({ ...ctx, error_message: e instanceof Error ? e.message : String(e) })
  }
}

export async function POST(req: NextRequest) {
  let supabase: ReturnType<typeof getSupabase>
  try {
    supabase = getSupabase()
  } catch (e) {
    console.error(e)
    return NextResponse.json({ error: 'server misconfigured: SUPABASE_SERVICE_ROLE_KEY missing' }, { status: 500 })
  }

  const body = await req.json()
  const { invoice_id, invoice_no, container_name, hub, requested_by } = body

  if (!invoice_id || !container_name || !hub || !requested_by) {
    return NextResponse.json({ error: 'missing fields' }, { status: 400 })
  }
  if (!HUBS.includes(hub)) {
    return NextResponse.json({ error: 'invalid hub' }, { status: 400 })
  }

  // Upsert the request (one per container)
  const { error } = await supabase
    .from('container_hub_requests')
    .upsert({ invoice_id, invoice_no, container_name, hub, requested_by, status: 'pending', confirmed_at: null }, { onConflict: 'invoice_id,container_name' })

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })

  // Get all admin emails
  const { data: admins } = await supabase
    .from('page_permissions')
    .select('email')
    .eq('is_admin', true)

  const adminEmails = (admins ?? []).map((a: { email: string }) => a.email).filter(Boolean)

  // Send LINE notification if configured
  const lineToken = process.env.LINE_CHANNEL_ACCESS_TOKEN
  const lineGroupId = process.env.LINE_GROUP_ID
  if (lineToken && lineGroupId) {
    const invoiceLink = `${APP_URL}/dashboard/${invoice_id}`
    const message = [
      '🔔 คำขอเปลี่ยน Hub',
      '',
      `ตู้: ${container_name}`,
      `Invoice: ${invoice_no}`,
      `Hub ที่ขอ: ${hub}`,
      `ขอโดย: ${requested_by}`,
      '',
      `📄 ดู Invoice: ${invoiceLink}`,
    ].join('\n')
    await sendLineNotification(supabase, lineToken, lineGroupId, message, {
      context: 'hub_request_created', invoice_id, invoice_no, container_name,
    })
  }

  return NextResponse.json({ ok: true })
}

export async function PATCH(req: NextRequest) {
  let supabase: ReturnType<typeof getSupabase>
  try {
    supabase = getSupabase()
  } catch (e) {
    console.error(e)
    return NextResponse.json({ error: 'server misconfigured: SUPABASE_SERVICE_ROLE_KEY missing' }, { status: 500 })
  }

  const body = await req.json()
  const { invoice_id, container_name, action } = body

  if (!invoice_id || !container_name) {
    return NextResponse.json({ error: 'missing fields' }, { status: 400 })
  }

  if (action === 'reject') {
    const { invoice_no, hub, requested_by } = body

    // Only delete a request that is still pending — if another admin already
    // confirmed it, this guard stops the delete from silently reverting an
    // already-confirmed hub assignment back to the default.
    const { data, error } = await supabase
      .from('container_hub_requests')
      .delete()
      .eq('invoice_id', invoice_id)
      .eq('container_name', container_name)
      .eq('status', 'pending')
      .select('id')
    if (error) return NextResponse.json({ error: error.message }, { status: 500 })
    if (!data || data.length === 0) {
      return NextResponse.json({ error: 'already_handled', message: 'รายการนี้ถูกดำเนินการไปแล้วโดยแอดมินคนอื่น' }, { status: 409 })
    }

    // LINE notification after a successful delete
    const lineToken = process.env.LINE_CHANNEL_ACCESS_TOKEN
    const lineGroupId = process.env.LINE_GROUP_ID
    if (lineToken && lineGroupId) {
      await sendLineNotification(
        supabase, lineToken, lineGroupId,
        `❌ ปฏิเสธคำขอเปลี่ยน Hub\n\nตู้: ${container_name}\nInvoice: ${invoice_no}\nHub ที่ขอ: ${hub}\nขอโดย: ${requested_by}`,
        { context: 'hub_request_rejected', invoice_id, invoice_no, container_name },
      )
    }

    return NextResponse.json({ ok: true })
  }

  if (action === 'update_date') {
    const { hub_arrival_date } = body
    if (!hub_arrival_date) {
      return NextResponse.json({ error: 'missing hub_arrival_date' }, { status: 400 })
    }
    // Only updates a request that's already confirmed — editing the ETA of a
    // still-pending request should go through the normal confirm flow instead.
    const { data, error } = await supabase
      .from('container_hub_requests')
      .update({ hub_arrival_date })
      .eq('invoice_id', invoice_id)
      .eq('container_name', container_name)
      .eq('status', 'confirmed')
      .select('id')
    if (error) return NextResponse.json({ error: error.message }, { status: 500 })
    if (!data || data.length === 0) {
      return NextResponse.json({ error: 'not_found', message: 'ไม่พบคำร้องที่ยืนยันแล้วสำหรับตู้นี้' }, { status: 404 })
    }
    return NextResponse.json({ ok: true })
  }

  // Default: confirm — same guard so a stale "confirm" click after another
  // admin already rejected/confirmed the same row doesn't silently no-op or
  // overwrite it.
  const { confirmed_by, hub_arrival_date } = body
  const { data, error } = await supabase
    .from('container_hub_requests')
    .update({ status: 'confirmed', confirmed_at: new Date().toISOString(), confirmed_by: confirmed_by ?? null, hub_arrival_date: hub_arrival_date ?? null })
    .eq('invoice_id', invoice_id)
    .eq('container_name', container_name)
    .eq('status', 'pending')
    .select('id')

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  if (!data || data.length === 0) {
    return NextResponse.json({ error: 'already_handled', message: 'รายการนี้ถูกดำเนินการไปแล้วโดยแอดมินคนอื่น' }, { status: 409 })
  }
  return NextResponse.json({ ok: true })
}

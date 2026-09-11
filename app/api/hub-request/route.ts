import { NextRequest, NextResponse } from 'next/server'
import { createClient } from '@supabase/supabase-js'

const supabase = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL!,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
)

const HUBS = ['มัยลาภ', 'ขอนแก่น', 'พิษณุโลก', 'สุราษฎร์ธานี'] as const
const APP_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://import-project-zeta.vercel.app'

export async function POST(req: NextRequest) {
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
    const adminLink = `${APP_URL}/admin`
    const message = [
      '🔔 คำขอเปลี่ยน Hub',
      '',
      `ตู้: ${container_name}`,
      `Invoice: ${invoice_no}`,
      `Hub ที่ขอ: ${hub}`,
      `ขอโดย: ${requested_by}`,
      '',
      `👉 ยืนยันที่: ${adminLink}`,
    ].join('\n')
    try {
      await fetch('https://api.line.me/v2/bot/message/push', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${lineToken}`,
        },
        body: JSON.stringify({
          to: lineGroupId,
          messages: [{ type: 'text', text: message }],
        }),
      })
    } catch (e) {
      console.error('LINE notification failed', e)
    }
  }

  return NextResponse.json({ ok: true })
}

export async function PATCH(req: NextRequest) {
  // Admin confirms a request
  const body = await req.json()
  const { invoice_id, container_name } = body

  if (!invoice_id || !container_name) {
    return NextResponse.json({ error: 'missing fields' }, { status: 400 })
  }

  const { confirmed_by, hub_arrival_date } = body

  const { error } = await supabase
    .from('container_hub_requests')
    .update({ status: 'confirmed', confirmed_at: new Date().toISOString(), confirmed_by: confirmed_by ?? null, hub_arrival_date: hub_arrival_date ?? null })
    .eq('invoice_id', invoice_id)
    .eq('container_name', container_name)

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json({ ok: true })
}

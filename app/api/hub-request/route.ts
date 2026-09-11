import { NextRequest, NextResponse } from 'next/server'
import { createClient } from '@supabase/supabase-js'

const supabase = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL!,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
)

const HUBS = ['มัยลาภ', 'ขอนแก่น', 'พิษณุโลก', 'สุราษฎร์ธานี'] as const
const APP_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://my-app-mu-one-69.vercel.app'

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

  // Send email via Resend if API key is configured
  const resendKey = process.env.RESEND_API_KEY
  if (resendKey && adminEmails.length > 0) {
    const deepLink = `${APP_URL}/dashboard/${invoice_id}?container=${encodeURIComponent(container_name)}`
    try {
      const { Resend } = await import('resend')
      const resend = new Resend(resendKey)
      await resend.emails.send({
        from: 'Import PO <noreply@rbs-groups.com>',
        to: adminEmails,
        subject: `[Hub Request] ตู้ ${container_name} → Hub ${hub}`,
        html: `
          <p><strong>${requested_by}</strong> ขอเปลี่ยนปลายทางตู้</p>
          <ul>
            <li>Invoice: <strong>${invoice_no}</strong></li>
            <li>ตู้: <strong>${container_name}</strong></li>
            <li>Hub ที่ขอ: <strong>${hub}</strong></li>
          </ul>
          <p><a href="${deepLink}" style="background:#2563eb;color:white;padding:10px 20px;border-radius:6px;text-decoration:none;display:inline-block">
            ดู Invoice และยืนยัน
          </a></p>
        `,
      })
    } catch (e) {
      console.error('email send failed', e)
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

  const { error } = await supabase
    .from('container_hub_requests')
    .update({ status: 'confirmed', confirmed_at: new Date().toISOString() })
    .eq('invoice_id', invoice_id)
    .eq('container_name', container_name)

  if (error) return NextResponse.json({ error: error.message }, { status: 500 })
  return NextResponse.json({ ok: true })
}

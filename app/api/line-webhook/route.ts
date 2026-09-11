import { NextRequest, NextResponse } from 'next/server'

export async function POST(req: NextRequest) {
  // Webhook endpoint — LINE requires this to exist and return 200
  // Notifications are sent proactively from /api/hub-request, not here
  await req.json()
  return NextResponse.json({ ok: true })
}

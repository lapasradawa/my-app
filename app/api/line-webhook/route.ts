import { NextRequest, NextResponse } from 'next/server'

export async function POST(req: NextRequest) {
  const body = await req.json()
  const events = body.events ?? []

  for (const event of events) {
    const groupId = event.source?.groupId
    const replyToken = event.replyToken

    console.log('LINE event:', JSON.stringify({ type: event.type, groupId, hasReplyToken: !!replyToken }))

    if (groupId && replyToken) {
      const res = await fetch('https://api.line.me/v2/bot/message/reply', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${process.env.LINE_CHANNEL_ACCESS_TOKEN}`,
        },
        body: JSON.stringify({
          replyToken,
          messages: [{ type: 'text', text: `Group ID: ${groupId}` }],
        }),
      })
      const result = await res.text()
      console.log('LINE reply result:', res.status, result)
    }
  }

  return NextResponse.json({ ok: true })
}

import { NextRequest, NextResponse } from 'next/server'

export async function POST(req: NextRequest) {
  const body = await req.json()
  const events = body.events ?? []

  for (const event of events) {
    const groupId = event.source?.groupId
    const replyToken = event.replyToken

    if (groupId && replyToken) {
      // Reply to group with its own ID so admin can copy it
      await fetch('https://api.line.me/v2/bot/message/reply', {
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
    }
  }

  return NextResponse.json({ ok: true })
}

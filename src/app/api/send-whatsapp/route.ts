import { createClient } from '@supabase/supabase-js'
import { NextResponse } from 'next/server'

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL!,
  process.env.SUPABASE_SERVICE_ROLE_KEY!
)

export async function POST(req: Request) {
  try {
    const { message } = (await req.json()) as { message: string }

    if (!message || !message.trim()) {
      return NextResponse.json({ ok: false, message: 'اكتب نص الرسالة أول' }, { status: 400 })
    }

    const { data: branches, error } = await supabaseAdmin
      .from('branches')
      .select('id, name, manager_phone')

    if (error) throw error

    const validBranches = (branches || []).filter(
      (b) => b.manager_phone && b.manager_phone.trim() !== ''
    )

    if (validBranches.length === 0) {
      return NextResponse.json({ ok: true, sent: 0, failed: 0, failedBranches: [] })
    }

    const INSTANCE_ID = process.env.ULTRAMSG_INSTANCE_ID!
    const TOKEN = process.env.ULTRAMSG_TOKEN!

    let sent = 0
    let failed = 0
    const failedBranches: string[] = []

    for (const branch of validBranches) {
      try {
        // ينظف الرقم من أي رموز أو مسافات، يخلي بس الأرقام
        let phone = branch.manager_phone!.replace(/[^0-9]/g, '')

        // يحوّل الصيغة المحلية (05xxxxxxxx) للصيغة الدولية (9665xxxxxxxx)
        if (phone.startsWith('0')) {
          phone = '966' + phone.slice(1)
        }
        // لو كتبه بدون صفر وبدون 966 (يبدأ بـ 5 مباشرة)
        if (phone.startsWith('5') && phone.length === 9) {
          phone = '966' + phone
        }

        const res = await fetch(`https://api.ultramsg.com/${INSTANCE_ID}/messages/chat`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            token: TOKEN,
            to: phone,
            body: message,
          }),
        })

        if (res.ok) {
          sent++
        } else {
          failed++
          failedBranches.push(branch.name)
        }
      } catch {
        failed++
        failedBranches.push(branch.name)
      }
      // فاصل بسيط بين كل رسالة عشان ما تنحظر من UltraMsg
      await new Promise((r) => setTimeout(r, 300))
    }

    return NextResponse.json({ ok: true, sent, failed, failedBranches })
  } catch (err: any) {
    return NextResponse.json({ ok: false, message: err.message || 'خطأ غير متوقع' }, { status: 500 })
  }
}
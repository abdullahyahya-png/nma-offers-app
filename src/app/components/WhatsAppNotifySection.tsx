'use client'

// الصقه بمسار: src/app/components/WhatsAppNotifySection.tsx
// يحتاج: عمود manager_phone بجدول branches (شغّل SQL أول)
// يحتاج: /api/send-whatsapp route شغال ومربوط بـ UltraMsg

import { useEffect, useState } from 'react'
import { supabase } from '../../lib/supabase'
import { Send, Save, Phone } from 'lucide-react'

interface Branch {
  id: string
  name: string
  manager_phone: string | null
}

export default function WhatsAppNotifySection() {
  const [branches, setBranches] = useState<Branch[]>([])
  const [phoneInputs, setPhoneInputs] = useState<Record<string, string>>({})
  const [savingId, setSavingId] = useState<string | null>(null)
  const [message, setMessage] = useState('🟢 تنبيه من نظام عروض NMA\nفيه تحديث عروض جديد، الرجاء الدخول للنظام ومراجعة المنتجات وطباعة الملصقات.')
  const [sending, setSending] = useState(false)
  const [status, setStatus] = useState('')

  useEffect(() => {
    const load = async () => {
      const { data } = await supabase.from('branches').select('id, name, manager_phone').order('name')
      if (data) {
        setBranches(data)
        const inputs: Record<string, string> = {}
        data.forEach((b) => { inputs[b.id] = b.manager_phone || '' })
        setPhoneInputs(inputs)
      }
    }
    load()
  }, [])

  const handleSavePhone = async (branchId: string) => {
    setSavingId(branchId)
    const phone = phoneInputs[branchId]?.trim() || null
    const { error } = await supabase.from('branches').update({ manager_phone: phone }).eq('id', branchId)
    setSavingId(null)
    if (error) {
      setStatus(`خطأ بحفظ رقم الفرع: ${error.message}`)
    } else {
      setStatus('تم حفظ الرقم بنجاح')
      setBranches((prev) => prev.map((b) => (b.id === branchId ? { ...b, manager_phone: phone } : b)))
    }
  }

  const handleSendToAll = async () => {
    if (!message.trim()) {
      setStatus('اكتب نص الرسالة أول')
      return
    }
    const numbersCount = branches.filter((b) => b.manager_phone && b.manager_phone.trim() !== '').length
    if (numbersCount === 0) {
      setStatus('ما فيه أي رقم محفوظ لأي فرع')
      return
    }
    if (!confirm(`بترسل الرسالة لـ ${numbersCount} فرع، متأكد؟`)) return

    setSending(true)
    setStatus('جاري الإرسال...')
    try {
      const res = await fetch('/api/send-whatsapp', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message }),
      })
      const data = await res.json()
      if (data.ok) {
        setStatus(
          `تم الإرسال بنجاح لـ ${data.sent} فرع` +
          (data.failed > 0 ? ` — فشل الإرسال لـ ${data.failed} فرع (${data.failedBranches.join('، ')})` : '')
        )
      } else {
        setStatus(`خطأ: ${data.message}`)
      }
    } catch (err: any) {
      setStatus(`خطأ بالاتصال: ${err.message}`)
    }
    setSending(false)
  }

  return (
    <div className="space-y-6 max-w-3xl">
      {status && (
        <div className="p-3.5 bg-[var(--yellow)]/15 border-2 border-[var(--yellow)]/40 rounded-lg text-sm text-[#8a6300] font-bold">
          {status}
        </div>
      )}

      <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm">
        <h2 className="font-black text-sm text-[var(--navy)] mb-1 flex items-center gap-2">
          <Send size={16} />
          إرسال إشعار واتساب لكل الفروع
        </h2>
        <p className="text-xs text-gray-500 font-medium mb-4">
          الرسالة تنرسل لكل فرع عنده رقم محفوظ بالأسفل
        </p>
        <textarea
          value={message}
          onChange={(e) => setMessage(e.target.value)}
          rows={4}
          className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-3 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
        />
        <button
          onClick={handleSendToAll}
          disabled={sending}
          className="mt-4 bg-emerald-600 hover:bg-emerald-500 text-white px-6 py-3 rounded-xl font-black transition-colors disabled:opacity-50 flex items-center gap-2"
        >
          <Send size={16} />
          {sending ? 'جاري الإرسال...' : 'إرسال للكل'}
        </button>
      </div>

      <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
        <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
          <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
            <Phone size={16} />
            أرقام واتساب مدراء الفروع
          </h2>
        </div>
        <div className="divide-y-2 divide-[var(--navy)]/10">
          {branches.map((b) => (
            <div key={b.id} className="flex items-center gap-3 p-4">
              <span className="text-sm font-bold text-[var(--navy)] w-48 shrink-0 truncate">{b.name}</span>
              <input
                value={phoneInputs[b.id] || ''}
                onChange={(e) => setPhoneInputs((prev) => ({ ...prev, [b.id]: e.target.value }))}
                placeholder="مثال: 0537196361"
                dir="ltr"
                className="flex-1 bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
              />
              <button
                onClick={() => handleSavePhone(b.id)}
                disabled={savingId === b.id}
                className="flex items-center gap-1.5 bg-[var(--navy)] hover:bg-[#0f1a4d] text-white text-xs font-bold px-4 py-2 rounded-lg transition-colors disabled:opacity-50 shrink-0"
              >
                <Save size={13} />
                حفظ
              </button>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}

'use client'

// الصقه بمسار: src/app/components/AuditOverviewSection.tsx
// يحتاج جدول label_checks (شغّلته بالـ SQL قبل)

import { useEffect, useState } from 'react'
import { supabase } from '../../lib/supabase'
import { ClipboardCheck } from 'lucide-react'

interface BranchAuditRow {
  branch_id: string
  branch_name: string
  total: number
  checked: number
}

export default function AuditOverviewSection() {
  const [rows, setRows] = useState<BranchAuditRow[]>([])
  const [loading, setLoading] = useState(true)

  const loadData = async () => {
    setLoading(true)
    const { data: branches } = await supabase.from('branches').select('id, name').order('name')
    const { data: checks } = await supabase.from('label_checks').select('branch_id, is_checked')

    const map: Record<string, { total: number; checked: number }> = {}
    checks?.forEach((c) => {
      if (!map[c.branch_id]) map[c.branch_id] = { total: 0, checked: 0 }
      map[c.branch_id].total++
      if (c.is_checked) map[c.branch_id].checked++
    })

    const result: BranchAuditRow[] = (branches || []).map((b) => ({
      branch_id: b.id,
      branch_name: b.name,
      total: map[b.id]?.total || 0,
      checked: map[b.id]?.checked || 0,
    }))

    setRows(result)
    setLoading(false)
  }

  useEffect(() => {
    loadData()
  }, [])

  return (
    <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm max-w-3xl">
      <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
        <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
          <ClipboardCheck size={16} />
          نسبة تدقيق الملصقات لكل فرع
        </h2>
        <p className="text-xs text-gray-500 font-medium mt-1">
          يوضح كم منتج أكّد الفرع إن عليه ملصق فعلي بالرف، من إجمالي منتجاته
        </p>
      </div>
      <div className="divide-y-2 divide-[var(--navy)]/10 max-h-[650px] overflow-y-auto">
        {loading && <p className="p-6 text-center text-gray-400 text-sm">جاري التحميل...</p>}
        {!loading && rows.length === 0 && (
          <p className="p-6 text-center text-gray-400 text-sm">ما فيه بيانات تدقيق بعد</p>
        )}
        {!loading && rows.map((row) => {
          const pct = row.total > 0 ? Math.round((row.checked / row.total) * 100) : 0
          const barColor = pct === 100 ? 'bg-emerald-500' : pct >= 50 ? 'bg-[var(--yellow)]' : 'bg-[var(--red)]'
          return (
            <div key={row.branch_id} className="p-4">
              <div className="flex items-center justify-between mb-2">
                <p className="text-sm font-bold text-[var(--navy)]">{row.branch_name}</p>
                <span className="text-xs font-black text-[var(--navy)]">
                  {row.total === 0 ? 'ما بدأ تدقيق بعد' : `${row.checked} من ${row.total} (${pct}%)`}
                </span>
              </div>
              {row.total > 0 && (
                <div className="bg-gray-100 rounded-full h-2.5 overflow-hidden">
                  <div className={`${barColor} h-full transition-all`} style={{ width: `${pct}%` }} />
                </div>
              )}
            </div>
          )
        })}
      </div>
    </div>
  )
}
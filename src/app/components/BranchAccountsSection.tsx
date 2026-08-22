'use client'

// الصقه بمسار: src/app/components/BranchAccountsSection.tsx
// يحتاج SQL الجديد (simplify_login.sql) — كلمة سر بس، بدون اسم مستخدم منفصل

import { useEffect, useState } from 'react'
import { supabase } from '../../lib/supabase'
import { KeyRound, Save, CheckCircle2 } from 'lucide-react'

interface BranchStatus {
  branch_id: string
  name: string
  has_password: boolean
}

export default function BranchAccountsSection() {
  const [branches, setBranches] = useState<BranchStatus[]>([])
  const [passwordInputs, setPasswordInputs] = useState<Record<string, string>>({})
  const [savingId, setSavingId] = useState<string | null>(null)
  const [savedIds, setSavedIds] = useState<Set<string>>(new Set())
  const [status, setStatus] = useState('')

  const loadData = async () => {
    const { data } = await supabase.from('branches_login_status').select('*').order('name')
    if (data) setBranches(data)
  }

  useEffect(() => {
    loadData()
  }, [])

  const handleSave = async (branchId: string) => {
    const password = (passwordInputs[branchId] || '').trim()

    if (!password) {
      setStatus('اكتب كلمة سر جديدة أول')
      return
    }
    if (password.length < 4) {
      setStatus('كلمة السر لازم تكون 4 أحرف/أرقام على الأقل')
      return
    }

    setSavingId(branchId)
    const { error } = await supabase.rpc('set_branch_password', {
      p_branch_id: branchId,
      p_password: password,
    })
    setSavingId(null)

    if (error) {
      setStatus(`خطأ: ${error.message}`)
    } else {
      setStatus('تم حفظ كلمة السر بنجاح')
      setBranches((prev) => prev.map((b) => (b.branch_id === branchId ? { ...b, has_password: true } : b)))
      setPasswordInputs((prev) => ({ ...prev, [branchId]: '' }))
      setSavedIds((prev) => new Set(prev).add(branchId))
      setTimeout(() => {
        setSavedIds((prev) => {
          const next = new Set(prev)
          next.delete(branchId)
          return next
        })
      }, 3000)
    }
  }

  return (
    <div className="space-y-4 max-w-3xl">
      {status && (
        <div className="p-3.5 bg-[var(--yellow)]/15 border-2 border-[var(--yellow)]/40 rounded-lg text-sm text-[#8a6300] font-bold">
          {status}
        </div>
      )}

      <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
        <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
          <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
            <KeyRound size={16} />
            كلمات سر دخول الفروع ({branches.length})
          </h2>
          <p className="text-xs text-gray-500 font-medium mt-1">
            اسم الفرع نفسه هو اسم الدخول. حط كلمة سر لكل فرع — أي موظف يعرفها يقدر يدخل بنفس الحساب.
          </p>
        </div>
        <div className="divide-y-2 divide-[var(--navy)]/10 max-h-[650px] overflow-y-auto">
          {branches.map((branch) => {
            const justSaved = savedIds.has(branch.branch_id)
            return (
              <div key={branch.branch_id} className="p-4 flex flex-wrap items-center gap-3">
                <div className="w-52 shrink-0">
                  <p className="text-sm font-bold text-[var(--navy)] truncate">{branch.name}</p>
                  {branch.has_password ? (
                    <p className="text-[11px] text-emerald-600 font-bold flex items-center gap-1">
                      <CheckCircle2 size={11} />
                      عنده كلمة سر
                    </p>
                  ) : (
                    <p className="text-[11px] text-gray-400 font-bold">ما عنده كلمة سر بعد</p>
                  )}
                </div>
                <input
                  value={passwordInputs[branch.branch_id] || ''}
                  onChange={(e) => setPasswordInputs((prev) => ({ ...prev, [branch.branch_id]: e.target.value }))}
                  placeholder="كلمة سر جديدة"
                  dir="ltr"
                  className="flex-1 min-w-[160px] bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                />
                <button
                  onClick={() => handleSave(branch.branch_id)}
                  disabled={savingId === branch.branch_id}
                  className={`flex items-center gap-1.5 text-xs font-bold px-4 py-2 rounded-lg transition-colors disabled:opacity-50 shrink-0 ${
                    justSaved ? 'bg-emerald-600 text-white' : 'bg-[var(--navy)] hover:bg-[#0f1a4d] text-white'
                  }`}
                >
                  {justSaved ? <CheckCircle2 size={13} /> : <Save size={13} />}
                  {justSaved ? 'تم الحفظ' : 'حفظ'}
                </button>
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}

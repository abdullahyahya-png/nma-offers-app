'use client'

// الملف الأصلي العادي (بانر + زر تثبيت) — يُستخدم بلوحة الأدمن (وأي صفحة ثانية تحتاج نفس السلوك القديم)
// يختلف عن InstallPWAButtonAuto.tsx (خاص بواجهة الطباعة، يفتح نافذة التثبيت تلقائياً بدون بانر وسيط)

import { useEffect, useState } from 'react'
import { Download, X } from 'lucide-react'

export default function InstallPWAButton() {
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null)
  const [showBanner, setShowBanner] = useState(false)
  const [dismissed, setDismissed] = useState(false)
  const [isIOS, setIsIOS] = useState(false)

  useEffect(() => {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').catch(() => {})
    }

    const alreadyInstalled = window.matchMedia('(display-mode: standalone)').matches
    if (alreadyInstalled) return

    const ua = window.navigator.userAgent
    const iosDevice = /iPad|iPhone|iPod/.test(ua) && !(window as any).MSStream
    setIsIOS(iosDevice)

    // يظهر البانر فوراً بدون انتظار (تعليمات يدوية افتراضياً، أو زر تثبيت حقيقي لو المتصفح دعمه)
    setShowBanner(true)

    const handler = (e: any) => {
      e.preventDefault()
      setDeferredPrompt(e)
    }
    window.addEventListener('beforeinstallprompt', handler)
    return () => window.removeEventListener('beforeinstallprompt', handler)
  }, [])

  const handleInstallClick = async () => {
    if (deferredPrompt) {
      deferredPrompt.prompt()
      await deferredPrompt.userChoice
      setDeferredPrompt(null)
      setShowBanner(false)
    }
  }

  if (!showBanner || dismissed) return null

  return (
    <div className="fixed bottom-4 inset-x-4 sm:inset-x-auto sm:right-4 sm:left-auto sm:w-96 z-50 bg-[var(--navy)] text-white rounded-2xl shadow-2xl p-4">
      <div className="flex items-start gap-3">
        <Download size={20} className="shrink-0 mt-0.5" />
        <div className="flex-1 min-w-0">
          <p className="text-sm font-black mb-1">ثبّت التطبيق</p>
          {deferredPrompt ? (
            <>
              <p className="text-xs text-white/80 leading-relaxed mb-2">ثبّت التطبيق على جهازك للوصول السريع</p>
              <button
                onClick={handleInstallClick}
                className="bg-white text-[var(--navy)] text-xs font-bold px-3 py-1.5 rounded-lg"
              >
                تثبيت الآن
              </button>
            </>
          ) : isIOS ? (
            <p className="text-xs text-white/80 leading-relaxed">
              اضغط زر المشاركة ⬆️ بأسفل الشاشة، ثم اختر "إضافة إلى الشاشة الرئيسية"
            </p>
          ) : (
            <p className="text-xs text-white/80 leading-relaxed">
              افتح قائمة المتصفح ⋮ واختر "تثبيت التطبيق" أو "Install app"
            </p>
          )}
        </div>
        <button onClick={() => setDismissed(true)} className="text-white/60 hover:text-white shrink-0">
          <X size={16} />
        </button>
      </div>
    </div>
  )
}
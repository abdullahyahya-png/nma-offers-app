'use client'

// الصقه بمسار: src/app/components/InstallPWAButton.tsx (نفس المكان، يستبدل القديم)

import { useEffect, useState } from 'react'
import { Download, X } from 'lucide-react'

export default function InstallPWAButton() {
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null)
  const [showBanner, setShowBanner] = useState(false)
  const [dismissed, setDismissed] = useState(false)
  const [isIOS, setIsIOS] = useState(false)
  const [showManualSteps, setShowManualSteps] = useState(false)

  useEffect(() => {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').catch(() => {})
    }

    // لو مثبت أصلاً، ما نعرض شي
    const alreadyInstalled = window.matchMedia('(display-mode: standalone)').matches
    if (alreadyInstalled) return

    const ua = window.navigator.userAgent
    const iosDevice = /iPad|iPhone|iPod/.test(ua) && !(window as any).MSStream
    if (iosDevice) {
      setIsIOS(true)
      setShowBanner(true)
      return
    }

    const handler = (e: any) => {
      e.preventDefault()
      setDeferredPrompt(e)
      setShowBanner(true)
    }
    window.addEventListener('beforeinstallprompt', handler)

    // لو كروم ما أطلق الحدث التلقائي خلال 3 ثواني، نوري تعليمات يدوية بدل الانتظار عليه
    const fallbackTimer = setTimeout(() => {
      setShowBanner((prev) => prev || true)
      setShowManualSteps((prevManual) => {
        // بس لو ما وصلنا deferredPrompt لسا
        return true
      })
    }, 3000)

    return () => {
      window.removeEventListener('beforeinstallprompt', handler)
      clearTimeout(fallbackTimer)
    }
  }, [])

  const handleInstall = async () => {
    if (!deferredPrompt) {
      setShowManualSteps(true)
      return
    }
    deferredPrompt.prompt()
    await deferredPrompt.userChoice
    setDeferredPrompt(null)
    setShowBanner(false)
  }

  if (!showBanner || dismissed) return null

  const showingManual = isIOS || (showManualSteps && !deferredPrompt)

  return (
    <div className="fixed bottom-4 inset-x-4 sm:inset-x-auto sm:right-4 sm:left-auto sm:w-96 z-50 bg-[var(--navy)] text-white rounded-2xl shadow-2xl p-4">
      <div className="flex items-start gap-3">
        <Download size={20} className="shrink-0 mt-0.5" />
        <div className="flex-1 min-w-0">
          <p className="text-sm font-black mb-1">ثبّت التطبيق</p>
          {showingManual ? (
            isIOS ? (
              <p className="text-xs text-white/80 leading-relaxed">
                اضغط زر المشاركة ⬆️ بأسفل الشاشة، ثم اختر "إضافة إلى الشاشة الرئيسية"
              </p>
            ) : (
              <p className="text-xs text-white/80 leading-relaxed">
                من قائمة المتصفح (⋮ فوق يمين الشاشة)، اختر "تثبيت التطبيق" أو "إضافة إلى الشاشة الرئيسية"
              </p>
            )
          ) : (
            <p className="text-xs text-white/70">وصول أسرع من شاشتك الرئيسية</p>
          )}
        </div>
        <button onClick={() => setDismissed(true)} className="text-white/60 hover:text-white shrink-0">
          <X size={16} />
        </button>
      </div>
      {!showingManual && (
        <button
          onClick={handleInstall}
          className="w-full mt-3 bg-white text-[var(--navy)] text-sm font-bold py-2 rounded-lg"
        >
          تثبيت الآن
        </button>
      )}
    </div>
  )
}
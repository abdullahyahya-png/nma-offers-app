'use client'

// الصقه بمسار: src/app/components/InstallPWAButton.tsx
// استخدمه بصفحة الأدمن وصفحة الفرع (استيراد وعرض مرة وحدة بكل صفحة)

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

    // آيفون/آيباد ما يدعم beforeinstallprompt، نعرض له تعليمات يدوية
    const ua = window.navigator.userAgent
    const iosDevice = /iPad|iPhone|iPod/.test(ua) && !(window as any).MSStream
    const alreadyInstalled = window.matchMedia('(display-mode: standalone)').matches
    if (iosDevice && !alreadyInstalled) {
      setIsIOS(true)
      setShowBanner(true)
    }

    const handler = (e: any) => {
      e.preventDefault()
      setDeferredPrompt(e)
      setShowBanner(true)
    }
    window.addEventListener('beforeinstallprompt', handler)
    return () => window.removeEventListener('beforeinstallprompt', handler)
  }, [])

  const handleInstall = async () => {
    if (!deferredPrompt) return
    deferredPrompt.prompt()
    await deferredPrompt.userChoice
    setDeferredPrompt(null)
    setShowBanner(false)
  }

  if (!showBanner || dismissed) return null

  return (
    <div className="fixed bottom-4 inset-x-4 sm:inset-x-auto sm:right-4 sm:left-auto sm:w-80 z-50 bg-[var(--navy)] text-white rounded-2xl shadow-2xl p-4 flex items-center gap-3">
      <Download size={20} className="shrink-0" />
      <div className="flex-1 min-w-0">
        <p className="text-sm font-black">ثبّت التطبيق</p>
        {isIOS ? (
          <p className="text-xs text-white/70">اضغط زر المشاركة ⬆️ ثم "إضافة إلى الشاشة الرئيسية"</p>
        ) : (
          <p className="text-xs text-white/70">وصول أسرع من شاشتك الرئيسية</p>
        )}
      </div>
      {!isIOS && (
        <button onClick={handleInstall} className="bg-white text-[var(--navy)] text-xs font-bold px-3 py-1.5 rounded-lg shrink-0">
          تثبيت
        </button>
      )}
      <button onClick={() => setDismissed(true)} className="text-white/60 hover:text-white shrink-0">
        <X size={16} />
      </button>
    </div>
  )
}
'use client'

// الصقه بمسار: src/app/components/InstallPWAButtonAuto.tsx
// خاص بواجهة الطباعة: يفتح نافذة التثبيت الأصلية تلقائياً بدون ضغطة وسيطة (Android/Chrome/Edge)
// آيفون فقط يحتاج تعليمات يدوية (نظام iOS ما يدعم فتح نافذة تثبيت تلقائية إطلاقاً)

import { useEffect, useState } from 'react'
import { Download, X } from 'lucide-react'

export default function InstallPWAButtonAuto() {
  const [showIOSBanner, setShowIOSBanner] = useState(false)
  const [dismissed, setDismissed] = useState(false)

  useEffect(() => {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').catch(() => {})
    }

    const alreadyInstalled = window.matchMedia('(display-mode: standalone)').matches
    if (alreadyInstalled) return

    const ua = window.navigator.userAgent
    const iosDevice = /iPad|iPhone|iPod/.test(ua) && !(window as any).MSStream

    if (iosDevice) {
      setShowIOSBanner(true)
      return
    }

    // أي جهاز/متصفح غير آيفون: نفتح نافذة التثبيت الأصلية فوراً أول ما المتصفح يسمح بيها
    const handler = (e: any) => {
      e.preventDefault()
      e.prompt()
    }
    window.addEventListener('beforeinstallprompt', handler)
    return () => window.removeEventListener('beforeinstallprompt', handler)
  }, [])

  if (!showIOSBanner || dismissed) return null

  return (
    <div className="fixed bottom-4 inset-x-4 sm:inset-x-auto sm:right-4 sm:left-auto sm:w-96 z-50 bg-[var(--navy)] text-white rounded-2xl shadow-2xl p-4">
      <div className="flex items-start gap-3">
        <Download size={20} className="shrink-0 mt-0.5" />
        <div className="flex-1 min-w-0">
          <p className="text-sm font-black mb-1">ثبّت التطبيق</p>
          <p className="text-xs text-white/80 leading-relaxed">
            اضغط زر المشاركة ⬆️ بأسفل الشاشة، ثم اختر "إضافة إلى الشاشة الرئيسية"
          </p>
        </div>
        <button onClick={() => setDismissed(true)} className="text-white/60 hover:text-white shrink-0">
          <X size={16} />
        </button>
      </div>
    </div>
  )
}
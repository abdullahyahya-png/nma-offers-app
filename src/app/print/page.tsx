'use client'

import { useEffect, useRef, useState } from 'react'
import jsPDF from 'jspdf'
import { supabase } from '../../lib/supabase'
import { renderPdfToCanvas } from '../../lib/pdfBackground'
import { Search, Printer, Download, Package, ListPlus, Layers, X } from 'lucide-react'

interface OfferItem {
  id: string
  barcode: string
  product_name: string
  previous_price: number
  offer_price: number
  is_active?: boolean
}

interface LabelData {
  name: string
  offerPriceText: string
  prevPriceText: string
  barcodeText: string
}

const REF_W = 618
const REF_H = 1034
const SCALE = 4

const POS = {
  offerPrice: { xFrac: 0.45, yFrac: 545 / REF_H, fontPx: 90 },
  prevPrice: { xFrac: 0.47, yFrac: 705 / REF_H, fontPx: 58 },
  name: { xFrac: 0.5, yFrac: 845 / REF_H, fontPx: 34, lineSpacingPx: 38 },
  barcode: { xFrac: 0.5, yFrac: 968 / REF_H, fontPx: 33 },
}

const NAVY = '#150971'
const RED = '#C00000'

function itemToLabelData(item: OfferItem): LabelData {
  return {
    name: item.product_name,
    offerPriceText: item.offer_price.toFixed(2),
    prevPriceText: item.previous_price.toFixed(2),
    barcodeText: item.barcode,
  }
}

export default function PrintPage() {
  const [allItems, setAllItems] = useState<OfferItem[]>([])
  const [searchText, setSearchText] = useState('')
  const [status, setStatus] = useState('')
  const [bgReady, setBgReady] = useState(false)
  const [printingId, setPrintingId] = useState<string | null>(null)
  const [printQueue, setPrintQueue] = useState<OfferItem[]>([])
  const [bulkBarcodesText, setBulkBarcodesText] = useState('')
  const [generatingQueue, setGeneratingQueue] = useState(false)
  const bgImageRef = useRef<HTMLCanvasElement | null>(null)

  const queueIds = new Set(printQueue.map((i) => i.id))

  const addToQueue = (item: OfferItem) => {
    if (queueIds.has(item.id)) return
    setPrintQueue((prev) => [...prev, item])
  }

  const removeFromQueue = (id: string) => {
    setPrintQueue((prev) => prev.filter((i) => i.id !== id))
  }

  const clearQueue = () => setPrintQueue([])

  const handleBulkAdd = () => {
    const codes = bulkBarcodesText
      .split(/[\n,\s]+/)
      .map((c) => c.trim())
      .filter(Boolean)

    if (codes.length === 0) {
      setStatus('اكتب أو الصق باركودات أول')
      return
    }

    const codeSet = new Set(codes)
    const matched = allItems.filter((item) => codeSet.has(item.barcode))
    const matchedBarcodes = new Set(matched.map((m) => m.barcode))
    const notFound = codes.filter((c) => !matchedBarcodes.has(c))

    setPrintQueue((prev) => {
      const existingIds = new Set(prev.map((i) => i.id))
      const newOnes = matched.filter((m) => !existingIds.has(m.id))
      return [...prev, ...newOnes]
    })

    setBulkBarcodesText('')
    setStatus(
      `تمت إضافة ${matched.length} منتج لقائمة الطباعة` +
      (notFound.length > 0 ? ` — ${notFound.length} باركود ما تم لقاه بالعروض الحالية` : '')
    )
  }

  const filteredItems = searchText.trim().length >= 1
    ? allItems.filter(
        (item) =>
          item.barcode.includes(searchText.trim()) ||
          item.product_name.includes(searchText.trim())
      )
    : allItems

  useEffect(() => {
    if (!status) return
    const timer = setTimeout(() => setStatus(''), 4000)
    return () => clearTimeout(timer)
  }, [status])

  const fetchOffers = async () => {
    const { data } = await supabase
      .from('offer_items')
      .select('*')
      .order('created_at', { ascending: false })
    if (data) setAllItems(data.filter((i) => i.is_active !== false))
  }

  useEffect(() => {
    fetchOffers()

    const { data: bgData } = supabase.storage.from('label-assets').getPublicUrl('label-bg.pdf')
    renderPdfToCanvas(`${bgData.publicUrl}?t=${Date.now()}`, REF_W * SCALE, REF_H * SCALE)
      .then((canvas) => {
        bgImageRef.current = canvas
        setBgReady(true)
      })
      .catch(() => {
        setBgReady(false)
      })

    // Realtime: أي تحديث ينزّله الأدمن يظهر هنا فوراً بدون Refresh
    const channel = supabase
      .channel('print-general-realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'offer_items' }, () => fetchOffers())
      .subscribe()

    return () => {
      supabase.removeChannel(channel)
    }
  }, [])

  const renderLabelCanvas = async (data: LabelData): Promise<HTMLCanvasElement> => {
    const canvas = document.createElement('canvas')
    canvas.width = REF_W * SCALE
    canvas.height = REF_H * SCALE
    const ctx = canvas.getContext('2d')!
    ctx.textBaseline = 'middle'

    if (bgImageRef.current) ctx.drawImage(bgImageRef.current, 0, 0, canvas.width, canvas.height)

    ctx.textAlign = 'center'
    ctx.direction = 'rtl' as any

    ctx.font = `900 ${POS.offerPrice.fontPx * SCALE}px Tajawal`
    ctx.fillStyle = NAVY
    ctx.fillText(data.offerPriceText, canvas.width * POS.offerPrice.xFrac, canvas.height * POS.offerPrice.yFrac)

    ctx.font = `700 ${POS.prevPrice.fontPx * SCALE}px Tajawal`
    ctx.fillStyle = RED
    const prevX = canvas.width * POS.prevPrice.xFrac
    const prevY = canvas.height * POS.prevPrice.yFrac
    ctx.fillText(data.prevPriceText, prevX, prevY)
    const prevWidth = ctx.measureText(data.prevPriceText).width
    ctx.beginPath()
    ctx.strokeStyle = RED
    ctx.lineWidth = 3 * SCALE
    ctx.moveTo(prevX - prevWidth / 2, prevY)
    ctx.lineTo(prevX + prevWidth / 2, prevY)
    ctx.stroke()

    ctx.font = `700 ${POS.name.fontPx * SCALE}px Tajawal`
    ctx.fillStyle = NAVY
    const maxWidth = canvas.width * 0.82
    const words = data.name.split(' ')
    const lines: string[] = []
    let currentLine = ''
    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word
      if (ctx.measureText(testLine).width > maxWidth && currentLine) {
        lines.push(currentLine)
        currentLine = word
      } else {
        currentLine = testLine
      }
    }
    if (currentLine) lines.push(currentLine)
    const nameLines = lines.slice(0, 2)
    const lineSpacing = POS.name.lineSpacingPx * SCALE
    const startY = canvas.height * POS.name.yFrac - ((nameLines.length - 1) * lineSpacing) / 2
    nameLines.forEach((line, i) => {
      ctx.fillText(line, canvas.width * POS.name.xFrac, startY + i * lineSpacing)
    })

    ctx.font = `700 ${POS.barcode.fontPx * SCALE}px Tajawal`
    ctx.fillStyle = RED
    ctx.fillText(data.barcodeText, canvas.width * POS.barcode.xFrac, canvas.height * POS.barcode.yFrac)

    return canvas
  }

  const handleQueueAction = async (mode: 'download' | 'print') => {
    if (printQueue.length === 0) {
      setStatus('قائمة الطباعة فاضية — أضف منتجات أول')
      return
    }
    if (!bgReady || !bgImageRef.current) {
      setStatus('جاري تحميل قالب الملصق، حاول بعد ثانيتين')
      return
    }
    setGeneratingQueue(true)
    try {
      await document.fonts.load('900 90px Tajawal')
      await document.fonts.load('700 58px Tajawal')
      await document.fonts.load('700 34px Tajawal')

      const doc = new jsPDF({ unit: 'pt', format: [296.28, 496.2], compress: true })
      for (let i = 0; i < printQueue.length; i++) {
        setStatus(`جاري توليد الملصق ${i + 1} من ${printQueue.length}...`)
        await new Promise((r) => setTimeout(r, 0))
        const canvas = await renderLabelCanvas(itemToLabelData(printQueue[i]))
        if (i > 0) doc.addPage([296.28, 496.2])
        doc.addImage(canvas, 'PNG', 0, 0, 296.28, 496.2, undefined, 'FAST')
      }

      if (mode === 'print') {
        doc.autoPrint()
        const blobUrl = doc.output('bloburl')
        window.open(blobUrl as unknown as string, '_blank')
        setStatus(`تم توليد ${printQueue.length} ملصق وفتح نافذة الطباعة`)
      } else {
        doc.save('ملصقات_مختارة.pdf')
        setStatus(`تم تحميل ${printQueue.length} ملصق بملف واحد`)
      }
    } catch (err: any) {
      setStatus(`صار خطأ: ${err?.message || 'غير معروف'}`)
    } finally {
      setGeneratingQueue(false)
    }
  }

  const handleAction = async (item: OfferItem, mode: 'download' | 'print') => {
    if (!bgReady || !bgImageRef.current) {
      setStatus('جاري تحميل قالب الملصق، حاول بعد ثانيتين')
      return
    }
    setPrintingId(item.id)
    setStatus('جاري تجهيز الملصق بأعلى جودة، لحظات...')
    try {
      await document.fonts.load('900 90px Tajawal')
      await document.fonts.load('700 58px Tajawal')
      await document.fonts.load('700 34px Tajawal')

      const canvas = await renderLabelCanvas(itemToLabelData(item))
      const doc = new jsPDF({ unit: 'pt', format: [296.28, 496.2], compress: true })
      doc.addImage(canvas, 'PNG', 0, 0, 296.28, 496.2, undefined, 'FAST')

      if (mode === 'print') {
        doc.autoPrint()
        const blobUrl = doc.output('bloburl')
        window.open(blobUrl as unknown as string, '_blank')
        setStatus('تم فتح نافذة الطباعة')
      } else {
        doc.save(`ملصق_${item.barcode}.pdf`)
        setStatus('تم تحميل الملصق بنجاح')
      }
    } catch (err: any) {
      setStatus(`صار خطأ: ${err?.message || 'غير معروف'}`)
    } finally {
      setPrintingId(null)
    }
  }

  return (
    <div className="min-h-screen bg-[var(--background)]">
      <header className="bg-white border-b-4 border-[var(--navy)]">
        <div className="max-w-4xl mx-auto px-4 py-4 flex items-center gap-3">
          <img src="/logo.png" alt="شعار العروض" className="w-12 h-12 object-contain shrink-0" />
          <div>
            <p className="text-[var(--red)] text-[11px] font-bold">واجهة الطباعة السريعة</p>
            <h1 className="text-[var(--navy)] text-base font-black">عروض عامة</h1>
          </div>
        </div>
      </header>

      <div className="max-w-4xl mx-auto px-4 py-6 space-y-4">
        {status && (
          <div className="p-3.5 bg-[var(--yellow)]/15 border-2 border-[var(--yellow)]/40 rounded-lg text-sm text-[#8a6300] font-bold">
            {status}
          </div>
        )}

        <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-4 shadow-sm">
          <div className="relative">
            <Search size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400" />
            <input
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              placeholder="ابحث بالاسم أو الباركود"
              autoFocus
              className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-3 pr-9 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
            />
          </div>
        </div>

        {/* إضافة عدة باركودات دفعة وحدة */}
        <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-4 shadow-sm">
          <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2 mb-2">
            <ListPlus size={16} />
            إضافة عدة باركودات دفعة وحدة
          </h2>
          <p className="text-xs text-gray-500 font-medium mb-3">
            الصق أو اكتب عدة باركودات (كل وحدة بسطر، أو مفصولة بفاصلة أو مسافة)
          </p>
          <textarea
            value={bulkBarcodesText}
            onChange={(e) => setBulkBarcodesText(e.target.value)}
            rows={3}
            placeholder="6281007020001&#10;6281007020002&#10;6281007020003"
            dir="ltr"
            className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-3 text-sm text-[var(--navy)] font-medium mb-3 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20 resize-none"
          />
          <button
            onClick={handleBulkAdd}
            className="flex items-center gap-1.5 bg-[var(--navy)] hover:bg-[#0f1a4d] text-white text-xs font-bold px-4 py-2 rounded-lg transition-colors"
          >
            <ListPlus size={13} />
            أضف الكل لقائمة الطباعة
          </button>
        </div>

        {/* قائمة الطباعة المختارة */}
        {printQueue.length > 0 && (
          <div className="bg-emerald-50 rounded-2xl border-2 border-emerald-300 overflow-hidden shadow-sm">
            <div className="p-4 border-b-2 border-emerald-200 bg-emerald-100/50 flex flex-wrap items-center justify-between gap-3">
              <h2 className="font-black text-sm text-emerald-800 flex items-center gap-2">
                <Layers size={16} />
                قائمة الطباعة ({printQueue.length})
              </h2>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => handleQueueAction('download')}
                  disabled={generatingQueue}
                  className="flex items-center gap-1.5 bg-white border-2 border-emerald-300 hover:bg-emerald-50 text-emerald-700 text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
                >
                  <Download size={13} />
                  تحميل الكل
                </button>
                <button
                  onClick={() => handleQueueAction('print')}
                  disabled={generatingQueue}
                  className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
                >
                  <Printer size={13} />
                  {generatingQueue ? 'جاري التوليد...' : 'طباعة الكل'}
                </button>
                <button
                  onClick={clearQueue}
                  className="text-emerald-700 hover:text-[var(--red)] text-xs font-bold underline"
                >
                  تفريغ
                </button>
              </div>
            </div>
            <div className="divide-y divide-emerald-200 max-h-64 overflow-y-auto">
              {printQueue.map((item) => (
                <div key={item.id} className="flex items-center justify-between gap-3 p-2.5 px-4">
                  <div className="min-w-0">
                    <p className="text-sm font-bold text-emerald-900 truncate">{item.product_name}</p>
                    <p className="text-[11px] text-emerald-700">{item.barcode}</p>
                  </div>
                  <button
                    onClick={() => removeFromQueue(item.id)}
                    className="text-emerald-600 hover:text-[var(--red)] shrink-0"
                  >
                    <X size={16} />
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
          <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
            <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
              <Package size={16} />
              كل العروض ({filteredItems.length})
            </h2>
          </div>
          <div className="overflow-x-auto">
            <div className="max-h-[650px] overflow-y-auto">
              <table className="w-full text-sm border-collapse">
                <thead className="sticky top-0 bg-[var(--navy)] text-white z-10">
                  <tr>
                    <th className="p-3 text-right font-bold border-2 border-white/20">الباركود</th>
                    <th className="p-3 text-right font-bold border-2 border-white/20">اسم المنتج</th>
                    <th className="p-3 text-right font-bold border-2 border-white/20">السعر السابق</th>
                    <th className="p-3 text-right font-bold border-2 border-white/20">سعر العرض</th>
                    <th className="p-3 text-center font-bold border-2 border-white/20">إجراءات</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredItems.length === 0 && (
                    <tr>
                      <td colSpan={5} className="p-6 text-center text-gray-400 text-sm">ما فيه نتائج مطابقة</td>
                    </tr>
                  )}
                  {filteredItems.map((item, i) => (
                    <tr key={item.id} className={`${i % 2 === 0 ? 'bg-white' : 'bg-[var(--navy)]/[0.03]'} hover:bg-[var(--yellow)]/10 transition-colors`}>
                      <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.barcode}</td>
                      <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.product_name}</td>
                      <td className="p-3 text-gray-500 font-bold line-through border-2 border-[var(--navy)]/10">{item.previous_price.toFixed(2)}</td>
                      <td className="p-3 text-[var(--red)] font-black border-2 border-[var(--navy)]/10">{item.offer_price.toFixed(2)}</td>
                      <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                        <div className="flex items-center justify-center gap-2">
                          <button
                            onClick={() => handleAction(item, 'download')}
                            disabled={printingId === item.id}
                            className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
                          >
                            <Download size={13} />
                            تحميل
                          </button>
                          <button
                            onClick={() => handleAction(item, 'print')}
                            disabled={printingId === item.id}
                            className="flex items-center gap-1.5 bg-[var(--navy)] hover:bg-[#0f1a4d] text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
                          >
                            <Printer size={13} />
                            طباعة
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
          </div>
        </div>
      </div>
    </div>
  )
}
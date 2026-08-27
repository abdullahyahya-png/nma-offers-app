'use client'

import { useEffect, useRef, useState } from 'react'
import jsPDF from 'jspdf'
import * as XLSX from 'xlsx'
import { supabase } from '../../lib/supabase'
import InstallPWAButton from '../components/InstallPWAButton'
import { Search, Printer, Download, Package, ListPlus, Layers, X, LayoutGrid, UploadCloud, ClipboardCheck } from 'lucide-react'

interface OfferItem {
  id: string
  barcode: string
  product_name: string
  previous_price: number
  offer_price: number
  is_active?: boolean
  is_makeup?: boolean
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
// دقة مخفّضة للتوليد الجماعي (عشرات/مئات الملصقات) — لسا واضحة جداً للطباعة (~300 نقطة/إنش)
const BULK_SCALE = 2

const POS = {
  offerPrice: { xFrac: 0.45, yFrac: 545 / REF_H, fontPx: 90 },
  prevPrice: { xFrac: 0.47, yFrac: 705 / REF_H, fontPx: 58 },
  name: { xFrac: 0.5, yFrac: 845 / REF_H, fontPx: 34, lineSpacingPx: 38 },
  barcode: { xFrac: 0.5, yFrac: 968 / REF_H, fontPx: 33 },
}

const NAVY = '#150971'
const RED = '#C00000'

function normalizeBarcode(raw: any): string {
  if (raw === null || raw === undefined) return ''
  let str = String(raw).trim().replace(/[\s,،]/g, '')
  if (!str) return ''
  if (/^\d+(?:\.\d+)?e[+-]?\d+$/i.test(str)) {
    const num = Number(str)
    if (Number.isFinite(num)) str = num.toFixed(0)
  }
  if (/^\d+\.0+$/.test(str)) str = str.replace(/\.0+$/, '')
  return str
}

function downloadItemsAsExcel(items: OfferItem[], filename: string) {
  const rows = items.map((item) => ({
    الباركود: item.barcode,
    'اسم المنتج': item.product_name,
    'السعر السابق': item.previous_price,
    'سعر العرض': item.offer_price,
  }))
  const worksheet = XLSX.utils.json_to_sheet(rows)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'العروض')
  XLSX.writeFile(workbook, `${filename}.xlsx`)
}

function itemToLabelData(item: OfferItem): LabelData {
  return {
    name: item.product_name,
    offerPriceText: item.offer_price.toFixed(2),
    prevPriceText: item.previous_price.toFixed(2),
    barcodeText: item.barcode,
  }
}

export default function PrintPage() {
  const [activeSection, setActiveSection] = useState<'general' | 'excel' | 'downloads' | 'queue' | 'periodicCheck' | 'makeup'>('general')
  const [allItems, setAllItems] = useState<OfferItem[]>([])
  const [searchText, setSearchText] = useState('')
  const [makeupSearchText, setMakeupSearchText] = useState('')
  const [status, setStatus] = useState('')
  const [bgReady, setBgReady] = useState(false)
  const [printingId, setPrintingId] = useState<string | null>(null)
  const [printQueue, setPrintQueue] = useState<OfferItem[]>([])
  const [generatingQueue, setGeneratingQueue] = useState(false)
  const [excelUploading, setExcelUploading] = useState(false)
  const [excelMatchedItems, setExcelMatchedItems] = useState<OfferItem[]>([])
  const [excelUnmatchedBarcodes, setExcelUnmatchedBarcodes] = useState<string[]>([])
  const [excelTotalRead, setExcelTotalRead] = useState<number | null>(null)
  const [excelGenerating, setExcelGenerating] = useState(false)
  const [downloadsGenerating, setDownloadsGenerating] = useState(false)
  const [readyDownloads, setReadyDownloads] = useState<{ url: string; filename: string }[]>([])
  const [periodicChecks, setPeriodicChecks] = useState<Record<string, boolean>>({})
  const [periodicUnavailable, setPeriodicUnavailable] = useState<Record<string, boolean>>({})
  const [bulkJob, setBulkJob] = useState<{
    items: OfferItem[]
    baseFilename: string
    totalChunks: number
    currentChunk: number
  } | null>(null)
  const [generatingChunk, setGeneratingChunk] = useState(false)
  const bgImageRef = useRef<HTMLImageElement | null>(null)

  const queueIds = new Set(printQueue.map((i) => i.id))

  const addToQueue = (item: OfferItem) => {
    if (queueIds.has(item.id)) return
    setPrintQueue((prev) => [...prev, item])
  }

  const removeFromQueue = (id: string) => {
    setPrintQueue((prev) => prev.filter((i) => i.id !== id))
  }

  const clearQueue = () => setPrintQueue([])

  const handleExcelFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return

    setExcelUploading(true)
    setExcelMatchedItems([])
    setExcelUnmatchedBarcodes([])
    setExcelTotalRead(null)

    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        const data = event.target?.result
        if (!data) throw new Error('تعذر قراءة ملف الإكسل')

        const workbook = XLSX.read(data, { type: 'array' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        if (!sheet) throw new Error('لم يتم العثور على أي ورقة داخل الملف')

        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false })
        if (rows.length < 2) throw new Error('الملف لا يحتوي على بيانات كافية')

        const headers = rows[0].map((cell) =>
          String(cell ?? '').trim().toLowerCase().replace(/[^a-z0-9؀-ۿ]/g, '')
        )
        const barcodeColumnIndex = headers.findIndex((header) =>
          ['باركود', 'الباركود', 'barcode', 'ean', 'upc', 'sku', 'كودالصنف', 'كودالمنتج'].some((k) => header.includes(k))
        )

        let detectedColumn = barcodeColumnIndex
        if (detectedColumn === -1) {
          const maxColumns = Math.max(...rows.slice(0, 50).map((row) => row.length), 0)
          let bestColumn = 0
          let bestScore = -1
          for (let col = 0; col < maxColumns; col++) {
            let score = 0
            for (const row of rows.slice(1, 101)) {
              const value = normalizeBarcode(row[col])
              if (value && /^\d{6,}$/.test(value)) score++
            }
            if (score > bestScore) {
              bestScore = score
              bestColumn = col
            }
          }
          detectedColumn = bestColumn
        }

        const barcodes = Array.from(
          new Set(rows.slice(1).map((row) => normalizeBarcode(row[detectedColumn])).filter(Boolean))
        )
        if (barcodes.length === 0) throw new Error('لم يتم العثور على أي باركود صالح داخل الملف')

        const itemMap = new Map<string, OfferItem>()
        for (const item of allItems) {
          const normalized = normalizeBarcode(item.barcode)
          if (normalized && !itemMap.has(normalized)) itemMap.set(normalized, item)
        }

        const matched: OfferItem[] = []
        const notFound: string[] = []
        for (const barcode of barcodes) {
          const item = itemMap.get(barcode)
          if (item) matched.push(item)
          else notFound.push(barcode)
        }

        setExcelTotalRead(barcodes.length)
        setExcelMatchedItems(matched)
        setExcelUnmatchedBarcodes(notFound)

        const columnName = barcodeColumnIndex !== -1 && rows[0][detectedColumn] ? ` من عمود "${rows[0][detectedColumn]}"` : ''
        setStatus(
          `تم قراءة ${barcodes.length} باركود${columnName} — ${matched.length} له عرض حالياً` +
          (notFound.length > 0 ? ` — ${notFound.length} غير مطابق` : '')
        )
      } catch (err: any) {
        setExcelMatchedItems([])
        setExcelUnmatchedBarcodes([])
        setExcelTotalRead(null)
        setStatus(`تعذر قراءة الملف: ${err?.message || 'خطأ غير معروف'}`)
      } finally {
        setExcelUploading(false)
        e.target.value = ''
      }
    }
    reader.onerror = () => {
      setExcelUploading(false)
      e.target.value = ''
      setStatus('حدث خطأ أثناء قراءة ملف الإكسل')
    }
    reader.readAsArrayBuffer(file)
  }

  const handleAddExcelResultsToQueue = () => {
    if (excelMatchedNonMakeup.length === 0) return
    setPrintQueue((prev) => {
      const existingIds = new Set(prev.map((i) => i.id))
      const newOnes = excelMatchedNonMakeup.filter((m) => !existingIds.has(m.id))
      return [...prev, ...newOnes]
    })
    setStatus(`تمت إضافة ${excelMatchedNonMakeup.length} منتج لقائمة الطباعة`)
  }

  const handleDownloadAllExcel = () => {
    if (allItems.length === 0) {
      setStatus('لا توجد عروض حالياً')
      return
    }
    downloadItemsAsExcel(allItems, 'كل_العروض')
    setStatus(`تم تحميل ملف إكسل فيه ${allItems.length} منتج`)
  }

  const filteredItems = searchText.trim().length >= 1
    ? allItems.filter((item) => item.barcode.includes(searchText.trim()) || item.product_name.includes(searchText.trim()))
    : allItems

  // منتجات المكياج تُستثنى من الطباعة/التحميل الجماعي العام (نفس البراند، درجات مختلفة — ملصق واحد يكفيها)
  const makeupItems = allItems.filter((item) => item.is_makeup)
  const nonMakeupItems = allItems.filter((item) => !item.is_makeup)
  const filteredMakeupItems = makeupSearchText.trim().length >= 1
    ? makeupItems.filter((item) => item.barcode.includes(makeupSearchText.trim()) || item.product_name.includes(makeupSearchText.trim()))
    : makeupItems

  // نطبّق نفس الاستثناء على نتائج مطابقة ملف الإكسل — الطباعة/التحميل الجماعي هنا يستبعد المكياج
  const excelMatchedNonMakeup = excelMatchedItems.filter((item) => !item.is_makeup)
  const excelMatchedMakeupCount = excelMatchedItems.length - excelMatchedNonMakeup.length

  useEffect(() => {
    if (!status) return
    const timer = setTimeout(() => setStatus(''), 4000)
    return () => clearTimeout(timer)
  }, [status])

  // البانر الأخضر "الملف جاهز" يقفل تلقائياً لو المستخدم نسى يضغط عليه أو يسكره
  useEffect(() => {
    if (readyDownloads.length === 0) return
    const timer = setTimeout(() => setReadyDownloads([]), 60000)
    return () => clearTimeout(timer)
  }, [readyDownloads])

  // التشييك الدوري محفوظ محلياً بهذا الجهاز بس (ما يشتركه أي جهاز ثاني)
  useEffect(() => {
    const savedChecks = localStorage.getItem('print_periodic_checks')
    if (savedChecks) {
      try {
        setPeriodicChecks(JSON.parse(savedChecks))
      } catch {
        localStorage.removeItem('print_periodic_checks')
      }
    }
    const savedUnavailable = localStorage.getItem('print_periodic_unavailable')
    if (savedUnavailable) {
      try {
        setPeriodicUnavailable(JSON.parse(savedUnavailable))
      } catch {
        localStorage.removeItem('print_periodic_unavailable')
      }
    }
  }, [])

  useEffect(() => {
    localStorage.setItem('print_periodic_checks', JSON.stringify(periodicChecks))
  }, [periodicChecks])

  useEffect(() => {
    localStorage.setItem('print_periodic_unavailable', JSON.stringify(periodicUnavailable))
  }, [periodicUnavailable])

  const togglePeriodicCheck = (barcode: string) => {
    setPeriodicChecks((prev) => ({ ...prev, [barcode]: !prev[barcode] }))
  }

  const togglePeriodicUnavailable = (barcode: string) => {
    setPeriodicUnavailable((prev) => ({ ...prev, [barcode]: !prev[barcode] }))
  }

  const handleStartNewPeriodicRound = () => {
    const confirmed = window.confirm('بتبدأ دورة تشييك جديدة — كل العلامات ترجع فاضية من جديد على هذا الجهاز. متأكد؟')
    if (!confirmed) return
    setPeriodicChecks({})
  }

  // نستثني "غير المتوفر بفرعي" من العدّاد — النسبة تحسب بس على المنتجات المتوفرة فعلاً
  const periodicRelevantItems = allItems.filter((item) => !periodicUnavailable[item.barcode])
  const periodicCheckedCount = periodicRelevantItems.filter((item) => periodicChecks[item.barcode]).length

  const fetchOffers = async () => {
    const { data } = await supabase.from('offer_items').select('*').order('created_at', { ascending: false })
    if (data) setAllItems(data.filter((i) => i.is_active !== false))
  }

  useEffect(() => {
    fetchOffers()

    const { data: bgData } = supabase.storage.from('label-assets').getPublicUrl('label-bg.png')
    const img = new Image()
    img.crossOrigin = 'anonymous'
    img.src = `${bgData.publicUrl}?t=${Date.now()}`
    img.onload = () => {
      bgImageRef.current = img
      setBgReady(true)
    }
    img.onerror = () => setBgReady(false)

    const channel = supabase
      .channel('print-general-realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'offer_items' }, () => fetchOffers())
      .subscribe()

    return () => {
      supabase.removeChannel(channel)
    }
  }, [])

  // ========================================
  // توليد الملصقات — بنفس الطريقة تماماً بدون فرق بين "مفرد" و"جماعي"
  // (نفس السطور، نفس الترتيب، الفرق الوحيد إنه بحلقة تكرار)
  // ========================================
  const renderLabelCanvas = async (
    data: LabelData,
    scale: number = SCALE,
    bgOverride?: HTMLImageElement | HTMLCanvasElement | null
  ): Promise<HTMLCanvasElement> => {
    const canvas = document.createElement('canvas')
    canvas.width = REF_W * scale
    canvas.height = REF_H * scale
    const ctx = canvas.getContext('2d')!
    ctx.textBaseline = 'middle'

    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)

    const bg = bgOverride !== undefined ? bgOverride : bgImageRef.current
    if (bg) ctx.drawImage(bg, 0, 0, canvas.width, canvas.height)

    ctx.textAlign = 'center'
    ctx.direction = 'rtl' as any

    ctx.font = `900 ${POS.offerPrice.fontPx * scale}px Tajawal`
    ctx.fillStyle = NAVY
    ctx.fillText(data.offerPriceText, canvas.width * POS.offerPrice.xFrac, canvas.height * POS.offerPrice.yFrac)

    ctx.font = `700 ${POS.prevPrice.fontPx * scale}px Tajawal`
    ctx.fillStyle = RED
    const prevX = canvas.width * POS.prevPrice.xFrac
    const prevY = canvas.height * POS.prevPrice.yFrac
    ctx.fillText(data.prevPriceText, prevX, prevY)
    const prevWidth = ctx.measureText(data.prevPriceText).width
    ctx.beginPath()
    ctx.strokeStyle = RED
    ctx.lineWidth = 3 * scale
    ctx.moveTo(prevX - prevWidth / 2, prevY)
    ctx.lineTo(prevX + prevWidth / 2, prevY)
    ctx.stroke()

    ctx.font = `700 ${POS.name.fontPx * scale}px Tajawal`
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
    const lineSpacing = POS.name.lineSpacingPx * scale
    const startY = canvas.height * POS.name.yFrac - ((nameLines.length - 1) * lineSpacing) / 2
    nameLines.forEach((line, i) => {
      ctx.fillText(line, canvas.width * POS.name.xFrac, startY + i * lineSpacing)
    })

    ctx.font = `700 ${POS.barcode.fontPx * scale}px Tajawal`
    ctx.fillStyle = RED
    ctx.fillText(data.barcodeText, canvas.width * POS.barcode.xFrac, canvas.height * POS.barcode.yFrac)

    return canvas
  }

  // نشتق خلفية مصغّرة من نفس الخلفية الأصلية اللي نجحت فعلاً من أول تحميل للصفحة
  // (بدل ما نستدعي مكتبة قراءة PDF من جديد كل مرة — تصغير كانفاس عادي، عملية موثوقة 100%)
  const getScaledBackground = (scale: number): HTMLCanvasElement | HTMLImageElement | null => {
    if (!bgImageRef.current) return null
    if (scale === SCALE) return bgImageRef.current
    const canvas = document.createElement('canvas')
    canvas.width = REF_W * scale
    canvas.height = REF_H * scale
    const ctx = canvas.getContext('2d')!
    ctx.drawImage(bgImageRef.current, 0, 0, canvas.width, canvas.height)
    return canvas
  }

  // دالة توليد وحيدة يستخدمها كل شي (مفرد، متعدد، الكل) — بدون أي تفريع منطق
  // ما فيه ضغط PDF، ما فيه JPEG، ما فيه تمرير مقاس مخصص لـ addPage — أبسط شكل ممكن لـ jsPDF
  const generatePdf = async (
    items: OfferItem[],
    scale: number,
    onProgress?: (msg: string) => void
  ): Promise<jsPDF> => {
    const bg = getScaledBackground(scale)

    // فحص فوري: نتأكد الخلفية نفسها ما تسبب مشكلة "Tainted Canvas" (خطأ أمني معروف
    // يصير لما صورة بمصدر خارجي تلوّث الكانفاس، ويصير أي تصدير بعدها يفشل بصمت)
    if (bg) {
      const testCanvas = document.createElement('canvas')
      testCanvas.width = 10
      testCanvas.height = 10
      const testCtx = testCanvas.getContext('2d')!
      testCtx.drawImage(bg, 0, 0, 10, 10)
      try {
        testCanvas.toDataURL()
      } catch (taintErr: any) {
        throw new Error(
          `مشكلة تلوّث الكانفاس (Tainted Canvas) — الخلفية جاية من مصدر يمنع تصديرها: ${taintErr?.message || taintErr}`
        )
      }
    }

    const doc = new jsPDF({ unit: 'pt', format: [296.28, 496.2] })

    for (let i = 0; i < items.length; i++) {
      if (onProgress) {
        onProgress(`جاري توليد الملصق ${i + 1} من ${items.length}...`)
        await new Promise((r) => setTimeout(r, 0))
      }
      let canvas: HTMLCanvasElement
      try {
        canvas = await renderLabelCanvas(itemToLabelData(items[i]), scale, bg)
      } catch (renderErr: any) {
        throw new Error(`فشل رسم الملصق رقم ${i + 1} (${items[i].barcode}): ${renderErr?.message || renderErr}`)
      }
      if (i > 0) doc.addPage()
      try {
        doc.addImage(canvas, 'PNG', 0, 0, 296.28, 496.2)
      } catch (embedErr: any) {
        throw new Error(`فشل تضمين الملصق رقم ${i + 1} (${items[i].barcode}) بالملف: ${embedErr?.message || embedErr}`)
      }
    }

    return doc
  }

  const BULK_CHUNK_SIZE = 100

  // يولّد جزء واحد بس بكل مرة (بدل كل الأجزاء دفعة وحدة) — يمنع تراكم الذاكرة وتجمد المتصفح
  const generateBulkChunk = async (
    items: OfferItem[],
    baseFilename: string,
    chunkIndex: number,
    totalChunks: number
  ) => {
    setGeneratingChunk(true)
    // نحرر ذاكرة الملف السابق فوراً قبل توليد الجزء الجديد
    readyDownloads.forEach((d) => {
      try {
        URL.revokeObjectURL(d.url)
      } catch {}
    })
    setReadyDownloads([])
    try {
      await document.fonts.load('900 90px Tajawal')
      await document.fonts.load('700 58px Tajawal')
      await document.fonts.load('700 34px Tajawal')

      const chunkItems = items.slice(chunkIndex * BULK_CHUNK_SIZE, (chunkIndex + 1) * BULK_CHUNK_SIZE)
      const doc = await generatePdf(chunkItems, BULK_SCALE, (msg) =>
        setStatus(`(جزء ${chunkIndex + 1} من ${totalChunks}) ${msg}`)
      )
      const partSuffix = totalChunks > 1 ? `_جزء${chunkIndex + 1}من${totalChunks}` : ''
      const url = doc.output('bloburl') as unknown as string
      setReadyDownloads([{ url, filename: `${baseFilename}${partSuffix}.pdf` }])
      setBulkJob({ items, baseFilename, totalChunks, currentChunk: chunkIndex })
      setStatus(`جهّزنا الجزء ${chunkIndex + 1} من ${totalChunks} (${chunkItems.length} ملصق) — اضغط زر التحميل تحت`)
    } catch (err: any) {
      setStatus(`صار خطأ: ${err?.message || 'غير معروف'}`)
    } finally {
      setGeneratingChunk(false)
    }
  }

  const startBulkDownloadJob = async (items: OfferItem[], baseFilename: string) => {
    if (items.length === 0) {
      setStatus('لا توجد عروض لتحميلها')
      return
    }
    const totalChunks = Math.ceil(items.length / BULK_CHUNK_SIZE)
    await generateBulkChunk(items, baseFilename, 0, totalChunks)
  }

  const handleGenerateNextBulkChunk = () => {
    if (!bulkJob) return
    const nextIndex = bulkJob.currentChunk + 1
    if (nextIndex >= bulkJob.totalChunks) return
    generateBulkChunk(bulkJob.items, bulkJob.baseFilename, nextIndex, bulkJob.totalChunks)
  }

  const handleDownloadAllLabels = async () => {
    if (nonMakeupItems.length === 0) {
      setStatus('لا توجد عروض حالياً (غير المكياج)')
      return
    }
    setDownloadsGenerating(true)
    await startBulkDownloadJob(nonMakeupItems, 'ملصقات_كل_المنتجات')
    setDownloadsGenerating(false)
  }

  const handleDownloadMakeupLabels = async () => {
    setDownloadsGenerating(true)
    await startBulkDownloadJob(makeupItems, 'ملصقات_المكياج')
    setDownloadsGenerating(false)
  }

  const handleExcelAction = async (mode: 'download' | 'print') => {
    if (excelMatchedNonMakeup.length === 0) {
      setStatus('ما فيه منتجات مطابقة (غير المكياج) لطباعتها — رفع ملف أول')
      return
    }
    if (mode === 'download') {
      setExcelGenerating(true)
      await startBulkDownloadJob(excelMatchedNonMakeup, 'ملصقات_الملف_المرفوع')
      setExcelGenerating(false)
      return
    }
    const printWindow = window.open('', '_blank')
    setExcelGenerating(true)
    try {
      await document.fonts.load('900 90px Tajawal')
      await document.fonts.load('700 58px Tajawal')
      await document.fonts.load('700 34px Tajawal')

      const doc = await generatePdf(excelMatchedNonMakeup, BULK_SCALE, setStatus)
      doc.autoPrint()
      const blobUrl = doc.output('bloburl')
      if (printWindow) {
        printWindow.location.href = blobUrl as unknown as string
        setStatus('تم فتح نافذة الطباعة')
      } else {
        setStatus('المتصفح منع فتح نافذة الطباعة — اسمح بالنوافذ المنبثقة وحاول من جديد')
      }
    } catch (err: any) {
      if (printWindow) printWindow.close()
      setStatus(`صار خطأ: ${err?.message || 'غير معروف'}`)
    } finally {
      setExcelGenerating(false)
    }
  }

  const handleQueueAction = async (mode: 'download' | 'print') => {
    if (printQueue.length === 0) {
      setStatus('قائمة الطباعة فاضية — أضف منتجات أول')
      return
    }
    if (mode === 'download') {
      setGeneratingQueue(true)
      await startBulkDownloadJob(printQueue, 'ملصقات_مختارة')
      setGeneratingQueue(false)
      return
    }
    const printWindow = window.open('', '_blank')
    setGeneratingQueue(true)
    try {
      await document.fonts.load('900 90px Tajawal')
      await document.fonts.load('700 58px Tajawal')
      await document.fonts.load('700 34px Tajawal')

      const doc = await generatePdf(printQueue, BULK_SCALE, setStatus)
      doc.autoPrint()
      const blobUrl = doc.output('bloburl')
      if (printWindow) {
        printWindow.location.href = blobUrl as unknown as string
        setStatus(`تم توليد ${printQueue.length} ملصق وفتح نافذة الطباعة`)
      } else {
        setStatus('المتصفح منع فتح نافذة الطباعة — اسمح بالنوافذ المنبثقة لهذا الموقع وحاول من جديد')
      }
    } catch (err: any) {
      if (printWindow) printWindow.close()
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
    const printWindow = mode === 'print' ? window.open('', '_blank') : null
    setPrintingId(item.id)
    setStatus('جاري تجهيز الملصق بأعلى جودة، لحظات...')
    try {
      await document.fonts.load('900 90px Tajawal')
      await document.fonts.load('700 58px Tajawal')
      await document.fonts.load('700 34px Tajawal')

      const canvas = await renderLabelCanvas(itemToLabelData(item))
      const doc = new jsPDF({ unit: 'pt', format: [296.28, 496.2] })
      doc.addImage(canvas, 'PNG', 0, 0, 296.28, 496.2)

      if (mode === 'print') {
        doc.autoPrint()
        const blobUrl = doc.output('bloburl')
        if (printWindow) {
          printWindow.location.href = blobUrl as unknown as string
          setStatus('تم فتح نافذة الطباعة')
        } else {
          setStatus('المتصفح منع فتح نافذة الطباعة — اسمح بالنوافذ المنبثقة وحاول من جديد')
        }
      } else {
        doc.save(`ملصق_${item.barcode}.pdf`)
        setStatus('تم تحميل الملصق بنجاح')
      }
    } catch (err: any) {
      if (printWindow) printWindow.close()
      setStatus(`صار خطأ: ${err?.message || 'غير معروف'}`)
    } finally {
      setPrintingId(null)
    }
  }

  return (
    <div className="min-h-screen w-full bg-[var(--background)] overflow-x-hidden">
      <InstallPWAButton />
      <header className="bg-white border-b-4 border-[var(--navy)]">
        <div className="w-full max-w-6xl mx-auto px-4 py-4 flex items-center gap-3">
          <img src="/logo.png" alt="شعار العروض" className="w-12 h-12 object-contain shrink-0" />
          <div>
            <p className="text-[var(--red)] text-sm font-bold">واجهة الطباعة السريعة ومتابعة العروض</p>
            <h1 className="text-[var(--navy)] text-base font-black">عروض عامة</h1>
          </div>
        </div>
      </header>

      <div className="w-full max-w-6xl mx-auto px-3 sm:px-4 py-4 sm:py-6 flex flex-col md:flex-row gap-4 md:gap-6 items-start">
        <aside className="w-full md:w-64 md:shrink-0 md:sticky md:top-6">
          <div className="w-full grid grid-cols-2 md:flex md:flex-col gap-1.5 bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-2 shadow-sm">
            <button
              onClick={() => setActiveSection('general')}
              className={`w-full flex items-center justify-between gap-1.5 px-2.5 py-2.5 rounded-lg text-xs sm:text-sm font-bold transition-colors ${
                activeSection === 'general' ? 'bg-[var(--navy)]/10 text-[var(--navy)]' : 'text-gray-600 hover:bg-gray-50'
              }`}
            >
              <span className="flex items-center gap-2">
                <LayoutGrid size={15} />
                كل العروض
              </span>
              <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-gray-200 text-gray-600">{allItems.length}</span>
            </button>
            <button
              onClick={() => setActiveSection('excel')}
              className={`w-full flex items-center gap-1.5 px-2.5 py-2.5 rounded-lg text-xs sm:text-sm font-bold transition-colors ${
                activeSection === 'excel' ? 'bg-[var(--navy)]/10 text-[var(--navy)]' : 'text-gray-600 hover:bg-gray-50'
              }`}
            >
              <UploadCloud size={15} />
              رفع ملف إكسل
            </button>
            <button
              onClick={() => setActiveSection('downloads')}
              className={`w-full flex items-center gap-1.5 px-2.5 py-2.5 rounded-lg text-xs sm:text-sm font-bold transition-colors ${
                activeSection === 'downloads' ? 'bg-[var(--navy)]/10 text-[var(--navy)]' : 'text-gray-600 hover:bg-gray-50'
              }`}
            >
              <Download size={15} />
              التنزيلات
            </button>
            <button
              onClick={() => setActiveSection('queue')}
              className={`w-full flex items-center justify-between gap-1.5 px-2.5 py-2.5 rounded-lg text-xs sm:text-sm font-bold transition-colors ${
                activeSection === 'queue' ? 'bg-[var(--navy)]/10 text-[var(--navy)]' : 'text-gray-600 hover:bg-gray-50'
              }`}
            >
              <span className="flex items-center gap-2">
                <Printer size={15} />
                قائمة الطباعة
              </span>
              {printQueue.length > 0 && (
                <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-emerald-500 text-white">{printQueue.length}</span>
              )}
            </button>
            <button
              onClick={() => setActiveSection('periodicCheck')}
              className={`w-full flex items-center justify-between gap-1.5 px-2.5 py-2.5 rounded-lg text-xs sm:text-sm font-bold transition-colors ${
                activeSection === 'periodicCheck' ? 'bg-[var(--navy)]/10 text-[var(--navy)]' : 'text-gray-600 hover:bg-gray-50'
              }`}
            >
              <span className="flex items-center gap-2">
                <ClipboardCheck size={15} />
                التشييك الدوري
              </span>
              {periodicCheckedCount > 0 && (
                <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-emerald-500 text-white">{periodicCheckedCount}/{allItems.length}</span>
              )}
            </button>
            <button
              onClick={() => setActiveSection('makeup')}
              className={`w-full flex items-center justify-between gap-1.5 px-2.5 py-2.5 rounded-lg text-xs sm:text-sm font-bold transition-colors ${
                activeSection === 'makeup' ? 'bg-[var(--navy)]/10 text-[var(--navy)]' : 'text-gray-600 hover:bg-gray-50'
              }`}
            >
              <span className="flex items-center gap-2">
                <Layers size={15} />
                عروض المكياج
              </span>
              <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-pink-200 text-pink-700">{makeupItems.length}</span>
            </button>
          </div>
        </aside>

        <div className="flex-1 min-w-0 w-full space-y-4">
          {status && (
            <div className="p-3.5 bg-[var(--yellow)]/15 border-2 border-[var(--yellow)]/40 rounded-lg text-sm text-[#8a6300] font-bold">
              {status}
            </div>
          )}

          {readyDownloads.length > 0 && (
            <div className="p-4 bg-emerald-50 border-2 border-emerald-400 rounded-lg space-y-2 relative">
              <button
                onClick={() => {
                  setReadyDownloads([])
                  setBulkJob(null)
                }}
                className="absolute top-3 left-3 text-emerald-700 hover:text-emerald-900"
                title="إغلاق"
              >
                <X size={16} />
              </button>
              <p className="text-sm text-emerald-800 font-bold pl-6">
                ✅ {bulkJob ? `الجزء ${bulkJob.currentChunk + 1} من ${bulkJob.totalChunks} جاهز` : 'الملف جاهز'} — اضغط للتحميل
              </p>
              <div className="flex flex-wrap items-center gap-2">
                {readyDownloads.map((d, idx) => (
                  <a
                    key={idx}
                    href={d.url}
                    download={d.filename}
                    className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-sm font-bold px-4 py-2 rounded-lg transition-colors"
                  >
                    <Download size={14} />
                    تحميل هذا الجزء
                  </a>
                ))}
                {bulkJob && bulkJob.currentChunk + 1 < bulkJob.totalChunks && (
                  <button
                    onClick={handleGenerateNextBulkChunk}
                    disabled={generatingChunk}
                    className="flex items-center gap-1.5 bg-white border-2 border-emerald-500 hover:bg-emerald-100 text-emerald-700 text-sm font-bold px-4 py-2 rounded-lg transition-colors disabled:opacity-50"
                  >
                    {generatingChunk ? 'جاري التوليد...' : `توليد الجزء التالي (${bulkJob.currentChunk + 2} من ${bulkJob.totalChunks})`}
                  </button>
                )}
              </div>
            </div>
          )}

          {activeSection === 'excel' && (
            <div className="space-y-4">
              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-5 shadow-sm max-w-xl">
                <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2 mb-1">
                  <UploadCloud size={16} />
                  رفع ملف إكسل باركودات
                </h2>
                <p className="text-xs text-gray-500 font-medium mb-4">
                  ارفع أي ملف Excel فيه عمود باركودات (أي فرع أو أي مصدر)، ونطابقها تلقائياً مع العروض الحالية
                </p>
                <label className="flex items-center justify-center gap-2 border-2 border-dashed border-[var(--navy)]/25 rounded-xl p-6 cursor-pointer hover:border-[var(--navy)] hover:bg-[var(--navy)]/5 transition-colors">
                  <UploadCloud size={20} className="text-[var(--navy)]" />
                  <span className="text-[var(--navy)] font-bold text-sm">
                    {excelUploading ? 'جاري القراءة...' : 'اختر ملف Excel'}
                  </span>
                  <input type="file" accept=".xlsx,.xls" onChange={handleExcelFileUpload} className="hidden" disabled={excelUploading} />
                </label>
              </div>

              {excelTotalRead !== null && (
                <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm max-w-xl">
                  <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
                    <h3 className="text-sm font-black text-[var(--navy)]">نتيجة المطابقة</h3>
                  </div>
                  <div className="p-4 grid grid-cols-2 gap-3 text-sm">
                    <div className="bg-gray-50 rounded-lg p-3">
                      <p className="text-gray-500 text-xs font-bold">إجمالي الباركودات بالملف</p>
                      <p className="text-[var(--navy)] font-black text-lg">{excelTotalRead}</p>
                    </div>
                    <div className="bg-emerald-50 rounded-lg p-3">
                      <p className="text-emerald-700 text-xs font-bold">له عرض حالياً</p>
                      <p className="text-emerald-700 font-black text-lg">{excelMatchedItems.length}</p>
                    </div>
                    <div className="bg-[var(--red)]/5 rounded-lg p-3 col-span-2">
                      <p className="text-[var(--red)] text-xs font-bold">ما تم لقاه بالعروض الحالية</p>
                      <p className="text-[var(--red)] font-black text-lg">{excelUnmatchedBarcodes.length}</p>
                    </div>
                    {excelMatchedMakeupCount > 0 && (
                      <div className="bg-pink-50 rounded-lg p-3 col-span-2">
                        <p className="text-pink-700 text-xs font-bold">
                          منهم {excelMatchedMakeupCount} فئة مكياج — مستثناة تلقائياً من الطباعة/التحميل هنا (روح تبويب "عروض المكياج" لطباعتها)
                        </p>
                      </div>
                    )}
                  </div>
                  {excelMatchedNonMakeup.length > 0 && (
                    <div className="p-4 pt-0 flex flex-wrap items-center gap-2">
                      <button
                        onClick={() => handleExcelAction('download')}
                        disabled={excelGenerating}
                        className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-4 py-2 rounded-lg transition-colors disabled:opacity-50"
                      >
                        <Download size={13} />
                        تحميل ملصقات المطابقة
                      </button>
                      <button
                        onClick={() => handleExcelAction('print')}
                        disabled={excelGenerating}
                        className="flex items-center gap-1.5 bg-[var(--navy)] hover:bg-[#0f1a4d] text-white text-xs font-bold px-4 py-2 rounded-lg transition-colors disabled:opacity-50"
                      >
                        <Printer size={13} />
                        {excelGenerating ? 'جاري التجهيز...' : 'طباعة مباشرة'}
                      </button>
                      <button
                        onClick={handleAddExcelResultsToQueue}
                        className="flex items-center gap-1.5 bg-white border-2 border-emerald-300 hover:bg-emerald-50 text-emerald-700 text-xs font-bold px-4 py-2 rounded-lg transition-colors"
                      >
                        <ListPlus size={13} />
                        أضف لقائمة الطباعة
                      </button>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {activeSection === 'downloads' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-5 shadow-sm max-w-xl">
              <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2 mb-1">
                <Download size={16} />
                التنزيلات
              </h2>
              <p className="text-xs text-gray-500 font-medium mb-4">
                تحميل بيانات أو ملصقات كل المنتجات دفعة وحدة
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <button
                  onClick={handleDownloadAllExcel}
                  className="flex flex-col items-center gap-2 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] rounded-xl p-5 transition-colors"
                >
                  <Package size={24} />
                  <span className="text-sm font-bold">تحميل إكسل كل المنتجات</span>
                  <span className="text-[11px] text-gray-500">({allItems.length} منتج بكل تفاصيلها)</span>
                </button>
                <button
                  onClick={handleDownloadAllLabels}
                  disabled={downloadsGenerating}
                  className="flex flex-col items-center gap-2 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] rounded-xl p-5 transition-colors disabled:opacity-50"
                >
                  <Printer size={24} />
                  <span className="text-sm font-bold">
                    {downloadsGenerating ? 'جاري التوليد...' : 'تحميل ملصقات كل المنتجات'}
                  </span>
                  <span className="text-[11px] text-gray-500">({nonMakeupItems.length} منتج، بدون فئة المكياج)</span>
                </button>
              </div>
            </div>
          )}

          {activeSection === 'queue' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
              <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5 flex flex-wrap items-center justify-between gap-3">
                <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
                  <Printer size={16} />
                  قائمة الطباعة ({printQueue.length})
                </h2>
                {printQueue.length > 0 && (
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => handleQueueAction('download')}
                      disabled={generatingQueue}
                      className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
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
                    <button onClick={clearQueue} className="text-[var(--red)] hover:underline text-xs font-bold">
                      تفريغ
                    </button>
                  </div>
                )}
              </div>
              <div className="divide-y-2 divide-[var(--navy)]/10 max-h-[600px] overflow-y-auto">
                {printQueue.length === 0 && (
                  <p className="p-6 text-center text-gray-400 text-sm">
                    قائمة الطباعة فاضية — أضف منتجات من "كل العروض" أو "إضافة عدة باركودات"
                  </p>
                )}
                {printQueue.map((item) => (
                  <div key={item.id} className="flex items-center justify-between gap-3 p-3.5">
                    <div className="min-w-0">
                      <p className="text-sm font-bold text-[var(--navy)] truncate">{item.product_name}</p>
                      <p className="text-xs text-gray-500">{item.barcode}</p>
                    </div>
                    <button onClick={() => removeFromQueue(item.id)} className="text-gray-400 hover:text-[var(--red)] shrink-0">
                      <X size={16} />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {activeSection === 'periodicCheck' && (
            <div className="space-y-4">
              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-5 shadow-sm">
                <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
                  <div>
                    <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
                      <ClipboardCheck size={16} />
                      التشييك الدوري
                    </h2>
                    <p className="text-xs text-gray-500 font-medium mt-1">
                      أكّد إن ملصق كل منتج موجود وصحيح
                    </p>
                  </div>
                  <button
                    onClick={handleStartNewPeriodicRound}
                    className="flex items-center gap-1.5 bg-white border-2 border-[var(--red)]/20 hover:bg-[var(--red)]/5 text-[var(--red)] text-xs font-bold px-4 py-2 rounded-lg transition-colors"
                  >
                    بدء دورة تشييك جديدة
                  </button>
                </div>
                <div className="flex items-center gap-3">
                  <div className="flex-1 bg-gray-100 rounded-full h-3 overflow-hidden">
                    <div
                      className="bg-emerald-500 h-full transition-all"
                      style={{ width: `${periodicRelevantItems.length > 0 ? (periodicCheckedCount / periodicRelevantItems.length) * 100 : 0}%` }}
                    />
                  </div>
                  <span className="text-xs font-black text-[var(--navy)] shrink-0">
                    {periodicCheckedCount} من {periodicRelevantItems.length}
                  </span>
                </div>
              </div>

              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                <div className="divide-y-2 divide-[var(--navy)]/10 max-h-[600px] overflow-y-auto">
                  {allItems.length === 0 && (
                    <p className="p-6 text-center text-gray-400 text-sm">ما فيه عروض حالياً</p>
                  )}
                  {allItems.map((item) => {
                    const isChecked = !!periodicChecks[item.barcode]
                    const isUnavailable = !!periodicUnavailable[item.barcode]
                    return (
                      <div
                        key={item.id}
                        className={`flex flex-col sm:flex-row sm:items-center justify-between gap-2 sm:gap-3 p-3 ${
                          isUnavailable ? 'bg-gray-50 opacity-60' : isChecked ? 'bg-emerald-50' : 'bg-white'
                        }`}
                      >
                        <div className="flex items-start gap-3 min-w-0">
                          <input
                            type="checkbox"
                            checked={isChecked}
                            disabled={isUnavailable}
                            onChange={() => togglePeriodicCheck(item.barcode)}
                            className="w-5 h-5 mt-0.5 cursor-pointer accent-emerald-600 shrink-0 disabled:cursor-not-allowed"
                          />
                          <div className="min-w-0">
                            <p className={`text-sm font-bold break-words ${isUnavailable ? 'text-gray-400 line-through' : 'text-[var(--navy)]'}`}>
                              {item.product_name}
                            </p>
                            <p className="text-[11px] text-gray-500">{item.barcode}</p>
                          </div>
                        </div>
                        <button
                          onClick={() => togglePeriodicUnavailable(item.barcode)}
                          className={`text-[11px] font-bold px-2.5 py-1.5 rounded-lg shrink-0 self-start sm:self-auto transition-colors ${
                            isUnavailable
                              ? 'bg-gray-200 text-gray-600 hover:bg-gray-300'
                              : 'bg-white border border-gray-300 text-gray-500 hover:bg-gray-100'
                          }`}
                        >
                          {isUnavailable ? 'مو متوفر' : 'غير متوفر بفرعي'}
                        </button>
                      </div>
                    )
                  })}
                </div>
              </div>
            </div>
          )}

          {activeSection === 'makeup' && (
            <div className="space-y-4">
              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-4 shadow-sm">
                <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
                  <div className="relative flex-1 min-w-[200px]">
                    <Search size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input
                      value={makeupSearchText}
                      onChange={(e) => setMakeupSearchText(e.target.value)}
                      placeholder="ابحث بالاسم أو الباركود"
                      className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-3 pr-9 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                    />
                  </div>
                  {makeupItems.length > 0 && (
                    <button
                      onClick={handleDownloadMakeupLabels}
                      disabled={downloadsGenerating}
                      className="flex items-center gap-1.5 bg-pink-600 hover:bg-pink-500 text-white text-xs font-bold px-4 py-2.5 rounded-lg transition-colors disabled:opacity-50 shrink-0"
                    >
                      <Printer size={13} />
                      {downloadsGenerating ? 'جاري التوليد...' : `تحميل كل ملصقات المكياج (${makeupItems.length})`}
                    </button>
                  )}
                </div>
                <p className="text-[11px] text-gray-400 font-medium">
                  منتجات المكياج مستثناة من الطباعة الجماعية العامة — تطبعها من هنا لحالها
                </p>
              </div>

              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
                  <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
                    <Layers size={16} />
                    عروض المكياج ({filteredMakeupItems.length})
                  </h2>
                </div>

                <div className="md:hidden divide-y-2 divide-[var(--navy)]/10 max-h-[600px] overflow-y-auto">
                  {filteredMakeupItems.length === 0 && (
                    <p className="p-6 text-center text-gray-400 text-sm">ما فيه نتائج</p>
                  )}
                  {filteredMakeupItems.map((item) => (
                    <div key={item.id} className="p-3.5 space-y-2">
                      <div>
                        <p className="text-sm font-bold text-[var(--navy)]">{item.product_name}</p>
                        <p className="text-xs text-gray-500 mt-0.5">
                          {item.barcode} ·{' '}
                          <span className="line-through text-gray-400">{item.previous_price.toFixed(2)}</span>{' '}
                          <span className="text-[var(--red)] font-bold">{item.offer_price.toFixed(2)}</span>
                        </p>
                      </div>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => addToQueue(item)}
                          disabled={queueIds.has(item.id)}
                          className="flex items-center gap-1.5 bg-white border-2 border-emerald-300 hover:bg-emerald-50 text-emerald-700 text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-40"
                        >
                          <ListPlus size={13} />
                          {queueIds.has(item.id) ? 'مضاف' : 'أضف'}
                        </button>
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
                    </div>
                  ))}
                </div>

                <div className="hidden md:block overflow-x-auto">
                  <div className="max-h-[600px] overflow-y-auto">
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
                        {filteredMakeupItems.length === 0 && (
                          <tr>
                            <td colSpan={5} className="p-6 text-center text-gray-400 text-sm">ما فيه نتائج</td>
                          </tr>
                        )}
                        {filteredMakeupItems.map((item, i) => (
                          <tr key={item.id} className={`${i % 2 === 0 ? 'bg-white' : 'bg-pink-50/40'} hover:bg-pink-100/50 transition-colors`}>
                            <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.barcode}</td>
                            <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.product_name}</td>
                            <td className="p-3 text-gray-500 font-bold line-through border-2 border-[var(--navy)]/10">{item.previous_price.toFixed(2)}</td>
                            <td className="p-3 text-[var(--red)] font-black border-2 border-[var(--navy)]/10">{item.offer_price.toFixed(2)}</td>
                            <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                              <div className="flex items-center justify-center gap-2">
                                <button
                                  onClick={() => addToQueue(item)}
                                  disabled={queueIds.has(item.id)}
                                  title="أضف لقائمة الطباعة"
                                  className="flex items-center gap-1.5 bg-white border-2 border-emerald-300 hover:bg-emerald-50 text-emerald-700 text-xs font-bold px-2.5 py-2 rounded-lg transition-colors disabled:opacity-40"
                                >
                                  <ListPlus size={13} />
                                </button>
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
          )}

          {activeSection === 'general' && (
            <>
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

              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
                  <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
                    <Package size={16} />
                    كل العروض ({filteredItems.length})
                  </h2>
                </div>

                <div className="md:hidden divide-y-2 divide-[var(--navy)]/10 max-h-[600px] overflow-y-auto">
                  {filteredItems.length === 0 && (
                    <p className="p-6 text-center text-gray-400 text-sm">ما فيه نتائج مطابقة</p>
                  )}
                  {filteredItems.map((item) => (
                    <div key={item.id} className="p-3.5 space-y-2">
                      <div>
                        <p className="text-sm font-bold text-[var(--navy)]">{item.product_name}</p>
                        <p className="text-xs text-gray-500 mt-0.5">
                          {item.barcode} ·{' '}
                          <span className="line-through text-gray-400">{item.previous_price.toFixed(2)}</span>{' '}
                          <span className="text-[var(--red)] font-bold">{item.offer_price.toFixed(2)}</span>
                        </p>
                      </div>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => addToQueue(item)}
                          disabled={queueIds.has(item.id)}
                          className="flex items-center gap-1.5 bg-white border-2 border-emerald-300 hover:bg-emerald-50 text-emerald-700 text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-40"
                        >
                          <ListPlus size={13} />
                          {queueIds.has(item.id) ? 'مضاف' : 'أضف'}
                        </button>
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
                    </div>
                  ))}
                </div>

                <div className="hidden md:block overflow-x-auto">
                  <div className="max-h-[600px] overflow-y-auto">
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
                                  onClick={() => addToQueue(item)}
                                  disabled={queueIds.has(item.id)}
                                  title="أضف لقائمة الطباعة"
                                  className="flex items-center gap-1.5 bg-white border-2 border-emerald-300 hover:bg-emerald-50 text-emerald-700 text-xs font-bold px-2.5 py-2 rounded-lg transition-colors disabled:opacity-40"
                                >
                                  <ListPlus size={13} />
                                </button>
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
            </>
          )}
        </div>
      </div>
    </div>
  )
}
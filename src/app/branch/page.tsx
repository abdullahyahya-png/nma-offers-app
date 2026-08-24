'use client'

import { useEffect, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import { supabase } from '../../lib/supabase'
import { renderPdfToCanvas } from '../../lib/pdfBackground'
import InstallPWAButton from '../components/InstallPWAButton'
import {
  Search, UploadCloud, Download, Printer, Package, Sparkles, Bell,
  XCircle, CheckCircle2, LayoutGrid, Layers, ChevronDown, ChevronUp,
  MessageCircle, Send, Store, LogOut, ClipboardCheck,
} from 'lucide-react'

interface OfferItem {
  id: string
  barcode: string
  product_name: string
  previous_price: number
  offer_price: number
  is_active?: boolean
  batch_id?: string | null
  cancelled_batch_id?: string | null
}

interface Branch {
  id: string
  name: string
}

interface OfferBatch {
  id: string
  label: string
  batch_type: 'new' | 'cancel'
  created_at: string
}

interface LabelData {
  name: string
  offerPriceText: string
  prevPriceText: string
  barcodeText: string
}

interface Message {
  id: string
  sender_role: 'admin' | 'branch'
  sender_branch_id: string | null
  target_branch_id: string | null
  body: string
  created_at: string
  is_read?: boolean
}

type SectionId = 'general' | 'branch' | 'updates' | 'cancelled' | 'custom' | 'messages' | 'upload' | 'search' | 'audit'

const REF_W = 618
const REF_H = 1034
const DEFAULT_SCALE = 4

const POS = {
  offerPrice: { xFrac: 0.45, yFrac: 545 / REF_H, fontPx: 90 },
  prevPrice: { xFrac: 0.47, yFrac: 705 / REF_H, fontPx: 58 },
  name: { xFrac: 0.5, yFrac: 845 / REF_H, fontPx: 34, lineSpacingPx: 38 },
  barcode: { xFrac: 0.5, yFrac: 968 / REF_H, fontPx: 33 },
}

const NAVY = '#150971'
const RED = '#C00000'

const SECTIONS: { id: SectionId; label: string; icon: any }[] = [
  { id: 'general', label: 'عروض عامة', icon: LayoutGrid },
  { id: 'branch', label: 'عروض الفرع', icon: Package },
  { id: 'updates', label: 'آخر التحديثات', icon: Bell },
  { id: 'cancelled', label: 'عروض ملغاة', icon: XCircle },
  { id: 'audit', label: 'تدقيق الملصقات', icon: ClipboardCheck },
  { id: 'custom', label: 'ملصق مخصص', icon: Layers },
  { id: 'messages', label: 'التواصل مع الإدارة', icon: MessageCircle },
  { id: 'upload', label: 'رفع ملف الفرع', icon: UploadCloud },
  { id: 'search', label: 'بحث بالباركود', icon: Search },
]

function formatDate(dateStr: string) {
  const d = new Date(dateStr)
  return d.toLocaleDateString('ar-SA', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })
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

function downloadBarcodeListAsExcel(barcodes: string[], filename: string) {
  const rows = barcodes.map((b) => ({ الباركود: b }))
  const worksheet = XLSX.utils.json_to_sheet(rows)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'باركودات')
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

export default function BranchPage() {
  const [branches, setBranches] = useState<Branch[]>([])
  const [selectedBranch, setSelectedBranch] = useState('')
  const [loggedInBranchName, setLoggedInBranchName] = useState('')
  const [sessionChecked, setSessionChecked] = useState(false)
  const [loginBranchName, setLoginBranchName] = useState('')
  const [loginPassword, setLoginPassword] = useState('')
  const [loginError, setLoginError] = useState('')
  const [loggingIn, setLoggingIn] = useState(false)
  const [allItems, setAllItems] = useState<OfferItem[]>([])
  const [cancelledItems, setCancelledItems] = useState<OfferItem[]>([])
  // قائمة باركودات الفرع الدائمة — تُطابق أي منتج بأي تحديث (قديم أو جديد) بنفس الباركود
  const [branchBarcodes, setBranchBarcodes] = useState<Set<string>>(new Set())
  const [labelChecks, setLabelChecks] = useState<Record<string, { is_checked: boolean; checked_at: string | null }>>({})
  const [auditLoading, setAuditLoading] = useState(false)
  const [activeSection, setActiveSection] = useState<SectionId>('branch')
  const [status, setStatus] = useState('')

  // يتحقق لو فيه جلسة دخول محفوظة بالمتصفح
  useEffect(() => {
    const saved = localStorage.getItem('branch_session')
    if (saved) {
      try {
        const parsed = JSON.parse(saved)
        if (parsed?.branchId && parsed?.branchName) {
          setSelectedBranch(parsed.branchId)
          setLoggedInBranchName(parsed.branchName)
        }
      } catch {
        localStorage.removeItem('branch_session')
      }
    }
    setSessionChecked(true)
  }, [])

  const handleLogin = async () => {
    if (!loginBranchName || !loginPassword.trim()) {
      setLoginError('اختر فرعك واكتب كلمة السر')
      return
    }
    setLoggingIn(true)
    setLoginError('')
    const { data, error } = await supabase.rpc('verify_branch_login', {
      p_branch_name: loginBranchName,
      p_password: loginPassword,
    })
    setLoggingIn(false)

    if (error || !data || data.length === 0) {
      setLoginError('كلمة السر غلط، أو الفرع ما له كلمة سر محفوظة بعد')
      return
    }

    const { branch_id, branch_name } = data[0]
    localStorage.setItem('branch_session', JSON.stringify({ branchId: branch_id, branchName: branch_name }))
    setSelectedBranch(branch_id)
    setLoggedInBranchName(branch_name)
    setLoginBranchName('')
    setLoginPassword('')
  }

  const handleLogout = () => {
    localStorage.removeItem('branch_session')
    setSelectedBranch('')
    setLoggedInBranchName('')
    setActiveSection('branch')
  }

  useEffect(() => {
    if (!status) return
    const timer = setTimeout(() => setStatus(''), 5000)
    return () => clearTimeout(timer)
  }, [status])
  const [generating, setGenerating] = useState(false)
  const bgImageRef = useRef<HTMLCanvasElement | null>(null)

  const [searchBarcode, setSearchBarcode] = useState('')

  const [recentBatches, setRecentBatches] = useState<OfferBatch[]>([])
  const [recentCancelBatches, setRecentCancelBatches] = useState<OfferBatch[]>([])
  const [confirmedBatchIds, setConfirmedBatchIds] = useState<string[]>([])
  const [confirmedRemovalIds, setConfirmedRemovalIds] = useState<string[]>([])
  const [previewUpdateBatchId, setPreviewUpdateBatchId] = useState<string | null>(null)
  const [previewCancelBatchId, setPreviewCancelBatchId] = useState<string | null>(null)
  const [expandedBranchBatchId, setExpandedBranchBatchId] = useState<string | null>(null)
  const [branchViewMode, setBranchViewMode] = useState<'grouped' | 'all'>('grouped')

  const [customName, setCustomName] = useState('')
  const [customPrevPrice, setCustomPrevPrice] = useState('')
  const [customOfferPrice, setCustomOfferPrice] = useState('')
  const [customNote, setCustomNote] = useState('')

  const [messages, setMessages] = useState<Message[]>([])
  const [newMessageText, setNewMessageText] = useState('')
  const [sendingMessage, setSendingMessage] = useState(false)

  // عروض الفرع المطابقة فعلياً (بالباركود) — تشمل أي تحديث قديم أو جديد بنفس الباركود
  const matchedItems = allItems.filter((item) => branchBarcodes.has(item.barcode))

  const auditCheckedCount = matchedItems.filter((item) => labelChecks[item.barcode]?.is_checked).length

  // أول ما يفتح تبويب التدقيق، نجهز صف لكل منتج بالفرع (لو مو موجود) عشان النسبة تحسب صح
  useEffect(() => {
    if (activeSection === 'audit' && selectedBranch && matchedItems.length > 0) {
      const missing = matchedItems.filter((item) => !(item.barcode in labelChecks)).map((item) => item.barcode)
      if (missing.length > 0) {
        ensureAuditRows(selectedBranch, missing)
      }
    }
  }, [activeSection, selectedBranch, matchedItems.length])

  // لوحة حالة الفرع: عدد التحديثات والإلغاءات اللي لسا ما اتأكدت + الرسائل
  const pendingUpdatesCount = recentBatches.filter((b) => !confirmedBatchIds.includes(b.id)).length
  const pendingRemovalsCount = recentCancelBatches.filter((b) => !confirmedRemovalIds.includes(b.id)).length
  const unreadMessagesCount = messages.filter((m) => m.sender_role === 'admin' && !m.is_read).length

  // تجميع منتجات الفرع حسب كل تحديث (لعرضها منفصلة بتبويب "عروض الفرع")
  const branchBatchGroups = recentBatches
    .map((batch) => ({
      batch,
      items: allItems.filter((i) => i.batch_id === batch.id && branchBarcodes.has(i.barcode)),
    }))
    .filter((g) => g.items.length > 0)

  const noBatchMatchedItems = allItems.filter((i) => !i.batch_id && branchBarcodes.has(i.barcode))

  const [searchScope, setSearchScope] = useState<'mine' | 'all'>('mine')
  const [uploadReport, setUploadReport] = useState<{
    totalRead: number
    uniqueCount: number
    duplicatesCount: number
    foundCount: number
    notFoundCount: number
    notFoundBarcodes: string[]
  } | null>(null)


  const searchResults = searchBarcode.trim().length >= 2
    ? (searchScope === 'mine' ? matchedItems : allItems)
        .filter((item) => item.barcode.includes(searchBarcode.trim()))
        .slice(0, 20)
    : []

  const fetchBranchBarcodes = async (branchId: string) => {
    const { data } = await supabase
      .from('branch_barcodes')
      .select('barcode')
      .eq('branch_id', branchId)
    setBranchBarcodes(new Set(data ? data.map((d) => d.barcode) : []))
  }

  const fetchLabelChecks = async (branchId: string) => {
    const { data } = await supabase
      .from('label_checks')
      .select('barcode, is_checked, checked_at')
      .eq('branch_id', branchId)
    const map: Record<string, { is_checked: boolean; checked_at: string | null }> = {}
    data?.forEach((row) => {
      map[row.barcode] = { is_checked: row.is_checked, checked_at: row.checked_at }
    })
    setLabelChecks(map)
  }

  // يجهز صف تدقيق لكل منتج بالفرع أول ما يفتح التبويب (لو مو موجود أصلاً)، عشان النسبة تحسب صح من البداية
  const ensureAuditRows = async (branchId: string, barcodes: string[]) => {
    if (barcodes.length === 0) return
    const rows = barcodes.map((barcode) => ({ branch_id: branchId, barcode }))
    await supabase.from('label_checks').upsert(rows, { onConflict: 'branch_id,barcode', ignoreDuplicates: true })
    await fetchLabelChecks(branchId)
  }

  const handleToggleLabelCheck = async (barcode: string, currentlyChecked: boolean) => {
    if (!selectedBranch) return
    const newValue = !currentlyChecked
    setLabelChecks((prev) => ({
      ...prev,
      [barcode]: { is_checked: newValue, checked_at: newValue ? new Date().toISOString() : null },
    }))
    await supabase
      .from('label_checks')
      .upsert(
        [{ branch_id: selectedBranch, barcode, is_checked: newValue, checked_at: newValue ? new Date().toISOString() : null }],
        { onConflict: 'branch_id,barcode' }
      )
  }

  const handleStartNewAuditRound = async () => {
    if (!selectedBranch) return
    const confirmed = window.confirm('بتبدأ دورة تدقيق جديدة — كل الملصقات ترجع "غير مدقّقة" من جديد. متأكد؟')
    if (!confirmed) return
    setAuditLoading(true)
    await supabase
      .from('label_checks')
      .update({ is_checked: false, checked_at: null })
      .eq('branch_id', selectedBranch)
    await fetchLabelChecks(selectedBranch)
    setAuditLoading(false)
    setStatus('تم بدء دورة تدقيق جديدة')
    logActivity('بدء دورة تدقيق ملصقات جديدة')
  }

  const fetchMessages = async (branchId: string) => {
    const { data } = await supabase
      .from('messages')
      .select('*')
      .or(`target_branch_id.eq.${branchId},target_branch_id.is.null`)
      .order('created_at', { ascending: false })
      .limit(30)
    if (data) setMessages(data)
  }

  const fetchConfirmations = async (branchId: string) => {
    const { data: confirmData } = await supabase
      .from('branch_batch_confirmations')
      .select('batch_id')
      .eq('branch_id', branchId)
    setConfirmedBatchIds(confirmData ? confirmData.map((d) => d.batch_id) : [])

    const { data: removalData } = await supabase
      .from('branch_cancel_confirmations')
      .select('batch_id')
      .eq('branch_id', branchId)
    setConfirmedRemovalIds(removalData ? removalData.map((d) => d.batch_id) : [])
  }

  const fetchOffersData = async () => {
    const { data: branchesData } = await supabase.from('branches').select('id, name').order('name')
    const { data: itemsData } = await supabase.from('offer_items').select('*').order('created_at', { ascending: false })
    const { data: batchesData } = await supabase
      .from('offer_batches')
      .select('*')
      .eq('batch_type', 'new')
      .order('created_at', { ascending: false })
      .limit(10)
    const { data: cancelBatchesData } = await supabase
      .from('offer_batches')
      .select('*')
      .eq('batch_type', 'cancel')
      .order('created_at', { ascending: false })
      .limit(10)

    if (branchesData) setBranches(branchesData)
    if (itemsData) {
      const active = itemsData.filter((i) => i.is_active !== false)
      setAllItems(active)
      setCancelledItems(itemsData.filter((i) => i.is_active === false))
    }
    if (batchesData) setRecentBatches(batchesData)
    if (cancelBatchesData) setRecentCancelBatches(cancelBatchesData)
  }

  // يسجل أي عملية يسويها الفرع بسجل النشاط (يشوفه الأدمن بتبويب "سجل النشاط")
  const logActivity = async (action: string, referenceId?: string, details?: string) => {
    if (!selectedBranch) return
    await supabase.from('activity_logs').insert([{
      branch_id: selectedBranch,
      actor_role: 'branch',
      action,
      reference_id: referenceId || null,
      details: details || null,
    }])
  }

  useEffect(() => {
    fetchOffersData()

    const { data: bgData } = supabase.storage.from('label-assets').getPublicUrl('label-bg.pdf')
    renderPdfToCanvas(`${bgData.publicUrl}?t=${Date.now()}`, REF_W * DEFAULT_SCALE, REF_H * DEFAULT_SCALE)
      .then((canvas) => {
        bgImageRef.current = canvas
      })
      .catch(() => {
        bgImageRef.current = null
      })

    // Realtime: أي تغيير على المنتجات أو التحديثات ينعكس فوراً بدون Refresh
    const channel = supabase
      .channel('branch-page-realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'offer_items' }, () => {
        fetchOffersData()
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'offer_batches' }, () => {
        fetchOffersData()
      })
      .subscribe()

    return () => {
      supabase.removeChannel(channel)
    }
  }, [])

  useEffect(() => {
    const load = async () => {
      if (!selectedBranch) {
        setConfirmedBatchIds([])
        setConfirmedRemovalIds([])
        setBranchBarcodes(new Set())
        setMessages([])
        setLabelChecks({})
        return
      }
      await fetchConfirmations(selectedBranch)
      await fetchBranchBarcodes(selectedBranch)
      await fetchMessages(selectedBranch)
      await fetchLabelChecks(selectedBranch)
    }
    load()

    if (!selectedBranch) return

    // Realtime: الرسائل الجديدة من الإدارة وتأكيدات الفروع تظهر لحظياً
    const channel = supabase
      .channel(`branch-${selectedBranch}-realtime`)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'messages' }, () => {
        fetchMessages(selectedBranch)
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'branch_batch_confirmations', filter: `branch_id=eq.${selectedBranch}` }, () => {
        fetchConfirmations(selectedBranch)
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'branch_cancel_confirmations', filter: `branch_id=eq.${selectedBranch}` }, () => {
        fetchConfirmations(selectedBranch)
      })
      .subscribe()

    return () => {
      supabase.removeChannel(channel)
    }
  }, [selectedBranch])

  // يعلّم رسائل الإدارة كمقروءة تلقائياً لما الفرع يفتح تبويب الرسائل
  useEffect(() => {
    if (activeSection !== 'messages' || !selectedBranch) return
    const unreadIds = messages.filter((m) => m.sender_role === 'admin' && !m.is_read).map((m) => m.id)
    if (unreadIds.length === 0) return
    supabase
      .from('messages')
      .update({ is_read: true, read_at: new Date().toISOString() })
      .in('id', unreadIds)
      .then(() => {
        setMessages((prev) => prev.map((m) => (unreadIds.includes(m.id) ? { ...m, is_read: true } : m)))
      })
  }, [activeSection, selectedBranch, messages])

  const handleConfirmBatch = async (batchId: string) => {
    if (!selectedBranch) {
      setStatus('اختر الفرع أول عشان تأكد استلام التحديث')
      return
    }
    const { error } = await supabase
      .from('branch_batch_confirmations')
      .upsert([{ branch_id: selectedBranch, batch_id: batchId }], { onConflict: 'branch_id,batch_id' })
    if (!error) {
      setConfirmedBatchIds((prev) => (prev.includes(batchId) ? prev : [...prev, batchId]))
      setStatus('تم تأكيد استلام التحديث')
      const batch = recentBatches.find((b) => b.id === batchId)
      logActivity('تأكيد تفعيل تحديث', batchId, batch?.label)
    }
  }

  const handleConfirmRemoval = async (batchId: string) => {
    if (!selectedBranch) {
      setStatus('اختر الفرع أول عشان تأكد إزالة الملصقات')
      return
    }
    const { error } = await supabase
      .from('branch_cancel_confirmations')
      .upsert([{ branch_id: selectedBranch, batch_id: batchId }], { onConflict: 'branch_id,batch_id' })
    if (!error) {
      setConfirmedRemovalIds((prev) => (prev.includes(batchId) ? prev : [...prev, batchId]))
      setStatus('تم تأكيد إزالة الملصقات لهذا الإلغاء')
      const batch = recentCancelBatches.find((b) => b.id === batchId)
      logActivity('تأكيد إزالة ملصقات', batchId, batch?.label)
    }
  }

  const handleDownloadBatch = (batch: OfferBatch, cancel = false) => {
    const batchItems = cancel
      ? cancelledItems.filter((item) => item.cancelled_batch_id === batch.id && branchBarcodes.has(item.barcode))
      : allItems.filter((item) => item.batch_id === batch.id && branchBarcodes.has(item.barcode))
    if (batchItems.length === 0) {
      setStatus('ما فيه منتجات تخص فرعك بهذا التحديث')
      return
    }
    downloadItemsAsExcel(batchItems, batch.label)
  }

  // تحميل كل منتجات التحديث بدون فلترة — عشان الفرع يراجعها ويحدد باركودات منتجاته منها
  const handleDownloadFullBatch = (batch: OfferBatch, cancel = false) => {
    const batchItems = cancel
      ? cancelledItems.filter((item) => item.cancelled_batch_id === batch.id)
      : allItems.filter((item) => item.batch_id === batch.id)
    if (batchItems.length === 0) {
      setStatus('ما فيه منتجات بهذا التحديث')
      return
    }
    downloadItemsAsExcel(batchItems, `${batch.label}_كامل`)
  }

  const handleDownloadTemplate = () => {
    if (allItems.length === 0) {
      setStatus('لا توجد عروض حالياً')
      return
    }
    downloadItemsAsExcel(allItems, 'قالب_باركودات_العروض')
  }

  // رفع ملف باركودات الفرع — يخزن الباركودات الخام (يشتغل مع أي تحديث حالي أو مستقبلي بنفس الباركود)
  // يطبّع الباركود: يحل مشكلة الصيغة العلمية (6.28E+12) وفواصل الآلاف ويشيل المسافات
  const normalizeBarcode = (raw: any): string => {
    if (raw === null || raw === undefined) return ''
    let str = String(raw).trim()
    // لو Excel حوّله لصيغة علمية زي 6.28101E+12
    if (/^[\d.]+E\+?\d+$/i.test(str)) {
      const num = Number(str)
      if (!isNaN(num)) str = num.toFixed(0)
    }
    // يشيل أي مسافات أو فواصل زايدة
    str = str.replace(/[\s,]/g, '')
    return str
  }

  const handleUploadBranchFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!selectedBranch) {
      setStatus('اختر الفرع أول')
      return
    }
    const file = e.target.files?.[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = async (event) => {
      const data = event.target?.result
      // raw:false يحافظ على تنسيق الخلية الأصلي (يمنع تحويل الباركودات الطويلة لصيغة علمية)
      const workbook = XLSX.read(data, { type: 'binary' })
      const sheetName = workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]
      const rows: any[] = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false })

      const rawBarcodes = rows
        .slice(1)
        .filter((row) => row.length >= 1 && row[0])
        .map((row) => normalizeBarcode(row[0]))
        .filter(Boolean)

      const totalRead = rawBarcodes.length
      const uniqueSet = new Set(rawBarcodes)
      const uploadedBarcodes = Array.from(uniqueSet)
      const duplicatesCount = totalRead - uploadedBarcodes.length

      if (uploadedBarcodes.length === 0) {
        setStatus('الملف لا يحتوي على أي باركود صالح')
        return
      }

      // تقرير: كم باركود له عرض موجود بالنظام حالياً، وكم مو موجود
      const allBarcodesInSystem = new Set(allItems.map((i) => i.barcode))
      const foundInSystem = uploadedBarcodes.filter((b) => allBarcodesInSystem.has(b))
      const notFoundInSystem = uploadedBarcodes.filter((b) => !allBarcodesInSystem.has(b))

      const rowsToInsert = uploadedBarcodes.map((barcode) => ({
        branch_id: selectedBranch,
        barcode,
      }))

      const { error } = await supabase
        .from('branch_barcodes')
        .upsert(rowsToInsert, { onConflict: 'branch_id,barcode', ignoreDuplicates: true })

      if (error) {
        setStatus(`خطأ: ${error.message}`)
        return
      }

      await fetchBranchBarcodes(selectedBranch)
      setUploadReport({
        totalRead,
        uniqueCount: uploadedBarcodes.length,
        duplicatesCount,
        foundCount: foundInSystem.length,
        notFoundCount: notFoundInSystem.length,
        notFoundBarcodes: notFoundInSystem,
      })
      setStatus(`تمت إضافة ${uploadedBarcodes.length} باركود لقائمة فرعك — تشتغل تلقائياً مع أي تحديث حالي أو مستقبلي`)
      logActivity('رفع باركودات', undefined, `${uploadedBarcodes.length} باركود`)
    }
    reader.readAsBinaryString(file)
  }

  const renderLabelCanvas = async (data: LabelData, scale: number = DEFAULT_SCALE): Promise<HTMLCanvasElement> => {
    const canvas = document.createElement('canvas')
    canvas.width = REF_W * scale
    canvas.height = REF_H * scale
    const ctx = canvas.getContext('2d')!
    ctx.textBaseline = 'middle'

    if (bgImageRef.current) {
      ctx.drawImage(bgImageRef.current, 0, 0, canvas.width, canvas.height)
    }

    ctx.textAlign = 'center'
    ctx.direction = 'rtl' as any

    const offerFontPx = POS.offerPrice.fontPx * scale
    ctx.font = `900 ${offerFontPx}px Tajawal`
    ctx.fillStyle = NAVY
    ctx.fillText(data.offerPriceText, canvas.width * POS.offerPrice.xFrac, canvas.height * POS.offerPrice.yFrac)

    const prevFontPx = POS.prevPrice.fontPx * scale
    ctx.font = `700 ${prevFontPx}px Tajawal`
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

    const nameFontPx = POS.name.fontPx * scale
    ctx.font = `700 ${nameFontPx}px Tajawal`
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

    const barcodeFontPx = POS.barcode.fontPx * scale
    ctx.font = `700 ${barcodeFontPx}px Tajawal`
    ctx.fillStyle = RED
    ctx.fillText(data.barcodeText, canvas.width * POS.barcode.xFrac, canvas.height * POS.barcode.yFrac)

    return canvas
  }

  // يولّد ملف PDF واحد بجودة كاملة — يمرر الـ Canvas مباشرة لـ jsPDF (أسرع بكثير من تحويله لنص Base64)
  // ويفعّل ضغط PDF الداخلي (compress) اللي يقلل حجم الملف النهائي بدون أي فقدان بالجودة المرئية
  const generateLabelsSingleFile = async (items: OfferItem[], filename: string, mode: 'download' | 'print' = 'download') => {
    const validItems = items.filter(
      (item) => item.product_name && item.barcode && !isNaN(item.previous_price) && !isNaN(item.offer_price)
    )
    const skippedCount = items.length - validItems.length

    if (validItems.length === 0) {
      setStatus('كل المنتجات ناقصة بيانات (اسم/سعر) — تأكد من صحة البيانات بجدول المنتجات')
      return
    }

    setGenerating(true)
    const scale = DEFAULT_SCALE

    try {
      await document.fonts.load('900 90px Tajawal')
      await document.fonts.load('700 58px Tajawal')
      await document.fonts.load('700 34px Tajawal')

      const doc = new jsPDF({ unit: 'pt', format: [296.28, 496.2], compress: true })

      for (let i = 0; i < validItems.length; i++) {
        setStatus(`جاري توليد الملصق ${i + 1} من ${validItems.length}...`)
        await new Promise((r) => setTimeout(r, 0))
        const canvas = await renderLabelCanvas(itemToLabelData(validItems[i]), scale)
        if (i > 0) doc.addPage([296.28, 496.2])
        // تمرير الـ canvas مباشرة (بدل base64) + ضغط FAST يخلي الحفظ أسرع بكثير مع نفس الدقة
        doc.addImage(canvas, 'PNG', 0, 0, 296.28, 496.2, undefined, 'FAST')
      }

      if (mode === 'print') {
        setStatus('جاري فتح نافذة الطباعة...')
        doc.autoPrint()
        const blobUrl = doc.output('bloburl')
        window.open(blobUrl as unknown as string, '_blank')
      } else {
        setStatus(`جاري حفظ الملف النهائي (${validItems.length} ملصق)...`)
        await new Promise((r) => setTimeout(r, 0))
        doc.save(`${filename}.pdf`)
      }

      setStatus(
        `تم توليد ${validItems.length} ملصق بنجاح ${mode === 'print' ? 'وفتح نافذة الطباعة' : 'بملف واحد'}` +
        (skippedCount > 0 ? ` — تم تخطي ${skippedCount} منتج بياناته ناقصة` : '')
      )
    } catch (err: any) {
      setStatus(`صار خطأ أثناء التوليد: ${err?.message || 'خطأ غير معروف'}`)
    } finally {
      setGenerating(false)
    }
  }

  const handleGenerateAllLabels = async (mode: 'download' | 'print' = 'download') => {
    if (matchedItems.length === 0) {
      setStatus('لا توجد منتجات متوفرة بالفرع')
      return
    }
    if (!bgImageRef.current) {
      setStatus('جاري تحميل صورة الخلفية، حاول بعد ثانيتين')
      return
    }
    await generateLabelsSingleFile(matchedItems, 'ملصقات_العروض', mode)
  }

  // توليد ملصقات لتحديث واحد بس (بس المنتجات المطابقة لفرعك)
  const handleGenerateBatchLabels = async (items: OfferItem[], filename: string, mode: 'download' | 'print' = 'download') => {
    if (items.length === 0) {
      setStatus('لا توجد منتجات تخص فرعك بهذا التحديث')
      return
    }
    if (!bgImageRef.current) {
      setStatus('جاري تحميل صورة الخلفية، حاول بعد ثانيتين')
      return
    }
    await generateLabelsSingleFile(items, filename, mode)
  }

  const handlePrintSingle = async (item: OfferItem) => {
    if (!bgImageRef.current) {
      setStatus('جاري تحميل صورة الخلفية، حاول بعد ثانيتين')
      return
    }
    await document.fonts.load('900 90px Tajawal')
    await document.fonts.load('700 58px Tajawal')
    await document.fonts.load('700 34px Tajawal')

    const canvas = await renderLabelCanvas(itemToLabelData(item), DEFAULT_SCALE)
    const doc = new jsPDF({ unit: 'pt', format: [296.28, 496.2], compress: true })
    doc.addImage(canvas, 'PNG', 0, 0, 296.28, 496.2, undefined, 'FAST')
    doc.save(`ملصق_${item.barcode}.pdf`)
  }

  const handleGenerateCustomLabel = async () => {
    if (!customName.trim() || !customPrevPrice || !customOfferPrice) {
      setStatus('عبّي اسم المنتج والسعر السابق وسعر العرض أول')
      return
    }
    if (!bgImageRef.current) {
      setStatus('جاري تحميل صورة الخلفية، حاول بعد ثانيتين')
      return
    }

    const data: LabelData = {
      name: customName.trim(),
      offerPriceText: Number(customOfferPrice).toFixed(2),
      prevPriceText: Number(customPrevPrice).toFixed(2),
      barcodeText: customNote.trim() || 'عرض خاص',
    }

    await document.fonts.load('900 90px Tajawal')
    await document.fonts.load('700 58px Tajawal')
    await document.fonts.load('700 34px Tajawal')

    const canvas = await renderLabelCanvas(data)
    const doc = new jsPDF({ unit: 'pt', format: [296.28, 496.2], compress: true })
    doc.addImage(canvas, 'PNG', 0, 0, 296.28, 496.2, undefined, 'FAST')
    doc.save(`ملصق_مخصص_${customName.trim()}.pdf`)

    setStatus('تم توليد الملصق المخصص بنجاح')
  }

  const handleSendMessage = async () => {
    if (!selectedBranch) {
      setStatus('اختر الفرع أول عشان ترسل رسالة')
      return
    }
    if (!newMessageText.trim()) return

    setSendingMessage(true)
    const { error } = await supabase.from('messages').insert([{
      sender_role: 'branch',
      sender_branch_id: selectedBranch,
      target_branch_id: null,
      body: newMessageText.trim(),
    }])
    setSendingMessage(false)

    if (error) {
      setStatus(`خطأ بالإرسال: ${error.message}`)
    } else {
      setNewMessageText('')
      setStatus('تم إرسال رسالتك للإدارة')
      logActivity('إرسال رسالة للإدارة')
      fetchMessages(selectedBranch)
    }
  }

  if (!sessionChecked) {
    return <div className="min-h-screen bg-[var(--background)]" />
  }

  if (!selectedBranch) {
    return (
      <div className="min-h-screen bg-[var(--background)] flex items-center justify-center p-6">
        <InstallPWAButton />
        <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 shadow-sm p-8 w-full max-w-sm">
          <div className="flex flex-col items-center mb-6">
            <img src="/logo.png" alt="شعار العروض" className="w-16 h-16 object-contain mb-3" />
            <p className="text-[var(--red)] text-xs font-bold">بوابة الفروع</p>
            <h1 className="text-[var(--navy)] text-lg font-black">تسجيل الدخول</h1>
          </div>

          {loginError && (
            <div className="mb-4 p-3 bg-[var(--red)]/5 border-2 border-[var(--red)]/20 rounded-lg text-xs text-[var(--red)] font-bold text-center">
              {loginError}
            </div>
          )}

          <label className="block text-xs font-bold text-gray-500 mb-1">اختر فرعك</label>
          <select
            value={loginBranchName}
            onChange={(e) => setLoginBranchName(e.target.value)}
            className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-bold mb-3 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
          >
            <option value="">-- اختر الفرع --</option>
            {branches.map((b) => (
              <option key={b.id} value={b.name}>{b.name}</option>
            ))}
          </select>

          <label className="block text-xs font-bold text-gray-500 mb-1">كلمة السر</label>
          <input
            value={loginPassword}
            onChange={(e) => setLoginPassword(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && handleLogin()}
            type="password"
            placeholder="••••••••"
            className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium mb-5 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
          />

          <button
            onClick={handleLogin}
            disabled={loggingIn}
            className="w-full bg-[var(--navy)] hover:bg-[#0f1a4d] text-white font-black py-3 rounded-xl transition-colors disabled:opacity-50"
          >
            {loggingIn ? 'جاري الدخول...' : 'دخول'}
          </button>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-[var(--background)]">
      <InstallPWAButton />
      <header className="bg-white border-b-4 border-[var(--navy)]">
        <div className="max-w-7xl mx-auto px-6 md:px-10 py-4 flex items-center gap-4">
          <img src="/logo.png" alt="شعار العروض" className="w-16 h-16 md:w-20 md:h-20 object-contain shrink-0" />
          <div>
            <p className="text-[var(--red)] text-xs font-bold">بوابة الفروع</p>
            <h1 className="text-[var(--navy)] text-xl md:text-2xl font-black">عروض المنتجات</h1>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 md:px-10 py-6 md:py-8 flex flex-col md:flex-row gap-6 items-start">
        <aside className="w-full md:w-72 md:shrink-0 md:sticky md:top-8 space-y-4">
          <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-4 shadow-sm">
            <label className="block text-xs font-bold text-gray-400 mb-2">فرعك</label>
            <div className="flex items-center justify-between bg-[var(--navy)]/5 rounded-lg p-2.5 mb-2">
              <span className="text-sm font-black text-[var(--navy)]">{loggedInBranchName}</span>
              <Store size={16} className="text-[var(--navy)]" />
            </div>
            <button
              onClick={handleLogout}
              className="w-full flex items-center justify-center gap-1.5 bg-white border-2 border-[var(--red)]/20 hover:bg-[var(--red)]/5 text-[var(--red)] text-xs font-bold py-2 rounded-lg transition-colors"
            >
              <LogOut size={13} />
              تسجيل الخروج
            </button>
          </div>

          <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-4 shadow-sm">
            <button
              onClick={() => setActiveSection('branch')}
              className="w-full flex items-center justify-between bg-[var(--navy)] text-white px-4 py-3 rounded-xl text-sm font-bold mb-4"
            >
              <span className="flex items-center gap-2">
                <Package size={16} />
                عروض الفرع
              </span>
              <span className="bg-white/20 px-2 py-0.5 rounded-full text-xs">{matchedItems.length}</span>
            </button>

            <p className="text-xs font-bold text-gray-400 mb-2">الأقسام</p>
            <div className="space-y-1">
              {SECTIONS.map((s) => {
                const Icon = s.icon
                const active = activeSection === s.id
                const count =
                  s.id === 'general' ? allItems.length :
                  s.id === 'branch' ? matchedItems.length :
                  s.id === 'updates' ? recentBatches.length :
                  s.id === 'cancelled' ? cancelledItems.length :
                  s.id === 'messages' ? unreadMessagesCount : null
                return (
                  <button
                    key={s.id}
                    onClick={() => setActiveSection(s.id)}
                    className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-sm font-bold transition-colors ${
                      active ? 'bg-[var(--navy)]/10 text-[var(--navy)]' : 'text-gray-600 hover:bg-gray-50'
                    }`}
                  >
                    <span className="flex items-center gap-2">
                      <Icon size={15} />
                      {s.label}
                    </span>
                    {count !== null && (
                      <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${
                        s.id === 'cancelled' && cancelledItems.length > 0 ? 'bg-[var(--red)] text-white' : 'bg-gray-200 text-gray-600'
                      }`}>
                        {count}
                      </span>
                    )}
                  </button>
                )
              })}
            </div>
          </div>
        </aside>

        <div className="flex-1 min-w-0">
          {selectedBranch && (
            <div className="mb-6 bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-5 shadow-sm">
              <h2 className="text-sm font-black text-[var(--navy)] mb-3">
                {loggedInBranchName || 'فرعك'}
              </h2>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                <button
                  onClick={() => setActiveSection('branch')}
                  className="flex items-center gap-2 bg-emerald-50 hover:bg-emerald-100 rounded-xl p-3 text-right transition-colors"
                >
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 shrink-0" />
                  <span className="text-sm font-bold text-emerald-700">{matchedItems.length} عرض نشط</span>
                </button>
                <button
                  onClick={() => setActiveSection('updates')}
                  className="flex items-center gap-2 bg-[var(--yellow)]/15 hover:bg-[var(--yellow)]/25 rounded-xl p-3 text-right transition-colors"
                >
                  <span className="w-2.5 h-2.5 rounded-full bg-[var(--yellow)] shrink-0" />
                  <span className="text-sm font-bold text-[#8a6300]">{pendingUpdatesCount} تحديث يحتاج تأكيد</span>
                </button>
                <button
                  onClick={() => setActiveSection('cancelled')}
                  className="flex items-center gap-2 bg-[var(--red)]/5 hover:bg-[var(--red)]/10 rounded-xl p-3 text-right transition-colors"
                >
                  <span className="w-2.5 h-2.5 rounded-full bg-[var(--red)] shrink-0" />
                  <span className="text-sm font-bold text-[var(--red)]">{pendingRemovalsCount} إلغاء يحتاج إزالة</span>
                </button>
                <button
                  onClick={() => setActiveSection('messages')}
                  className="flex items-center gap-2 bg-[var(--navy)]/5 hover:bg-[var(--navy)]/10 rounded-xl p-3 text-right transition-colors"
                >
                  <span className="w-2.5 h-2.5 rounded-full bg-[var(--navy)] shrink-0" />
                  <span className="text-sm font-bold text-[var(--navy)]">{unreadMessagesCount} رسالة جديدة</span>
                </button>
              </div>
            </div>
          )}

          {status && (
            <div className="mb-6 p-3.5 bg-[var(--yellow)]/15 border-2 border-[var(--yellow)]/40 rounded-lg text-sm text-[#8a6300] font-bold">
              {status}
            </div>
          )}

          {activeSection === 'search' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm">
              <h2 className="font-black text-sm text-[var(--navy)] mb-1">البحث عن عرض بالباركود</h2>
              <p className="text-xs text-gray-500 font-medium mb-3">اكتب جزء من الباركود بس، والنتائج تظهر تلقائياً</p>

              <div className="flex items-center gap-2 mb-3">
                <button
                  onClick={() => setSearchScope('mine')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-colors ${
                    searchScope === 'mine' ? 'bg-[var(--navy)] text-white' : 'bg-white border-2 border-[var(--navy)]/15 text-[var(--navy)] hover:bg-[var(--navy)]/10'
                  }`}
                >
                  البحث في عروض فرعي
                </button>
                <button
                  onClick={() => setSearchScope('all')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-colors ${
                    searchScope === 'all' ? 'bg-[var(--navy)] text-white' : 'bg-white border-2 border-[var(--navy)]/15 text-[var(--navy)] hover:bg-[var(--navy)]/10'
                  }`}
                >
                  البحث في جميع العروض
                </button>
              </div>

              <div className="relative">
                <Search size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input
                  value={searchBarcode}
                  onChange={(e) => setSearchBarcode(e.target.value)}
                  placeholder="امسح أو اكتب جزء من الباركود"
                  className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-3 pr-9 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                  autoFocus
                />
              </div>

              <div className="mt-4 space-y-2 max-h-[500px] overflow-y-auto">
                {searchBarcode.trim().length >= 2 && searchResults.length === 0 && (
                  <div className="p-4 bg-[var(--red)]/5 border-2 border-[var(--red)]/30 rounded-xl">
                    <p className="text-sm text-[var(--red)] font-bold">ما فيه نتائج مطابقة</p>
                  </div>
                )}
                {searchResults.map((item) => (
                  <div key={item.id} className="p-4 bg-[var(--yellow)]/10 border-2 border-[var(--yellow)]/30 rounded-xl flex items-center justify-between gap-4">
                    <div className="min-w-0">
                      <p className="font-black text-[var(--navy)] truncate">{item.product_name}</p>
                      <p className="text-xs text-gray-500 mt-1">
                        باركود: {item.barcode} · سابق: <span className="line-through">{item.previous_price.toFixed(2)}</span> · عرض: <span className="text-[var(--red)] font-bold">{item.offer_price.toFixed(2)}</span>
                      </p>
                    </div>
                    <button
                      onClick={() => handlePrintSingle(item)}
                      className="bg-[var(--navy)] hover:bg-[#0f1a4d] text-white px-4 py-2 rounded-lg text-xs font-bold transition-colors flex items-center gap-1.5 shrink-0"
                    >
                      <Printer size={13} />
                      طباعة
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {activeSection === 'upload' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm max-w-2xl">
              <h2 className="font-black text-sm text-[var(--navy)] mb-1">ملف باركودات فرعك</h2>
              <p className="text-xs text-gray-600 font-medium mb-4">
                ارفع باركودات المنتجات الموجودة بفرعك (عمود وحد فيه الباركود). كل ملف ترفعه يُضاف على قائمة فرعك الدائمة —
                وتشتغل تلقائياً مع أي تحديث حالي أو جديد ينزل مستقبلاً بنفس الباركود، بدون ما تحتاج ترفع من جديد.
              </p>
              <label className="flex items-center justify-center gap-2 border-2 border-dashed border-[var(--navy)]/25 rounded-lg p-6 cursor-pointer hover:border-[var(--navy)] hover:bg-[var(--navy)]/5 transition-colors">
                <UploadCloud size={18} className="text-[var(--navy)]" />
                <span className="text-[var(--navy)] text-sm font-bold">ارفع ملف فرعك</span>
                <input type="file" accept=".xlsx,.xls" onChange={handleUploadBranchFile} className="hidden" />
              </label>
              <p className="text-[11px] text-gray-400 font-medium mt-3">
                محتاج قائمة كل الباركودات أول؟ حمّلها من تبويب "عروض عامة"، أو من "معاينة" أي تحديث بتبويب "آخر التحديثات" وحدد باركودات منتجاتك منه.
              </p>

              {uploadReport && (
                <div className="mt-5 border-2 border-[var(--navy)]/15 rounded-xl overflow-hidden">
                  <div className="p-3 bg-[var(--navy)]/5 border-b-2 border-[var(--navy)]/10">
                    <h3 className="text-sm font-black text-[var(--navy)]">تقرير آخر ملف رفعته</h3>
                  </div>
                  <div className="p-4 grid grid-cols-2 gap-3 text-sm">
                    <div className="bg-gray-50 rounded-lg p-3">
                      <p className="text-gray-500 text-xs font-bold">إجمالي الصفوف اللي قرأناها</p>
                      <p className="text-[var(--navy)] font-black text-lg">{uploadReport.totalRead}</p>
                    </div>
                    <div className="bg-gray-50 rounded-lg p-3">
                      <p className="text-gray-500 text-xs font-bold">باركودات فريدة (بعد شيل التكرار)</p>
                      <p className="text-[var(--navy)] font-black text-lg">{uploadReport.uniqueCount}</p>
                    </div>
                    <div className="bg-[var(--yellow)]/10 rounded-lg p-3">
                      <p className="text-[#8a6300] text-xs font-bold">مكرر داخل نفس الملف</p>
                      <p className="text-[#8a6300] font-black text-lg">{uploadReport.duplicatesCount}</p>
                    </div>
                    <div className="bg-emerald-50 rounded-lg p-3">
                      <p className="text-emerald-700 text-xs font-bold">له عرض موجود بالنظام حالياً</p>
                      <p className="text-emerald-700 font-black text-lg">{uploadReport.foundCount}</p>
                    </div>
                    <div className="bg-[var(--red)]/5 rounded-lg p-3 col-span-2">
                      <p className="text-[var(--red)] text-xs font-bold">باركود مو موجود بالنظام (ما إله عرض حالياً)</p>
                      <p className="text-[var(--red)] font-black text-lg">{uploadReport.notFoundCount}</p>
                    </div>
                  </div>
                  {uploadReport.notFoundBarcodes.length > 0 && (
                    <div className="px-4 pb-4">
                      <button
                        onClick={() => downloadBarcodeListAsExcel(uploadReport.notFoundBarcodes, 'باركودات_غير_موجودة_بالنظام')}
                        className="flex items-center gap-2 bg-white border-2 border-[var(--red)]/30 hover:bg-[var(--red)]/5 text-[var(--red)] px-4 py-2 rounded-lg text-xs font-bold transition-colors"
                      >
                        <Download size={13} />
                        تحميل قائمة الباركودات غير الموجودة
                      </button>
                    </div>
                  )}
                  <div className="px-4 pb-4">
                    <button
                      onClick={() => setActiveSection('branch')}
                      className="text-xs font-bold text-[var(--navy)] underline hover:no-underline"
                    >
                      روح لعروض فرعك الآن
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {activeSection === 'custom' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm max-w-xl">
              <div className="flex items-center gap-2 mb-1">
                <Layers size={16} className="text-[var(--navy)]" />
                <h2 className="font-black text-sm text-[var(--navy)]">ملصق مخصص</h2>
              </div>
              <p className="text-xs text-gray-600 font-medium mb-4">
                لأي حالة خاصة — مثلاً منتجات نفس النوع بأسعار متشابهة ومختلفة النكهة/الحجم. عبّي البيانات وبنجهز لك الملصق فوراً.
              </p>
              <div className="space-y-3">
                <div>
                  <label className="block text-xs font-bold text-gray-500 mb-1">اسم/عنوان الملصق</label>
                  <input
                    value={customName}
                    onChange={(e) => setCustomName(e.target.value)}
                    placeholder="مثال: ريكسونا مزيل عرق - تشكيلة متنوعة"
                    className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                  />
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-gray-500 mb-1">السعر السابق</label>
                    <input
                      value={customPrevPrice}
                      onChange={(e) => setCustomPrevPrice(e.target.value)}
                      type="number"
                      placeholder="0.00"
                      className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-gray-500 mb-1">سعر العرض</label>
                    <input
                      value={customOfferPrice}
                      onChange={(e) => setCustomOfferPrice(e.target.value)}
                      type="number"
                      placeholder="0.00"
                      className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                    />
                  </div>
                </div>
                <div>
                  <label className="block text-xs font-bold text-gray-500 mb-1">ملاحظة مكان الباركود (اختياري)</label>
                  <input
                    value={customNote}
                    onChange={(e) => setCustomNote(e.target.value)}
                    placeholder="مثال: عدة أنواع متوفرة"
                    className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                  />
                </div>
              </div>
              <button
                onClick={handleGenerateCustomLabel}
                className="mt-5 w-full bg-[var(--navy)] hover:bg-[#0f1a4d] text-white px-6 py-3 rounded-xl text-sm font-black transition-colors flex items-center justify-center gap-2"
              >
                <Sparkles size={17} />
                توليد الملصق (PDF)
              </button>
            </div>
          )}

          {activeSection === 'messages' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm max-w-2xl">
              <div className="flex items-center gap-2 mb-4">
                <MessageCircle size={16} className="text-[var(--navy)]" />
                <h2 className="font-black text-sm text-[var(--navy)]">التواصل مع الإدارة</h2>
              </div>

              <div className="flex gap-2 mb-5">
                <input
                  value={newMessageText}
                  onChange={(e) => setNewMessageText(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleSendMessage()}
                  placeholder="اكتب رسالتك للإدارة..."
                  className="flex-1 bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                />
                <button
                  onClick={handleSendMessage}
                  disabled={sendingMessage}
                  className="bg-[var(--navy)] hover:bg-[#0f1a4d] text-white px-5 py-2.5 rounded-lg text-sm font-bold transition-colors flex items-center gap-2 disabled:opacity-50"
                >
                  <Send size={15} />
                  إرسال
                </button>
              </div>

              <p className="text-xs font-bold text-gray-400 mb-2">آخر الرسائل</p>
              <div className="space-y-2 max-h-96 overflow-y-auto">
                {messages.length === 0 && (
                  <p className="text-center text-gray-400 text-sm py-6">ما فيه رسائل بعد</p>
                )}
                {messages.map((m) => (
                  <div
                    key={m.id}
                    className={`p-3 rounded-xl text-sm ${
                      m.sender_role === 'admin' ? 'bg-[var(--navy)]/5 border border-[var(--navy)]/10' : 'bg-emerald-50 border border-emerald-100'
                    }`}
                  >
                    <p className="font-bold text-[var(--navy)] mb-1">
                      {m.sender_role === 'admin' ? (m.target_branch_id ? 'الإدارة (رسالة خاصة)' : 'الإدارة (للجميع)') : 'رسالتك'}
                    </p>
                    <p className="text-gray-700">{m.body}</p>
                    <p className="text-[10px] text-gray-400 mt-1">{formatDate(m.created_at)}</p>
                  </div>
                ))}
              </div>
            </div>
          )}

          {activeSection === 'audit' && (
            <div className="space-y-4">
              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-5 shadow-sm">
                <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
                  <div>
                    <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
                      <ClipboardCheck size={16} />
                      تدقيق الملصقات
                    </h2>
                    <p className="text-xs text-gray-500 font-medium mt-1">
                      أكّد إن كل منتج فعلاً عليه ملصق بالرف — للجولة الدورية
                    </p>
                  </div>
                  <button
                    onClick={handleStartNewAuditRound}
                    disabled={auditLoading}
                    className="flex items-center gap-1.5 bg-white border-2 border-[var(--red)]/20 hover:bg-[var(--red)]/5 text-[var(--red)] text-xs font-bold px-4 py-2 rounded-lg transition-colors disabled:opacity-50"
                  >
                    {auditLoading ? 'جاري البدء...' : 'بدء دورة تدقيق جديدة'}
                  </button>
                </div>
                <div className="flex items-center gap-3">
                  <div className="flex-1 bg-gray-100 rounded-full h-3 overflow-hidden">
                    <div
                      className="bg-emerald-500 h-full transition-all"
                      style={{ width: `${matchedItems.length > 0 ? (auditCheckedCount / matchedItems.length) * 100 : 0}%` }}
                    />
                  </div>
                  <span className="text-xs font-black text-[var(--navy)] shrink-0">
                    {auditCheckedCount} من {matchedItems.length}
                  </span>
                </div>
              </div>

              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                <div className="divide-y-2 divide-[var(--navy)]/10 max-h-[600px] overflow-y-auto">
                  {matchedItems.length === 0 && (
                    <p className="p-6 text-center text-gray-400 text-sm">ما فيه منتجات بفرعك للتدقيق</p>
                  )}
                  {matchedItems.map((item) => {
                    const check = labelChecks[item.barcode]
                    const isChecked = check?.is_checked || false
                    return (
                      <div
                        key={item.id}
                        className={`flex items-center justify-between gap-3 p-3.5 ${isChecked ? 'bg-emerald-50' : 'bg-white'}`}
                      >
                        <div className="flex items-center gap-3 min-w-0">
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => handleToggleLabelCheck(item.barcode, isChecked)}
                            className="w-5 h-5 cursor-pointer accent-emerald-600 shrink-0"
                          />
                          <div className="min-w-0">
                            <p className="text-sm font-bold text-[var(--navy)] truncate">{item.product_name}</p>
                            <p className="text-[11px] text-gray-500">{item.barcode}</p>
                          </div>
                        </div>
                        {isChecked && check?.checked_at && (
                          <span className="text-[11px] text-emerald-600 font-bold shrink-0 flex items-center gap-1">
                            <CheckCircle2 size={12} />
                            {formatDate(check.checked_at)}
                          </span>
                        )}
                      </div>
                    )
                  })}
                </div>
              </div>
            </div>
          )}

          {activeSection === 'cancelled' && (
            <div className="space-y-4">
              {recentCancelBatches.length === 0 && (
                <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--red)]/40 p-6 text-center">
                  <p className="text-gray-400 text-sm">ما فيه إلغاءات حالياً</p>
                </div>
              )}
              {recentCancelBatches.map((batch) => {
                const totalBatchItems = cancelledItems.filter((i) => i.cancelled_batch_id === batch.id).length
                const batchItems = cancelledItems.filter((i) => i.cancelled_batch_id === batch.id && branchBarcodes.has(i.barcode))
                const isPreview = previewCancelBatchId === batch.id
                const removed = confirmedRemovalIds.includes(batch.id)
                return (
                  <div key={batch.id} className="bg-[var(--card)] rounded-2xl border-2 border-[var(--red)]/30 overflow-hidden shadow-sm">
                    <div className="flex flex-wrap items-center justify-between gap-3 p-4 bg-[var(--red)]/5 border-b-2 border-[var(--red)]/15">
                      <div className="min-w-0">
                        <p className="text-sm font-black text-[var(--navy)] truncate flex items-center gap-2">
                          <XCircle size={15} className="text-[var(--red)]" />
                          {batch.label}
                        </p>
                        <p className="text-[11px] text-gray-500 font-medium">
                          {formatDate(batch.created_at)} · {batchItems.length} من {totalBatchItems} منتج يخص فرعك
                        </p>
                      </div>
                      <div className="flex items-center gap-2 shrink-0 flex-wrap">
                        <button
                          onClick={() => setPreviewCancelBatchId(isPreview ? null : batch.id)}
                          className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                        >
                          {isPreview ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
                          معاينة
                        </button>
                        <button
                          onClick={() => handleDownloadFullBatch(batch, true)}
                          className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                        >
                          <Download size={13} />
                          تحميل كل منتجات الإلغاء
                        </button>
                        <button
                          onClick={() => handleDownloadBatch(batch, true)}
                          className="flex items-center gap-1.5 bg-white border-2 border-[var(--red)]/20 hover:bg-[var(--red)]/10 text-[var(--red)] text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                        >
                          <Download size={13} />
                          تحميل الخاص بفرعي
                        </button>
                        {removed ? (
                          <span className="flex items-center gap-1.5 text-emerald-600 text-xs font-bold px-2">
                            <CheckCircle2 size={15} />
                            تم الإزالة
                          </span>
                        ) : (
                          <button
                            onClick={() => handleConfirmRemoval(batch.id)}
                            className="bg-[var(--red)] hover:bg-[#c11a20] text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                          >
                            تأكيد الإزالة
                          </button>
                        )}
                      </div>
                    </div>
                    {isPreview && (
                      <div className="p-4">
                        <p className="text-[11px] text-gray-500 font-bold mb-2">
                          هذي كل منتجات الإلغاء (مو بس منتجات فرعك) — علامة ✓ خضراء تعني إنها موجودة بقائمة باركودات فرعك.
                        </p>
                        <div className="max-h-56 overflow-y-auto space-y-1.5">
                          {cancelledItems.filter((i) => i.cancelled_batch_id === batch.id).map((item) => {
                            const isMine = branchBarcodes.has(item.barcode)
                            return (
                              <div
                                key={item.id}
                                className={`flex items-center justify-between gap-3 text-xs border rounded-lg px-3 py-2 ${
                                  isMine ? 'bg-emerald-50 border-emerald-200' : 'bg-white border-[var(--red)]/15'
                                }`}
                              >
                                <span className="text-[var(--navy)] font-bold flex items-center gap-1.5 min-w-0 truncate">
                                  {isMine && <CheckCircle2 size={12} className="text-emerald-600 shrink-0" />}
                                  {item.product_name}
                                </span>
                                <div className="flex items-center gap-3 shrink-0">
                                  <span className="text-gray-500">{item.barcode}</span>
                                  <span className="text-gray-400 line-through">{item.previous_price.toFixed(2)}</span>
                                  <span className="text-[var(--red)] font-black">{item.offer_price.toFixed(2)}</span>
                                </div>
                              </div>
                            )
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                )
              })}
            </div>
          )}

          {activeSection === 'updates' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
              <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
                <h2 className="font-black text-sm text-[var(--navy)]">آخر التحديثات</h2>
              </div>
              <div className="divide-y-2 divide-[var(--navy)]/10 max-h-[600px] overflow-y-auto">
                {recentBatches.length === 0 && (
                  <p className="p-6 text-center text-gray-400 text-sm">ماله تحديثات لسا</p>
                )}
                {recentBatches.map((batch) => {
                  const confirmed = confirmedBatchIds.includes(batch.id)
                  const totalBatchItems = allItems.filter((i) => i.batch_id === batch.id).length
                  const batchItems = allItems.filter((i) => i.batch_id === batch.id && branchBarcodes.has(i.barcode))
                  const isPreview = previewUpdateBatchId === batch.id
                  return (
                    <div key={batch.id}>
                      <div className="flex flex-wrap items-center justify-between gap-3 p-4">
                        <div className="min-w-0">
                          <p className="text-sm font-bold text-[var(--navy)] truncate">{batch.label}</p>
                          <p className="text-[11px] text-gray-500 font-medium">
                            {formatDate(batch.created_at)} · {batchItems.length} من {totalBatchItems} منتج يخص فرعك
                          </p>
                        </div>
                        <div className="flex items-center gap-2 shrink-0 flex-wrap">
                          <button
                            onClick={() => setPreviewUpdateBatchId(isPreview ? null : batch.id)}
                            className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                          >
                            {isPreview ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
                            معاينة
                          </button>
                          <button
                            onClick={() => handleDownloadFullBatch(batch)}
                            className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                          >
                            <Download size={13} />
                            تحميل كل منتجات التحديث
                          </button>
                          <button
                            onClick={() => handleDownloadBatch(batch)}
                            className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                          >
                            <Download size={13} />
                            تحميل الخاص بفرعي
                          </button>
                          <button
                            onClick={() => handleGenerateBatchLabels(batchItems, batch.label, 'download')}
                            disabled={generating}
                            className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
                          >
                            <Sparkles size={13} />
                            تحميل ملصقات هذا التحديث
                          </button>
                          <button
                            onClick={() => handleGenerateBatchLabels(batchItems, batch.label, 'print')}
                            disabled={generating}
                            className="flex items-center gap-1.5 bg-[var(--navy)] hover:bg-[#0f1a4d] text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
                          >
                            <Printer size={13} />
                            طباعة مباشرة
                          </button>
                          {confirmed ? (
                            <span className="flex items-center gap-1.5 text-emerald-600 text-xs font-bold px-2">
                              <CheckCircle2 size={15} />
                              تم التفعيل
                            </span>
                          ) : (
                            <button
                              onClick={() => handleConfirmBatch(batch.id)}
                              className="bg-[var(--navy)] hover:bg-[#0f1a4d] text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                            >
                              تأكيد التفعيل
                            </button>
                          )}
                        </div>
                      </div>
                      {isPreview && (
                        <div className="bg-[var(--navy)]/5 px-4 pb-4">
                          <p className="text-[11px] text-gray-500 font-bold mb-2 pt-1">
                            هذي كل منتجات التحديث (مو بس منتجات فرعك) — عشان تقدر تحدد أي باركود ينقص من قائمة فرعك.
                          </p>
                          <div className="max-h-52 overflow-y-auto space-y-1.5">
                            {allItems.filter((i) => i.batch_id === batch.id).map((item) => {
                              const isMine = branchBarcodes.has(item.barcode)
                              return (
                                <div
                                  key={item.id}
                                  className={`flex items-center justify-between gap-3 text-xs border rounded-lg px-3 py-2 ${
                                    isMine ? 'bg-emerald-50 border-emerald-200' : 'bg-white border-[var(--navy)]/10'
                                  }`}
                                >
                                  <span className="text-[var(--navy)] font-bold flex items-center gap-1.5 min-w-0 truncate">
                                    {isMine && <CheckCircle2 size={12} className="text-emerald-600 shrink-0" />}
                                    {item.product_name}
                                  </span>
                                  <div className="flex items-center gap-3 shrink-0">
                                    <span className="text-gray-500">{item.barcode}</span>
                                    <span className="text-gray-400 line-through">{item.previous_price.toFixed(2)}</span>
                                    <span className="text-[var(--red)] font-black">{item.offer_price.toFixed(2)}</span>
                                  </div>
                                </div>
                              )
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            </div>
          )}

          {activeSection === 'general' && (
            <>
              <div className="flex items-center justify-end mb-4">
                <button
                  onClick={handleDownloadTemplate}
                  className="flex items-center gap-2 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] px-4 py-2.5 rounded-lg text-sm font-bold transition-colors"
                >
                  <Download size={15} />
                  تحميل كل العروض النشطة
                </button>
              </div>

              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
                  <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
                    <Package size={16} />
                    كل العروض ({allItems.length})
                  </h2>
                </div>
                <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                  <table className="w-full text-sm border-collapse">
                    <thead className="sticky top-0 bg-[var(--navy)] text-white">
                      <tr>
                        <th className="p-3 text-right font-bold border-2 border-white/20">الباركود</th>
                        <th className="p-3 text-right font-bold border-2 border-white/20">اسم المنتج</th>
                        <th className="p-3 text-right font-bold border-2 border-white/20">السعر السابق</th>
                        <th className="p-3 text-right font-bold border-2 border-white/20">سعر العرض</th>
                        <th className="p-3 text-center font-bold border-2 border-white/20">طباعة ملصق</th>
                      </tr>
                    </thead>
                    <tbody>
                      {allItems.map((item, i) => (
                        <tr key={item.id} className={`${i % 2 === 0 ? 'bg-white' : 'bg-[var(--navy)]/[0.03]'} hover:bg-[var(--yellow)]/10 transition-colors`}>
                          <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.barcode}</td>
                          <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.product_name}</td>
                          <td className="p-3 text-gray-500 font-bold line-through border-2 border-[var(--navy)]/10">{item.previous_price.toFixed(2)}</td>
                          <td className="p-3 text-[var(--red)] font-black border-2 border-[var(--navy)]/10">{item.offer_price.toFixed(2)}</td>
                          <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                            <button
                              onClick={() => handlePrintSingle(item)}
                              className="inline-flex items-center gap-1 text-[var(--navy)] hover:text-[var(--red)] text-xs font-bold transition-colors"
                            >
                              <Printer size={13} />
                              طباعة
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}

          {activeSection === 'branch' && (
            <>
              {matchedItems.length > 0 && (
                <div className="mb-4 flex flex-wrap items-center gap-3">
                  <button
                    onClick={() => handleGenerateAllLabels('download')}
                    disabled={generating}
                    className="bg-emerald-600 hover:bg-emerald-500 text-white px-6 py-3 rounded-xl font-black transition-colors disabled:opacity-50 flex items-center gap-2 shadow-lg shadow-emerald-600/20"
                  >
                    <Sparkles size={17} />
                    {generating ? 'جاري التوليد...' : `تحميل كل ملصقات الفرع (${matchedItems.length})`}
                  </button>
                  <button
                    onClick={() => handleGenerateAllLabels('print')}
                    disabled={generating}
                    className="bg-[var(--navy)] hover:bg-[#0f1a4d] text-white px-6 py-3 rounded-xl font-black transition-colors disabled:opacity-50 flex items-center gap-2 shadow-lg shadow-[var(--navy)]/20"
                  >
                    <Printer size={17} />
                    طباعة مباشرة
                  </button>
                </div>
              )}

              {matchedItems.length === 0 && (
                <div className="p-6 bg-[var(--yellow)]/10 border-2 border-[var(--yellow)]/40 rounded-2xl text-center mb-4">
                  <p className="text-sm font-bold text-[#8a6300]">
                    ما رفعت باركودات فرعك بعد. روح تبويب "رفع ملف الفرع" وارفع ملف فيه باركودات منتجاتك.
                  </p>
                </div>
              )}

              {/* تبديل بين عرض مقسّم حسب كل تحديث، أو عرض كل منتجات الفرع بقائمة واحدة */}
              <div className="flex items-center gap-2 mb-4">
                <button
                  onClick={() => setBranchViewMode('grouped')}
                  className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors ${
                    branchViewMode === 'grouped' ? 'bg-[var(--navy)] text-white' : 'bg-white border-2 border-[var(--navy)]/15 text-[var(--navy)] hover:bg-[var(--navy)]/10'
                  }`}
                >
                  حسب كل تحديث
                </button>
                <button
                  onClick={() => setBranchViewMode('all')}
                  className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors ${
                    branchViewMode === 'all' ? 'bg-[var(--navy)] text-white' : 'bg-white border-2 border-[var(--navy)]/15 text-[var(--navy)] hover:bg-[var(--navy)]/10'
                  }`}
                >
                  عرض الكل بدون تحديد تحديث ({matchedItems.length})
                </button>
              </div>

              {branchViewMode === 'all' && (
                <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                  <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
                    <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
                      <Package size={16} />
                      كل عروض الفرع ({matchedItems.length})
                    </h2>
                  </div>
                  <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                    <table className="w-full text-sm border-collapse">
                      <thead className="sticky top-0 bg-[var(--navy)] text-white">
                        <tr>
                          <th className="p-3 text-right font-bold border-2 border-white/20">الباركود</th>
                          <th className="p-3 text-right font-bold border-2 border-white/20">اسم المنتج</th>
                          <th className="p-3 text-right font-bold border-2 border-white/20">السعر السابق</th>
                          <th className="p-3 text-right font-bold border-2 border-white/20">سعر العرض</th>
                          <th className="p-3 text-center font-bold border-2 border-white/20">طباعة ملصق</th>
                        </tr>
                      </thead>
                      <tbody>
                        {matchedItems.map((item, i) => (
                          <tr key={item.id} className={`${i % 2 === 0 ? 'bg-white' : 'bg-[var(--navy)]/[0.03]'} hover:bg-[var(--yellow)]/10 transition-colors`}>
                            <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.barcode}</td>
                            <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.product_name}</td>
                            <td className="p-3 text-gray-500 font-bold line-through border-2 border-[var(--navy)]/10">{item.previous_price.toFixed(2)}</td>
                            <td className="p-3 text-[var(--red)] font-black border-2 border-[var(--navy)]/10">{item.offer_price.toFixed(2)}</td>
                            <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                              <button
                                onClick={() => handlePrintSingle(item)}
                                className="inline-flex items-center gap-1 text-[var(--navy)] hover:text-[var(--red)] text-xs font-bold transition-colors"
                              >
                                <Printer size={13} />
                                طباعة
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* التحديثات مقسّمة كل واحد لحاله حسب باركودات فرعك */}
              {branchViewMode === 'grouped' && (
              <div className="space-y-4">
                {branchBatchGroups.map(({ batch, items }) => {
                  const expanded = expandedBranchBatchId === batch.id
                  return (
                    <div key={batch.id} className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                      <div className="flex flex-wrap items-center justify-between gap-3 p-4 bg-[var(--navy)]/5 border-b-2 border-[var(--navy)]/10">
                        <div className="min-w-0">
                          <p className="text-sm font-black text-[var(--navy)] truncate">{batch.label}</p>
                          <p className="text-[11px] text-gray-500 font-medium">{formatDate(batch.created_at)} · {items.length} منتج يخص فرعك</p>
                        </div>
                        <div className="flex items-center gap-2 shrink-0">
                          <button
                            onClick={() => setExpandedBranchBatchId(expanded ? null : batch.id)}
                            className="flex items-center gap-1.5 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                          >
                            {expanded ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
                            {expanded ? 'إخفاء المنتجات' : 'عرض المنتجات'}
                          </button>
                          <button
                            onClick={() => handleGenerateBatchLabels(items, batch.label, 'download')}
                            disabled={generating}
                            className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
                          >
                            <Sparkles size={13} />
                            تحميل هذا التحديث
                          </button>
                          <button
                            onClick={() => handleGenerateBatchLabels(items, batch.label, 'print')}
                            disabled={generating}
                            className="flex items-center gap-1.5 bg-[var(--navy)] hover:bg-[#0f1a4d] text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors disabled:opacity-50"
                          >
                            <Printer size={13} />
                            طباعة مباشرة
                          </button>
                        </div>
                      </div>
                      {expanded && (
                        <div className="overflow-x-auto">
                          <table className="w-full text-sm border-collapse">
                            <thead className="bg-[var(--navy)] text-white">
                              <tr>
                                <th className="p-3 text-right font-bold border-2 border-white/20">الباركود</th>
                                <th className="p-3 text-right font-bold border-2 border-white/20">اسم المنتج</th>
                                <th className="p-3 text-right font-bold border-2 border-white/20">السعر السابق</th>
                                <th className="p-3 text-right font-bold border-2 border-white/20">سعر العرض</th>
                                <th className="p-3 text-center font-bold border-2 border-white/20">طباعة ملصق</th>
                              </tr>
                            </thead>
                            <tbody>
                              {items.map((item, i) => (
                                <tr key={item.id} className={`${i % 2 === 0 ? 'bg-white' : 'bg-[var(--navy)]/[0.03]'} hover:bg-[var(--yellow)]/10 transition-colors`}>
                                  <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.barcode}</td>
                                  <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.product_name}</td>
                                  <td className="p-3 text-gray-500 font-bold line-through border-2 border-[var(--navy)]/10">{item.previous_price.toFixed(2)}</td>
                                  <td className="p-3 text-[var(--red)] font-black border-2 border-[var(--navy)]/10">{item.offer_price.toFixed(2)}</td>
                                  <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                                    <button
                                      onClick={() => handlePrintSingle(item)}
                                      className="inline-flex items-center gap-1 text-[var(--navy)] hover:text-[var(--red)] text-xs font-bold transition-colors"
                                    >
                                      <Printer size={13} />
                                      طباعة
                                    </button>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  )
                })}

                {noBatchMatchedItems.length > 0 && (
                  <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                    <div className="p-4 bg-[var(--navy)]/5 border-b-2 border-[var(--navy)]/10">
                      <h3 className="text-sm font-black text-[var(--navy)]">منتجات فرعك بدون تحديث محدد ({noBatchMatchedItems.length})</h3>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-sm border-collapse">
                        <thead className="bg-[var(--navy)] text-white">
                          <tr>
                            <th className="p-3 text-right font-bold border-2 border-white/20">الباركود</th>
                            <th className="p-3 text-right font-bold border-2 border-white/20">اسم المنتج</th>
                            <th className="p-3 text-right font-bold border-2 border-white/20">السعر السابق</th>
                            <th className="p-3 text-right font-bold border-2 border-white/20">سعر العرض</th>
                            <th className="p-3 text-center font-bold border-2 border-white/20">طباعة ملصق</th>
                          </tr>
                        </thead>
                        <tbody>
                          {noBatchMatchedItems.map((item, i) => (
                            <tr key={item.id} className={`${i % 2 === 0 ? 'bg-white' : 'bg-[var(--navy)]/[0.03]'} hover:bg-[var(--yellow)]/10 transition-colors`}>
                              <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.barcode}</td>
                              <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.product_name}</td>
                              <td className="p-3 text-gray-500 font-bold line-through border-2 border-[var(--navy)]/10">{item.previous_price.toFixed(2)}</td>
                              <td className="p-3 text-[var(--red)] font-black border-2 border-[var(--navy)]/10">{item.offer_price.toFixed(2)}</td>
                              <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                                <button
                                  onClick={() => handlePrintSingle(item)}
                                  className="inline-flex items-center gap-1 text-[var(--navy)] hover:text-[var(--red)] text-xs font-bold transition-colors"
                                >
                                  <Printer size={13} />
                                  طباعة
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  )
}
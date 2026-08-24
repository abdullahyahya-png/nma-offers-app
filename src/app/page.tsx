'use client'

export const dynamic = 'force-dynamic'

import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'
import { renderPdfToCanvas } from '../lib/pdfBackground'
import WhatsAppNotifySection from './components/WhatsAppNotifySection'
import BranchAccountsSection from './components/BranchAccountsSection'
import InstallPWAButton from './components/InstallPWAButton'
import AuditOverviewSection from './components/AuditOverviewSection'
import {
  UploadCloud,
  ImagePlus,
  PlusCircle,
  Trash2,
  Pencil,
  Check,
  Search,
  Package,
  TrendingUp,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  XCircle,
  X,
  History,
  ChevronUp,
  FileCheck,
  Download,
  LayoutGrid,
  Bell,
  Settings,
  LogOut,
  CheckCircle2,
  MessageCircle,
  Send,
  Store,
  Phone,
  KeyRound,
  ClipboardCheck,
} from 'lucide-react'

interface OfferItem {
  id?: string
  barcode: string
  product_name: string
  previous_price: number
  offer_price: number
  created_at?: string
  is_active?: boolean
  batch_id?: string | null
  cancelled_batch_id?: string | null
}

interface OfferBatch {
  id: string
  label: string
  batch_type: 'new' | 'cancel'
  created_at: string
}

interface Branch {
  id: string
  name: string
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

interface ActivityLog {
  id: string
  branch_id: string | null
  actor_role: 'admin' | 'branch'
  action: string
  reference_id: string | null
  details: string | null
  created_at: string
}

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100]

type SectionId = 'stats' | 'upload' | 'label' | 'cancel' | 'history' | 'manual' | 'messages' | 'table' | 'whatsapp' | 'activity' | 'accounts' | 'audit' | 'settings'
type TableFilter = 'active' | 'all' | 'cancelled'

const SECTIONS: { id: SectionId; label: string; icon: any }[] = [
  { id: 'stats', label: 'الإحصائيات', icon: LayoutGrid },
  { id: 'upload', label: 'رفع تحديث جديد', icon: UploadCloud },
  { id: 'label', label: 'ملصق العروض', icon: ImagePlus },
  { id: 'cancel', label: 'إدارة الإلغاء', icon: XCircle },
  { id: 'history', label: 'سجل التحديثات', icon: History },
  { id: 'manual', label: 'إضافة يدوية', icon: PlusCircle },
  { id: 'messages', label: 'الرسائل', icon: MessageCircle },
  { id: 'whatsapp', label: 'واتساب الفروع', icon: Phone },
  { id: 'accounts', label: 'حسابات الفروع', icon: KeyRound },
  { id: 'audit', label: 'تدقيق الملصقات', icon: ClipboardCheck },
  { id: 'activity', label: 'سجل النشاط', icon: History },
  { id: 'table', label: 'جدول المنتجات', icon: Package },
]

function stripExtension(filename: string) {
  return filename.replace(/\.(xlsx|xls)$/i, '')
}

function formatDate(dateStr: string) {
  const d = new Date(dateStr)
  return d.toLocaleDateString('ar-SA', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
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

export default function AdminPage() {
  const [items, setItems] = useState<OfferItem[]>([])
  const [batches, setBatches] = useState<OfferBatch[]>([])
  const [branches, setBranches] = useState<Branch[]>([])
  const [expandedBatchId, setExpandedBatchId] = useState<string | null>(null)
  const [status, setStatus] = useState('')
  const [uploading, setUploading] = useState(false)
  const [uploadingCancel, setUploadingCancel] = useState(false)
  const [uploadingBg, setUploadingBg] = useState(false)
  const [loading, setLoading] = useState(true)
  const [tableFilter, setTableFilter] = useState<TableFilter>('active')
  const [activeSection, setActiveSection] = useState<SectionId>('stats')
  const [labelPreviewUrl, setLabelPreviewUrl] = useState<string | null>(null)
  const [labelExists, setLabelExists] = useState(true)

  const [newBarcode, setNewBarcode] = useState('')
  const [newName, setNewName] = useState('')
  const [newPrevPrice, setNewPrevPrice] = useState('')
  const [newOfferPrice, setNewOfferPrice] = useState('')

  const [pendingFile, setPendingFile] = useState<File | null>(null)
  const [batchLabel, setBatchLabel] = useState('')

  const [pendingCancelFile, setPendingCancelFile] = useState<File | null>(null)
  const [cancelBatchLabel, setCancelBatchLabel] = useState('')

  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(10)

  const [messages, setMessages] = useState<Message[]>([])
  const [activityLogs, setActivityLogs] = useState<ActivityLog[]>([])
  const [messageTarget, setMessageTarget] = useState('')
  const [newMessageText, setNewMessageText] = useState('')
  const [sendingMessage, setSendingMessage] = useState(false)

  const [activationMap, setActivationMap] = useState<Record<string, Set<string>>>({})
  const [removalMap, setRemovalMap] = useState<Record<string, Set<string>>>({})
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set())
  const [tableSearchQuery, setTableSearchQuery] = useState('')
  const [editingId, setEditingId] = useState<string | null>(null)
  const [editPrevPrice, setEditPrevPrice] = useState('')
  const [editOfferPrice, setEditOfferPrice] = useState('')
  const [editName, setEditName] = useState('')
  const [savingEdit, setSavingEdit] = useState(false)
  const [confirmDialog, setConfirmDialog] = useState<{ message: string; onConfirm: () => void } | null>(null)
  const [showFactoryResetModal, setShowFactoryResetModal] = useState(false)
  const [factoryResetInput, setFactoryResetInput] = useState('')
  const [resettingFactory, setResettingFactory] = useState(false)

  const askConfirm = (message: string, onConfirm: () => void) => {
    setConfirmDialog({ message, onConfirm })
  }

  const wheelCooldownRef = useRef(false)

  const handleTableWheel = (e: React.WheelEvent) => {
    if (wheelCooldownRef.current) return
    if (e.deltaY > 15) {
      setPage((p) => Math.min(totalPages, p + 1))
    } else if (e.deltaY < -15) {
      setPage((p) => Math.max(1, p - 1))
    } else {
      return
    }
    wheelCooldownRef.current = true
    setTimeout(() => {
      wheelCooldownRef.current = false
    }, 450)
  }

  const fetchItems = async () => {
    setLoading(true)
    const { data, error } = await supabase
      .from('offer_items')
      .select('*')
      .order('created_at', { ascending: false })
    if (!error && data) setItems(data)
    setLoading(false)
  }

  const fetchBatches = async () => {
    const { data, error } = await supabase
      .from('offer_batches')
      .select('*')
      .order('created_at', { ascending: false })
    if (!error && data) setBatches(data)
  }

  const fetchBranches = async () => {
    const { data } = await supabase.from('branches').select('id, name, manager_phone').order('name')
    if (data) setBranches(data)
  }

  const fetchMessages = async () => {
    const { data } = await supabase
      .from('messages')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(50)
    if (data) setMessages(data)
  }

  const fetchActivityLogs = async () => {
    const { data } = await supabase
      .from('activity_logs')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(100)
    if (data) setActivityLogs(data)
  }

  // يسجل عملية إدارية بسجل النشاط
  const logActivity = async (action: string, referenceId?: string, details?: string) => {
    await supabase.from('activity_logs').insert([{
      branch_id: null,
      actor_role: 'admin',
      action,
      reference_id: referenceId || null,
      details: details || null,
    }])
  }

  const fetchConfirmations = async () => {
    const { data: activations } = await supabase.from('branch_batch_confirmations').select('branch_id, batch_id')
    const { data: removals } = await supabase.from('branch_cancel_confirmations').select('branch_id, batch_id')

    const actMap: Record<string, Set<string>> = {}
    activations?.forEach((row) => {
      if (!actMap[row.batch_id]) actMap[row.batch_id] = new Set()
      actMap[row.batch_id].add(row.branch_id)
    })
    setActivationMap(actMap)

    const remMap: Record<string, Set<string>> = {}
    removals?.forEach((row) => {
      if (!remMap[row.batch_id]) remMap[row.batch_id] = new Set()
      remMap[row.batch_id].add(row.branch_id)
    })
    setRemovalMap(remMap)
  }

  const refreshLabelPreview = async () => {
    const { data } = supabase.storage.from('label-assets').getPublicUrl('label-bg.pdf')
    const url = `${data.publicUrl}?t=${Date.now()}`
    console.log('محاولة تحميل قالب الملصق من:', url)
    try {
      const canvas = await renderPdfToCanvas(url, 600, 1000)
      setLabelPreviewUrl(canvas.toDataURL('image/png'))
      setLabelExists(true)
    } catch (err) {
      console.error('فشل تحميل/رسم قالب الملصق:', err)
      setLabelPreviewUrl(null)
      setLabelExists(false)
    }
  }

  useEffect(() => {
    fetchItems()
    fetchBatches()
    fetchBranches()
    fetchMessages()
    fetchConfirmations()
    fetchActivityLogs()
    refreshLabelPreview()

    // Realtime: الرسائل الجديدة من الفروع وتأكيداتهم تظهر لحظياً بدون Refresh
    const channel = supabase
      .channel('admin-page-realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'messages' }, () => {
        fetchMessages()
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'branch_batch_confirmations' }, () => {
        fetchConfirmations()
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'branch_cancel_confirmations' }, () => {
        fetchConfirmations()
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'activity_logs' }, () => {
        fetchActivityLogs()
      })
      .subscribe()

    return () => {
      supabase.removeChannel(channel)
    }
  }, [])

  // يعلّم رسائل الفروع كمقروءة تلقائياً لما الأدمن يفتح تبويب الرسائل
  useEffect(() => {
    if (activeSection !== 'messages') return
    const unreadIds = messages.filter((m) => m.sender_role === 'branch' && !m.is_read).map((m) => m.id)
    if (unreadIds.length === 0) return
    supabase
      .from('messages')
      .update({ is_read: true, read_at: new Date().toISOString() })
      .in('id', unreadIds)
      .then(() => {
        setMessages((prev) => prev.map((m) => (unreadIds.includes(m.id) ? { ...m, is_read: true } : m)))
      })
  }, [activeSection, messages])

  const activeItemsList = useMemo(() => items.filter((i) => i.is_active !== false), [items])
  const cancelledItemsList = useMemo(() => items.filter((i) => i.is_active === false), [items])
  const cancelledCount = cancelledItemsList.length

  const visibleItems = useMemo(() => {
    if (tableFilter === 'active') return activeItemsList
    if (tableFilter === 'cancelled') return cancelledItemsList
    return items
  }, [tableFilter, items, activeItemsList, cancelledItemsList])

  const searchedItems = useMemo(() => {
    const q = tableSearchQuery.trim()
    if (!q) return visibleItems
    return visibleItems.filter((i) => i.barcode.includes(q) || i.product_name.includes(q))
  }, [visibleItems, tableSearchQuery])

  const avgDiscount = useMemo(() => {
    if (activeItemsList.length === 0) return 0
    const total = activeItemsList.reduce((sum, item) => {
      if (!item.previous_price) return sum
      return sum + (1 - item.offer_price / item.previous_price)
    }, 0)
    return Math.round((total / activeItemsList.length) * 100)
  }, [activeItemsList])

  const totalPages = Math.max(1, Math.ceil(searchedItems.length / pageSize))
  const paginatedItems = searchedItems.slice((page - 1) * pageSize, page * pageSize)

  useEffect(() => {
    setPage(1)
    setSelectedIds(new Set())
  }, [pageSize, visibleItems.length, tableFilter, tableSearchQuery])

  const itemsForBatch = (batch: OfferBatch) => {
    return batch.batch_type === 'new'
      ? items.filter((i) => i.batch_id === batch.id)
      : items.filter((i) => i.cancelled_batch_id === batch.id)
  }

  const branchName = (id: string | null) => {
    if (!id) return null
    return branches.find((b) => b.id === id)?.name || 'فرع غير معروف'
  }

  const handleDownloadCancelled = () => {
    if (cancelledItemsList.length === 0) {
      setStatus('ما فيه عروض ملغاة حالياً')
      return
    }
    downloadItemsAsExcel(cancelledItemsList, 'عروض_ملغاة')
  }

  const handleSelectUpdateFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setPendingFile(file)
    setBatchLabel(stripExtension(file.name))
  }

  const handleConfirmUpload = () => {
    if (!pendingFile) return
    const reader = new FileReader()
    reader.onload = async (event) => {
      const data = event.target?.result
      const workbook = XLSX.read(data, { type: 'binary' })
      const sheetName = workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]
      const rows: any[] = XLSX.utils.sheet_to_json(sheet, { header: 1 })
      const rawParsedItems = rows.slice(1)
        .filter((row) => row.length >= 4 && row[0])
        .map((row) => ({
          barcode: String(row[0]).trim(),
          product_name: String(row[1]).trim(),
          previous_price: Number(row[2]),
          offer_price: Number(row[3]),
        }))

      // نشيل التكرار داخل نفس الملف (نفس الباركود بأكثر من صف) — نحتفظ بآخر صف مكرر
      const dedupedMap = new Map<string, typeof rawParsedItems[0]>()
      rawParsedItems.forEach((item) => dedupedMap.set(item.barcode, item))
      const parsedItems = Array.from(dedupedMap.values())
      const duplicatesInFile = rawParsedItems.length - parsedItems.length

      setUploading(true)
      setStatus(`جاري إنشاء التحديث "${batchLabel}"...`)

      const { data: batchData, error: batchError } = await supabase
        .from('offer_batches')
        .insert([{ label: batchLabel || stripExtension(pendingFile.name), batch_type: 'new' }])
        .select()
        .single()

      if (batchError || !batchData) {
        setUploading(false)
        setStatus(`خطأ بإنشاء التحديث: ${batchError?.message}`)
        return
      }

      const itemsWithBatch = parsedItems.map((item) => ({
        ...item,
        batch_id: batchData.id,
        is_active: true,
        cancelled_batch_id: null,
      }))

      const { error } = await supabase
        .from('offer_items')
        .upsert(itemsWithBatch, { onConflict: 'barcode' })

      setUploading(false)

      if (error) {
        setStatus(`خطأ: ${error.message}`)
      } else {
        setStatus(
          `تم حفظ ${parsedItems.length} منتج ضمن "${batchLabel}"` +
          (duplicatesInFile > 0 ? ` — تم تجاهل ${duplicatesInFile} صف مكرر بنفس الباركود داخل الملف` : '')
        )
        setPendingFile(null)
        setBatchLabel('')
        logActivity('رفع تحديث جديد', batchData.id, `${batchLabel} — ${parsedItems.length} منتج`)
        fetchItems()
        fetchBatches()
      }
    }
    reader.readAsBinaryString(pendingFile)
  }

  const handleSelectCancelFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setPendingCancelFile(file)
    setCancelBatchLabel(`إلغاء ${stripExtension(file.name)}`)
  }

  const handleConfirmCancelUpload = () => {
    if (!pendingCancelFile) return
    const reader = new FileReader()
    reader.onload = async (event) => {
      const data = event.target?.result
      const workbook = XLSX.read(data, { type: 'binary' })
      const sheetName = workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]
      const rows: any[] = XLSX.utils.sheet_to_json(sheet, { header: 1 })
      const barcodesToCancel = rows
        .slice(1)
        .filter((row) => row.length >= 1 && row[0])
        .map((row) => String(row[0]).trim())

      if (barcodesToCancel.length === 0) {
        setStatus('الملف فاضي أو مافيه باركودات')
        return
      }

      setUploadingCancel(true)
      setStatus(`جاري إلغاء ${barcodesToCancel.length} منتج...`)

      const { data: batchData, error: batchError } = await supabase
        .from('offer_batches')
        .insert([{ label: cancelBatchLabel || `إلغاء ${stripExtension(pendingCancelFile.name)}`, batch_type: 'cancel' }])
        .select()
        .single()

      if (batchError || !batchData) {
        setUploadingCancel(false)
        setStatus(`خطأ بإنشاء الإلغاء: ${batchError?.message}`)
        return
      }

      const { error } = await supabase
        .from('offer_items')
        .update({ is_active: false, cancelled_batch_id: batchData.id })
        .in('barcode', barcodesToCancel)

      setUploadingCancel(false)

      if (error) {
        setStatus(`خطأ: ${error.message}`)
      } else {
        setStatus(`تم إلغاء ${barcodesToCancel.length} منتج ضمن "${cancelBatchLabel}"`)
        setPendingCancelFile(null)
        setCancelBatchLabel('')
        logActivity('رفع تحديث إلغاء', batchData.id, `${cancelBatchLabel} — ${barcodesToCancel.length} منتج`)
        fetchItems()
        fetchBatches()
      }
    }
    reader.readAsBinaryString(pendingCancelFile)
  }

  const handleUploadBackground = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setUploadingBg(true)
    setStatus('جاري رفع قالب الملصق...')
    const { error: uploadError } = await supabase.storage
      .from('label-assets')
      .upload('label-bg.pdf', file, { upsert: true, contentType: 'application/pdf' })
    setUploadingBg(false)
    if (uploadError) {
      setStatus(`خطأ برفع القالب: ${uploadError.message}`)
    } else {
      setStatus('تم تحديث قالب الملصق بنجاح')
      refreshLabelPreview()
    }
  }

  const handleDeleteBackground = async () => {
    askConfirm('متأكد إنك تبي تحذف قالب الملصق الحالي؟ الفروع ما بيقدروا يطبعوا ملصقات لين ترفع وحدة بديلة.', async () => {
      const { error } = await supabase.storage.from('label-assets').remove(['label-bg.pdf'])
      if (error) {
        setStatus(`خطأ بالحذف: ${error.message}`)
      } else {
        setStatus('تم حذف قالب الملصق')
        refreshLabelPreview()
      }
    })
  }

  const handleAddManual = async () => {
    if (!newBarcode || !newName || !newPrevPrice || !newOfferPrice) {
      setStatus('عبّي كل الحقول أول')
      return
    }
    const { error } = await supabase.from('offer_items').upsert(
      [{
        barcode: newBarcode.trim(),
        product_name: newName.trim(),
        previous_price: Number(newPrevPrice),
        offer_price: Number(newOfferPrice),
        is_active: true,
      }],
      { onConflict: 'barcode' }
    )
    if (error) {
      setStatus(`خطأ: ${error.message}`)
    } else {
      setStatus('تمت الإضافة بنجاح')
      setNewBarcode('')
      setNewName('')
      setNewPrevPrice('')
      setNewOfferPrice('')
      fetchItems()
    }
  }

  const handleReactivate = async (id: string) => {
    const { error } = await supabase.from('offer_items').update({ is_active: true, cancelled_batch_id: null }).eq('id', id)
    if (!error) fetchItems()
  }

  const handleDelete = async (id: string) => {
    askConfirm('متأكد إنك تبي تحذف هذا المنتج نهائياً؟', async () => {
      const { error } = await supabase.from('offer_items').delete().eq('id', id)
      if (error) {
        setStatus(`خطأ بالحذف: ${error.message}`)
      } else {
        setItems((prev) => prev.filter((item) => item.id !== id))
      }
    })
  }

  const handleStartEdit = (item: OfferItem) => {
    setEditingId(item.id!)
    setEditName(item.product_name)
    setEditPrevPrice(String(item.previous_price))
    setEditOfferPrice(String(item.offer_price))
  }

  const handleCancelEdit = () => {
    setEditingId(null)
    setEditName('')
    setEditPrevPrice('')
    setEditOfferPrice('')
  }

  const handleSaveEdit = async (id: string) => {
    const name = editName.trim()
    const prev = Number(editPrevPrice)
    const offer = Number(editOfferPrice)
    if (!name) {
      setStatus('اسم المنتج ما يصير فاضي')
      return
    }
    if (isNaN(prev) || isNaN(offer) || prev < 0 || offer < 0) {
      setStatus('تأكد إن السعرين أرقام صحيحة')
      return
    }
    setSavingEdit(true)
    const { error } = await supabase
      .from('offer_items')
      .update({ product_name: name, previous_price: prev, offer_price: offer })
      .eq('id', id)
    setSavingEdit(false)

    if (error) {
      setStatus(`خطأ بالتعديل: ${error.message}`)
    } else {
      setItems((prevItems) =>
        prevItems.map((item) => (item.id === id ? { ...item, product_name: name, previous_price: prev, offer_price: offer } : item))
      )
      setStatus('تم تعديل المنتج بنجاح')
      logActivity('تعديل منتج', id)
      handleCancelEdit()
    }
  }

  const toggleSelectItem = (id: string) => {
    setSelectedIds((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  }

  const toggleSelectAllOnPage = () => {
    const pageIds = paginatedItems.map((i) => i.id!).filter(Boolean)
    const allSelected = pageIds.length > 0 && pageIds.every((id) => selectedIds.has(id))
    setSelectedIds((prev) => {
      const next = new Set(prev)
      if (allSelected) {
        pageIds.forEach((id) => next.delete(id))
      } else {
        pageIds.forEach((id) => next.add(id))
      }
      return next
    })
  }

  const handleBulkDelete = async () => {
    if (selectedIds.size === 0) return
    askConfirm(`متأكد إنك تبي تحذف ${selectedIds.size} منتج نهائياً؟ هذا الإجراء ما يُرجع.`, async () => {
      const idsArray = Array.from(selectedIds)
      const CHUNK_SIZE = 150 // تفادي خطأ Bad Request بسبب طول الرابط مع عدد كبير من المعرفات
      let deletedCount = 0
      let failedChunk = false

      for (let i = 0; i < idsArray.length; i += CHUNK_SIZE) {
        const chunk = idsArray.slice(i, i + CHUNK_SIZE)
        const { error } = await supabase.from('offer_items').delete().in('id', chunk)
        if (error) {
          setStatus(`خطأ بالحذف الجماعي (بعد حذف ${deletedCount} منتج): ${error.message}`)
          failedChunk = true
          break
        }
        deletedCount += chunk.length
      }

      const deletedIds = new Set(idsArray.slice(0, deletedCount))
      setItems((prev) => prev.filter((item) => !deletedIds.has(item.id!)))
      setSelectedIds((prev) => {
        const next = new Set(prev)
        deletedIds.forEach((id) => next.delete(id))
        return next
      })

      if (!failedChunk) {
        setStatus(`تم حذف ${deletedCount} منتج بنجاح`)
      }
    })
  }

  const handleDeleteBatch = async (batch: OfferBatch) => {
    const batchItems = itemsForBatch(batch)
    const itemsCount = batchItems.length
    askConfirm(
      `متأكد إنك تبي تحذف تحديث "${batch.label}" بالكامل؟\nبيتم حذف ${itemsCount} منتج مرتبط فيه نهائياً من كل الفروع، وهذا الإجراء ما يُرجع.`,
      async () => {
        // نحذف حسب رقم التحديث نفسه بدل قائمة طويلة من المعرفات (تفادي خطأ Bad Request مع التحديثات الكبيرة)
        const deleteColumn = batch.batch_type === 'new' ? 'batch_id' : 'cancelled_batch_id'
        const { error: itemsError } = await supabase
          .from('offer_items')
          .delete()
          .eq(deleteColumn, batch.id)

        if (itemsError) {
          setStatus(`خطأ بحذف منتجات التحديث: ${itemsError.message}`)
          return
        }

        if (batch.batch_type === 'new') {
          await supabase.from('branch_batch_confirmations').delete().eq('batch_id', batch.id)
        } else {
          await supabase.from('branch_cancel_confirmations').delete().eq('batch_id', batch.id)
        }

        const { error: batchError } = await supabase.from('offer_batches').delete().eq('id', batch.id)
        if (batchError) {
          setStatus(`خطأ بحذف التحديث: ${batchError.message}`)
          return
        }

        setStatus(`تم حذف تحديث "${batch.label}" وكل منتجاته (${itemsCount}) نهائياً`)
        fetchItems()
        fetchBatches()
        fetchConfirmations()
      }
    )
  }

  const handleFactoryReset = async () => {
    setResettingFactory(true)
    try {
      await supabase.from('offer_items').delete().not('id', 'is', null)
      await supabase.from('offer_batches').delete().not('id', 'is', null)
      await supabase.from('branch_batch_confirmations').delete().not('id', 'is', null)
      await supabase.from('branch_cancel_confirmations').delete().not('id', 'is', null)
      await supabase.from('branch_barcodes').delete().not('id', 'is', null)
      await supabase.from('label_checks').delete().not('id', 'is', null)
      await supabase.from('messages').delete().not('id', 'is', null)
      await supabase.from('activity_logs').delete().not('id', 'is', null)
      await supabase.storage.from('label-assets').remove(['label-bg.png', 'label-bg.pdf'])

      setStatus('تمت التهيئة الكاملة بنجاح — النظام رجع فاضي زي أول تشغيل')
      setFactoryResetInput('')
      setShowFactoryResetModal(false)
      fetchItems()
      fetchBatches()
      fetchMessages()
      fetchActivityLogs()
      fetchConfirmations()
      refreshLabelPreview()
    } catch (err: any) {
      setStatus(`صار خطأ أثناء التهيئة: ${err?.message || 'غير معروف'}`)
    } finally {
      setResettingFactory(false)
    }
  }

  const handleSendMessage = async () => {
    if (!newMessageText.trim()) return
    setSendingMessage(true)
    const { error } = await supabase.from('messages').insert([{
      sender_role: 'admin',
      sender_branch_id: null,
      target_branch_id: messageTarget || null,
      body: newMessageText.trim(),
    }])
    setSendingMessage(false)
    if (error) {
      setStatus(`خطأ بالإرسال: ${error.message}`)
    } else {
      setNewMessageText('')
      setStatus(messageTarget ? 'تم إرسال الرسالة للفرع المحدد' : 'تم إرسال الرسالة لكل الفروع')
      logActivity('إرسال رسالة', messageTarget || undefined, messageTarget ? undefined : 'لكل الفروع')
      fetchMessages()
    }
  }

  const discountBadgeClass = (discount: number) => {
    if (discount >= 20) return 'bg-emerald-100 text-emerald-700'
    if (discount >= 10) return 'bg-[#FFC72C]/25 text-[#8a6300]'
    return 'bg-rose-100 text-rose-700'
  }

  const goToTable = (filter: TableFilter) => {
    setTableFilter(filter)
    setActiveSection('table')
  }

  const unreadFromBranches = messages.filter((m) => m.sender_role === 'branch' && !m.is_read).length

  return (
    <div className="min-h-screen bg-[var(--background)]">
      <InstallPWAButton />
      <header className="bg-white border-b-4 border-[var(--navy)]">
        <div className="max-w-7xl mx-auto px-6 md:px-10 py-4 flex items-center gap-4">
          <img src="/logo.png" alt="شعار العروض" className="w-16 h-16 md:w-20 md:h-20 object-contain shrink-0" />
          <div>
            <p className="text-[var(--red)] text-xs font-bold">مركز التحكم</p>
            <h1 className="text-[var(--navy)] text-xl md:text-2xl font-black">إدارة عروض المنتجات</h1>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 md:px-10 py-6 md:py-8 flex flex-col md:flex-row gap-6 items-start">
        <aside className="w-full md:w-72 md:shrink-0 md:sticky md:top-8 space-y-4">
          <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-4 shadow-sm">
            <div className="flex items-center justify-between mb-3">
              <button
                onClick={() => setActiveSection('messages')}
                className="relative text-[var(--navy)]"
              >
                <Bell size={19} />
                {unreadFromBranches > 0 && (
                  <span className="absolute -top-1.5 -left-1.5 w-4 h-4 rounded-full bg-[var(--red)] text-white text-[9px] font-bold flex items-center justify-center">
                    {unreadFromBranches}
                  </span>
                )}
              </button>
              <div className="flex items-center gap-2">
                <p className="text-sm font-bold text-[var(--navy)]">مرحباً عبدالله</p>
                <div className="w-9 h-9 rounded-full bg-[var(--navy)] flex items-center justify-center font-bold text-white text-sm">
                  ع
                </div>
              </div>
            </div>
          </div>

          <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-4 shadow-sm">
            <p className="text-xs font-bold text-gray-400 mb-3">لوحة التحكم والعمليات</p>
            <button
              onClick={() => goToTable('all')}
              className="w-full flex items-center justify-between bg-[var(--navy)] text-white px-4 py-3 rounded-xl text-sm font-bold mb-4"
            >
              <span className="flex items-center gap-2">
                <LayoutGrid size={16} />
                جميع العروض
              </span>
              <span className="bg-white/20 px-2 py-0.5 rounded-full text-xs">{items.length}</span>
            </button>

            <p className="text-xs font-bold text-gray-400 mb-2">الحالة والتقدم</p>
            <div className="space-y-2 mb-4">
              <button
                onClick={() => goToTable('active')}
                className={`w-full flex items-center justify-between px-3 py-2 rounded-lg transition-colors ${
                  tableFilter === 'active' && activeSection === 'table' ? 'bg-emerald-100 ring-2 ring-emerald-400' : 'bg-emerald-50 hover:bg-emerald-100'
                }`}
              >
                <span className="flex items-center gap-2 text-sm font-bold text-emerald-700">
                  <CheckCircle2 size={15} />
                  نشطة
                </span>
                <span className="bg-emerald-500 text-white text-xs font-bold px-2 py-0.5 rounded-full">{activeItemsList.length}</span>
              </button>
              <button
                onClick={() => goToTable('cancelled')}
                className={`w-full flex items-center justify-between px-3 py-2 rounded-lg transition-colors ${
                  tableFilter === 'cancelled' && activeSection === 'table' ? 'bg-rose-100 ring-2 ring-rose-400' : 'bg-rose-50 hover:bg-rose-100'
                }`}
              >
                <span className="flex items-center gap-2 text-sm font-bold text-[var(--red)]">
                  <XCircle size={15} />
                  ملغاة
                </span>
                <span className="bg-[var(--red)] text-white text-xs font-bold px-2 py-0.5 rounded-full">{cancelledCount}</span>
              </button>
            </div>

            <p className="text-xs font-bold text-gray-400 mb-2">الأقسام</p>
            <div className="space-y-1">
              {SECTIONS.map((s) => {
                const Icon = s.icon
                const active = activeSection === s.id
                const count =
                  s.id === 'cancel' ? cancelledCount :
                  s.id === 'history' ? batches.length :
                  s.id === 'table' ? visibleItems.length :
                  s.id === 'activity' ? activityLogs.length :
                  s.id === 'messages' ? unreadFromBranches : null
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
                        s.id === 'messages' && unreadFromBranches > 0 ? 'bg-[var(--red)] text-white' : 'bg-gray-200 text-gray-600'
                      }`}>
                        {count}
                      </span>
                    )}
                  </button>
                )
              })}
            </div>
          </div>

          <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-2 shadow-sm">
            <button
              onClick={() => setActiveSection('settings')}
              className="w-full flex items-center gap-2 px-3 py-2.5 rounded-lg text-sm font-bold text-gray-600 hover:bg-gray-50 transition-colors"
            >
              <Settings size={15} />
              إعدادات الحساب
            </button>
            <button className="w-full flex items-center gap-2 px-3 py-2.5 rounded-lg text-sm font-bold text-[var(--red)] hover:bg-rose-50 transition-colors">
              <LogOut size={15} />
              تسجيل الخروج
            </button>
          </div>
        </aside>

        <div className="flex-1 min-w-0">
          {status && (
            <div className="mb-6 p-3.5 bg-[var(--yellow)]/15 border-2 border-[var(--yellow)]/40 rounded-lg text-sm text-[#8a6300] font-bold">
              {status}
            </div>
          )}

          {activeSection === 'stats' && (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
              <button onClick={() => goToTable('active')} className="text-right rounded-2xl p-6 bg-[var(--navy)] flex items-center justify-between shadow-lg shadow-[var(--navy)]/20 hover:opacity-90 transition-opacity">
                <div className="w-14 h-14 rounded-xl bg-white/15 flex items-center justify-center text-white shrink-0">
                  <Package size={24} />
                </div>
                <div className="text-left">
                  <p className="text-3xl font-black text-white">{activeItemsList.length}</p>
                  <p className="text-sm text-white/80 font-bold">عروض نشطة</p>
                </div>
              </button>

              <button onClick={() => goToTable('cancelled')} className="text-right rounded-2xl p-6 bg-[var(--red)] flex items-center justify-between shadow-lg shadow-[var(--red)]/20 hover:opacity-90 transition-opacity">
                <div className="w-14 h-14 rounded-xl bg-white/15 flex items-center justify-center text-white shrink-0">
                  <XCircle size={24} />
                </div>
                <div className="text-left">
                  <p className="text-3xl font-black text-white">{cancelledCount}</p>
                  <p className="text-sm text-white/80 font-bold">عروض ملغاة</p>
                </div>
              </button>

              <div className="rounded-2xl p-6 bg-[var(--yellow)] flex items-center justify-between shadow-lg shadow-[var(--yellow)]/30">
                <div className="w-14 h-14 rounded-xl bg-white/40 flex items-center justify-center text-[var(--navy)] shrink-0">
                  <TrendingUp size={24} />
                </div>
                <div className="text-left">
                  <p className="text-3xl font-black text-[var(--navy)]">{avgDiscount}%</p>
                  <p className="text-sm text-[var(--navy)]/80 font-bold">متوسط نسبة الخصم</p>
                </div>
              </div>
            </div>
          )}

          {activeSection === 'upload' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm max-w-xl">
              <div className="flex items-center justify-between mb-1">
                <h2 className="font-black text-sm text-[var(--navy)]">رفع تحديث عروض جديد</h2>
                <UploadCloud size={16} className="text-[var(--navy)]" />
              </div>
              <p className="text-xs text-gray-600 font-medium mb-3">باركود، اسم المنتج، السعر السابق، سعر العرض</p>

              {!pendingFile ? (
                <label className="flex flex-col items-center justify-center gap-2 border-2 border-dashed border-[var(--navy)]/25 rounded-xl p-6 cursor-pointer hover:border-[var(--navy)] hover:bg-[var(--navy)]/5 transition-colors">
                  <UploadCloud size={22} className="text-[var(--navy)]" />
                  <span className="text-[var(--navy)] font-bold text-sm">اختر ملف Excel</span>
                  <span className="mt-1 bg-[var(--navy)] text-white text-xs font-bold px-4 py-1.5 rounded-lg">
                    استعراض الملفات
                  </span>
                  <input type="file" accept=".xlsx,.xls" onChange={handleSelectUpdateFile} className="hidden" />
                </label>
              ) : (
                <div className="border-2 border-[var(--navy)]/30 rounded-xl p-4 bg-[var(--navy)]/5">
                  <div className="flex items-center gap-2 mb-3 text-[var(--navy)] text-sm font-bold">
                    <FileCheck size={16} />
                    {pendingFile.name}
                  </div>
                  <label className="block mb-2 text-xs text-gray-600 font-medium">اسم هذا التحديث</label>
                  <input
                    value={batchLabel}
                    onChange={(e) => setBatchLabel(e.target.value)}
                    className="w-full bg-white border-2 border-[var(--navy)]/20 rounded-lg p-2 text-sm text-[var(--navy)] font-medium mb-3 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                  />
                  <div className="flex gap-2">
                    <button
                      onClick={handleConfirmUpload}
                      disabled={uploading}
                      className="flex-1 bg-[var(--navy)] hover:bg-[#0f1a4d] text-white text-sm font-bold py-2 rounded-lg transition-colors disabled:opacity-50"
                    >
                      {uploading ? 'جاري الرفع...' : 'تأكيد ورفع التحديث'}
                    </button>
                    <button
                      onClick={() => { setPendingFile(null); setBatchLabel('') }}
                      className="px-4 bg-gray-100 hover:bg-gray-200 text-gray-600 text-sm font-bold rounded-lg transition-colors"
                    >
                      إلغاء
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {activeSection === 'label' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm max-w-xl">
              <div className="flex items-center justify-between mb-1">
                <h2 className="font-black text-sm text-[var(--navy)]">قالب الملصق</h2>
                <ImagePlus size={16} className="text-[var(--yellow)]" />
              </div>
              <p className="text-xs text-gray-600 font-medium mb-4">ملف PDF يُستخدم كقالب خلفية لكل ملصقات الفروع (جودة أعلى من الصور)</p>

              <div className="mb-4">
                <p className="text-xs font-bold text-gray-500 mb-2">المعاينة الحالية</p>
                {labelExists && labelPreviewUrl ? (
                  <div className="border-2 border-[var(--navy)]/15 rounded-xl overflow-hidden bg-gray-50 flex justify-center p-4">
                    <img src={labelPreviewUrl} alt="قالب الملصق الحالي" className="max-h-80 object-contain rounded-lg" />
                  </div>
                ) : (
                  <div className="border-2 border-dashed border-gray-300 rounded-xl p-8 text-center text-gray-400 text-sm font-medium">
                    لا يوجد قالب مرفوع حالياً
                  </div>
                )}
              </div>

              <div className="flex flex-col sm:flex-row gap-3">
                <label className="flex-1 flex items-center justify-center gap-2 border-2 border-dashed border-[var(--yellow)]/50 rounded-xl p-4 cursor-pointer hover:border-[var(--yellow)] hover:bg-[var(--yellow)]/10 transition-colors">
                  <ImagePlus size={18} className="text-[#8a6300]" />
                  <span className="text-[var(--navy)] font-bold text-sm">
                    {uploadingBg ? 'جاري الرفع...' : labelExists ? 'استبدال القالب' : 'رفع قالب جديد'}
                  </span>
                  <input type="file" accept=".pdf,application/pdf" onChange={handleUploadBackground} className="hidden" disabled={uploadingBg} />
                </label>
                {labelExists && (
                  <button
                    onClick={handleDeleteBackground}
                    className="flex items-center justify-center gap-2 bg-white border-2 border-[var(--red)]/30 hover:bg-[var(--red)]/5 text-[var(--red)] px-4 py-2.5 rounded-lg text-sm font-bold transition-colors"
                  >
                    <Trash2 size={15} />
                    حذف القالب
                  </button>
                )}
              </div>
              <p className="text-[11px] text-gray-400 font-medium mt-2">صيغة مقبولة: PDF بجودة عالية (صفحة واحدة تُستخدم)</p>
            </div>
          )}

          {activeSection === 'cancel' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--red)]/30 p-6 shadow-sm">
              <div className="flex items-center justify-between mb-1">
                <h2 className="font-black text-sm text-[var(--navy)]">إدارة العروض الملغاة</h2>
                <XCircle size={16} className="text-[var(--red)]" />
              </div>
              <p className="text-xs text-gray-600 font-medium mb-3">ملف فيه عمود باركود فقط للمنتجات اللي يُلغى عرضها</p>

              <div className="flex flex-col md:flex-row gap-3 items-start">
                {!pendingCancelFile ? (
                  <label className="flex items-center justify-center gap-3 border-2 border-dashed border-[var(--red)]/30 rounded-xl p-5 cursor-pointer hover:border-[var(--red)] hover:bg-[var(--red)]/5 transition-colors w-full md:w-1/2">
                    <XCircle size={20} className="text-[var(--red)]" />
                    <span className="text-[var(--navy)] font-bold text-sm">اختر ملف الإلغاء (Excel)</span>
                    <input type="file" accept=".xlsx,.xls" onChange={handleSelectCancelFile} className="hidden" />
                  </label>
                ) : (
                  <div className="border-2 border-[var(--red)]/40 rounded-xl p-4 bg-[var(--red)]/5 w-full md:w-1/2">
                    <div className="flex items-center gap-2 mb-3 text-[var(--red)] text-sm font-bold">
                      <FileCheck size={16} />
                      {pendingCancelFile.name}
                    </div>
                    <label className="block mb-2 text-xs text-gray-600 font-medium">اسم هذا الإلغاء</label>
                    <input
                      value={cancelBatchLabel}
                      onChange={(e) => setCancelBatchLabel(e.target.value)}
                      className="w-full bg-white border-2 border-[var(--red)]/20 rounded-lg p-2 text-sm text-[var(--navy)] font-medium mb-3 focus:outline-none focus:ring-2 focus:ring-[var(--red)]/20"
                    />
                    <div className="flex gap-2">
                      <button
                        onClick={handleConfirmCancelUpload}
                        disabled={uploadingCancel}
                        className="flex-1 bg-[var(--red)] hover:bg-[#c11a20] text-white text-sm font-bold py-2 rounded-lg transition-colors disabled:opacity-50"
                      >
                        {uploadingCancel ? 'جاري الإلغاء...' : 'تأكيد الإلغاء'}
                      </button>
                      <button
                        onClick={() => { setPendingCancelFile(null); setCancelBatchLabel('') }}
                        className="px-4 bg-gray-100 hover:bg-gray-200 text-gray-600 text-sm font-bold rounded-lg transition-colors"
                      >
                        إلغاء
                      </button>
                    </div>
                  </div>
                )}

                <button
                  onClick={handleDownloadCancelled}
                  className="flex items-center justify-center gap-2 bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 text-[var(--navy)] px-4 py-2.5 rounded-lg text-sm font-bold transition-colors whitespace-nowrap"
                >
                  <Download size={15} />
                  تحميل إكسل الملغاة ({cancelledCount})
                </button>
              </div>
            </div>
          )}

          {activeSection === 'history' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
              <div className="p-4 border-b-2 border-[var(--navy)]/10 flex items-center gap-2 bg-[var(--navy)]/5">
                <History size={16} className="text-[var(--navy)]" />
                <h2 className="font-black text-sm text-[var(--navy)]">سجل التحديثات ({batches.length})</h2>
              </div>
              <div className="divide-y-2 divide-[var(--navy)]/10 max-h-[700px] overflow-y-auto">
                {batches.length === 0 && (
                  <p className="p-6 text-center text-gray-500 text-sm font-medium">ماله تحديثات لسا</p>
                )}
                {batches.map((batch) => {
                  const batchItems = itemsForBatch(batch)
                  const isOpen = expandedBatchId === batch.id
                  const isCancel = batch.batch_type === 'cancel'
                  const respondedMap = isCancel ? removalMap : activationMap
                  const respondedSet = respondedMap[batch.id] || new Set()
                  return (
                    <div key={batch.id}>
                      <button
                        onClick={() => setExpandedBatchId(isOpen ? null : batch.id)}
                        className="w-full flex items-center justify-between p-4 hover:bg-[var(--navy)]/5 transition-colors text-right"
                      >
                        <div className="flex items-center gap-3">
                          <span className={`w-2.5 h-2.5 rounded-full ${isCancel ? 'bg-[var(--red)]' : 'bg-emerald-500'}`} />
                          <div>
                            <p className="text-sm font-bold text-[var(--navy)]">{batch.label}</p>
                            <p className="text-[11px] text-gray-500 font-medium">{formatDate(batch.created_at)}</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-3">
                          <span className={`text-xs font-bold px-2.5 py-1 rounded-md ${isCancel ? 'bg-[var(--red)]/10 text-[var(--red)]' : 'bg-emerald-100 text-emerald-700'}`}>
                            {isCancel ? 'إلغاء' : 'جديد'} · {batchItems.length}
                          </span>
                          <span className="text-xs font-bold px-2.5 py-1 rounded-md bg-[var(--navy)]/10 text-[var(--navy)]">
                            {respondedSet.size}/{branches.length} استجاب
                          </span>
                          <button
                            onClick={(e) => { e.stopPropagation(); handleDeleteBatch(batch) }}
                            className="text-gray-400 hover:text-[var(--red)] transition-colors p-1"
                            title="حذف هذا التحديث بالكامل"
                          >
                            <Trash2 size={15} />
                          </button>
                          {isOpen ? <ChevronUp size={16} className="text-gray-400" /> : <ChevronDown size={16} className="text-gray-400" />}
                        </div>
                      </button>
                      {isOpen && (
                        <div className="bg-[var(--navy)]/5 px-4 pb-4 space-y-4">
                          <div>
                            <p className="text-xs font-bold text-gray-500 mb-2">المنتجات ({batchItems.length})</p>
                            <div className="max-h-40 overflow-y-auto space-y-1.5">
                              {batchItems.map((item) => (
                                <div key={item.id} className="flex items-center justify-between text-xs bg-white border border-[var(--navy)]/10 rounded-lg px-3 py-2">
                                  <span className="text-[var(--navy)] font-bold">{item.product_name}</span>
                                  <span className="text-gray-500">{item.barcode}</span>
                                </div>
                              ))}
                              {batchItems.length === 0 && (
                                <p className="text-gray-400 text-xs text-center py-2">ما فيه منتجات مرتبطة</p>
                              )}
                            </div>
                          </div>

                          <div>
                            <p className="text-xs font-bold text-gray-500 mb-2">
                              {isCancel ? 'حالة الفروع (إزالة الملصقات)' : 'حالة الفروع (تفعيل التحديث)'}
                            </p>
                            <div className="grid sm:grid-cols-2 gap-1.5 max-h-56 overflow-y-auto">
                              {branches.map((b) => {
                                const done = respondedSet.has(b.id)
                                return (
                                  <div
                                    key={b.id}
                                    className={`flex items-center justify-between text-xs rounded-lg px-3 py-2 border ${
                                      done ? 'bg-emerald-50 border-emerald-200' : 'bg-white border-gray-200'
                                    }`}
                                  >
                                    <span className="flex items-center gap-1.5 text-[var(--navy)] font-bold">
                                      <Store size={12} />
                                      {b.name}
                                    </span>
                                    {done ? (
                                      <span className="flex items-center gap-1 text-emerald-600 font-bold">
                                        <CheckCircle2 size={12} />
                                        {isCancel ? 'أزال' : 'فعّل'}
                                      </span>
                                    ) : (
                                      <span className="text-gray-400 font-bold">لسه ما استجاب</span>
                                    )}
                                  </div>
                                )
                              })}
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            </div>
          )}

          {activeSection === 'manual' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm">
              <div className="flex items-center justify-between mb-4">
                <h2 className="font-black text-sm text-[var(--navy)]">إضافة منتج يدوياً</h2>
                <PlusCircle size={16} className="text-[var(--navy)]" />
              </div>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                <input value={newBarcode} onChange={(e) => setNewBarcode(e.target.value)} placeholder="الباركود"
                  className="bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium placeholder:text-gray-400 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20" />
                <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="اسم المنتج"
                  className="bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium placeholder:text-gray-400 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20" />
                <div className="relative">
                  <span className="absolute right-3 top-1/2 -translate-y-1/2 text-xs font-bold text-gray-400">SR</span>
                  <input value={newPrevPrice} onChange={(e) => setNewPrevPrice(e.target.value)} placeholder="السعر السابق" type="number"
                    className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 pr-10 text-sm text-[var(--navy)] font-medium placeholder:text-gray-400 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20" />
                </div>
                <div className="relative">
                  <span className="absolute right-3 top-1/2 -translate-y-1/2 text-xs font-bold text-gray-400">SR</span>
                  <input value={newOfferPrice} onChange={(e) => setNewOfferPrice(e.target.value)} placeholder="سعر العرض" type="number"
                    className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 pr-10 text-sm text-[var(--navy)] font-medium placeholder:text-gray-400 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20" />
                </div>
              </div>
              <button onClick={handleAddManual}
                className="mt-4 bg-[var(--navy)] hover:bg-[#0f1a4d] text-white px-6 py-2.5 rounded-lg text-sm font-bold transition-colors flex items-center gap-2">
                <PlusCircle size={16} />
                إضافة المنتج
              </button>
            </div>
          )}

          {activeSection === 'messages' && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm">
                <div className="flex items-center gap-2 mb-4">
                  <Send size={16} className="text-[var(--navy)]" />
                  <h2 className="font-black text-sm text-[var(--navy)]">إرسال رسالة</h2>
                </div>
                <label className="block text-xs font-bold text-gray-500 mb-1">إرسال إلى</label>
                <select
                  value={messageTarget}
                  onChange={(e) => setMessageTarget(e.target.value)}
                  className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-bold mb-3 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                >
                  <option value="">كل الفروع</option>
                  {branches.map((b) => (
                    <option key={b.id} value={b.id}>{b.name}</option>
                  ))}
                </select>
                <textarea
                  value={newMessageText}
                  onChange={(e) => setNewMessageText(e.target.value)}
                  placeholder="اكتب رسالتك هنا..."
                  rows={4}
                  className="w-full bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2.5 text-sm text-[var(--navy)] font-medium mb-3 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20 resize-none"
                />
                <button
                  onClick={handleSendMessage}
                  disabled={sendingMessage || !newMessageText.trim()}
                  className="w-full bg-[var(--navy)] hover:bg-[#0f1a4d] text-white px-5 py-2.5 rounded-lg text-sm font-bold transition-colors flex items-center justify-center gap-2 disabled:opacity-50"
                >
                  <Send size={15} />
                  {sendingMessage ? 'جاري الإرسال...' : 'إرسال'}
                </button>
              </div>

              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
                <div className="p-4 border-b-2 border-[var(--navy)]/10 bg-[var(--navy)]/5">
                  <h2 className="font-black text-sm text-[var(--navy)]">سجل الرسائل</h2>
                </div>
                <div className="divide-y divide-[var(--navy)]/10 max-h-[500px] overflow-y-auto">
                  {messages.length === 0 && (
                    <p className="p-6 text-center text-gray-400 text-sm">ما فيه رسائل بعد</p>
                  )}
                  {messages.map((m) => (
                    <div key={m.id} className={`p-4 ${m.sender_role === 'branch' ? 'bg-[var(--yellow)]/5' : ''}`}>
                      <p className="text-xs font-bold text-gray-500 mb-1">
                        {m.sender_role === 'admin'
                          ? `أنت → ${m.target_branch_id ? branchName(m.target_branch_id) : 'كل الفروع'}`
                          : `${branchName(m.sender_branch_id) || 'فرع'} → أنت`}
                        <span className="font-normal"> · {formatDate(m.created_at)}</span>
                      </p>
                      <p className="text-sm text-[var(--navy)] font-medium">{m.body}</p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {activeSection === 'whatsapp' && <WhatsAppNotifySection />}

          {activeSection === 'accounts' && <BranchAccountsSection />}

          {activeSection === 'audit' && <AuditOverviewSection />}

          {activeSection === 'settings' && (
            <div className="space-y-5 max-w-2xl">
              <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 p-6 shadow-sm">
                <h2 className="font-black text-sm text-[var(--navy)] mb-1">إعدادات الحساب</h2>
                <p className="text-xs text-gray-500 font-medium">إعدادات عامة للنظام</p>
              </div>

              <div className="bg-[var(--red)]/5 rounded-2xl border-2 border-[var(--red)]/30 p-6 shadow-sm">
                <h2 className="font-black text-sm text-[var(--red)] mb-1 flex items-center gap-2">
                  <XCircle size={16} />
                  منطقة الخطر
                </h2>
                <p className="text-xs text-gray-600 font-medium mb-4">
                  تهيئة كاملة: يمسح كل المنتجات، التحديثات، الإلغاءات، تأكيدات الفروع، باركودات الفروع، الرسائل، سجل النشاط، بيانات التدقيق، وقالب الملصق — نهائياً بدون رجعة.
                  <br />
                  <strong>ما يتأثر:</strong> أسماء الفروع وحساباتها (كلمات السر) تبقى زي ما هي.
                </p>
                <button
                  onClick={() => setShowFactoryResetModal(true)}
                  className="bg-[var(--red)] hover:bg-[#c11a20] text-white font-black px-6 py-3 rounded-xl transition-colors"
                >
                  تهيئة كاملة (حذف كل شي)
                </button>
              </div>
            </div>
          )}

          {activeSection === 'activity' && (
            <div className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm">
              <div className="p-4 border-b-2 border-[var(--navy)]/10 flex items-center gap-2 bg-[var(--navy)]/5">
                <History size={16} className="text-[var(--navy)]" />
                <h2 className="font-black text-sm text-[var(--navy)]">سجل النشاط ({activityLogs.length})</h2>
              </div>
              <div className="divide-y-2 divide-[var(--navy)]/10 max-h-[700px] overflow-y-auto">
                {activityLogs.length === 0 && (
                  <p className="p-6 text-center text-gray-500 text-sm font-medium">ما فيه أي نشاط مسجل بعد</p>
                )}
                {activityLogs.map((log) => {
                  const isAdmin = log.actor_role === 'admin'
                  const actorName = isAdmin ? 'الإدارة' : (branchName(log.branch_id) || 'فرع غير معروف')
                  return (
                    <div key={log.id} className="flex items-center justify-between gap-3 p-4">
                      <div className="flex items-center gap-3 min-w-0">
                        <span className={`w-2.5 h-2.5 rounded-full shrink-0 ${isAdmin ? 'bg-[var(--navy)]' : 'bg-emerald-500'}`} />
                        <div className="min-w-0">
                          <p className="text-sm font-bold text-[var(--navy)] truncate">
                            <span className={isAdmin ? 'text-[var(--navy)]' : 'text-emerald-700'}>{actorName}</span>
                            {' '}{log.action}
                            {log.details && <span className="text-gray-500 font-medium"> — {log.details}</span>}
                          </p>
                          <p className="text-[11px] text-gray-500 font-medium">{formatDate(log.created_at)}</p>
                        </div>
                      </div>
                    </div>
                  )
                })}
              </div>
            </div>
          )}

          {activeSection === 'table' && (
            <div
              onWheel={handleTableWheel}
              className="bg-[var(--card)] rounded-2xl border-2 border-[var(--navy)]/15 overflow-hidden shadow-sm"
            >
              <div className="p-4 border-b-2 border-[var(--navy)]/10 flex flex-wrap items-center justify-between gap-3 bg-[var(--navy)]/5">
                <h2 className="font-black text-sm text-[var(--navy)] flex items-center gap-2">
                  <Package size={16} />
                  جدول المنتجات ({searchedItems.length})
                </h2>
                <div className="flex items-center gap-3 flex-wrap">
                  <div className="relative">
                    <Search size={14} className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input
                      value={tableSearchQuery}
                      onChange={(e) => setTableSearchQuery(e.target.value)}
                      placeholder="ابحث بالاسم أو الباركود"
                      className="bg-white border-2 border-[var(--navy)]/15 rounded-lg p-2 pr-8 text-xs text-[var(--navy)] font-medium w-48 focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                    />
                  </div>
                  {selectedIds.size > 0 && (
                    <button
                      onClick={handleBulkDelete}
                      className="flex items-center gap-1.5 bg-[var(--red)] hover:bg-[#c11a20] text-white text-xs font-bold px-3 py-2 rounded-lg transition-colors"
                    >
                      <Trash2 size={13} />
                      حذف المحدد ({selectedIds.size})
                    </button>
                  )}
                  <div className="flex gap-1 bg-white rounded-lg border-2 border-[var(--navy)]/15 p-1">
                    <button
                      onClick={() => setTableFilter('active')}
                      className={`px-3 py-1.5 rounded-md text-xs font-bold transition-colors ${tableFilter === 'active' ? 'bg-emerald-500 text-white' : 'text-gray-500 hover:bg-gray-50'}`}
                    >
                      نشطة
                    </button>
                    <button
                      onClick={() => setTableFilter('cancelled')}
                      className={`px-3 py-1.5 rounded-md text-xs font-bold transition-colors ${tableFilter === 'cancelled' ? 'bg-[var(--red)] text-white' : 'text-gray-500 hover:bg-gray-50'}`}
                    >
                      ملغاة
                    </button>
                    <button
                      onClick={() => setTableFilter('all')}
                      className={`px-3 py-1.5 rounded-md text-xs font-bold transition-colors ${tableFilter === 'all' ? 'bg-[var(--navy)] text-white' : 'text-gray-500 hover:bg-gray-50'}`}
                    >
                      الكل
                    </button>
                  </div>
                </div>
              </div>
              <div className="overflow-x-auto">
                {paginatedItems.length > 0 && paginatedItems.every((i) => selectedIds.has(i.id!)) && selectedIds.size < searchedItems.length && (
                  <div className="p-2.5 bg-[var(--navy)]/5 border-b-2 border-[var(--navy)]/10 text-center text-xs font-bold text-[var(--navy)]">
                    تم تحديد {selectedIds.size} منتج بهذي الصفحة.{' '}
                    <button
                      onClick={() => setSelectedIds(new Set(searchedItems.map((i) => i.id!).filter(Boolean)))}
                      className="text-[var(--red)] underline hover:no-underline"
                    >
                      تحديد كل الـ{searchedItems.length} منتج المطابقة للفلتر
                    </button>
                  </div>
                )}
                {selectedIds.size === searchedItems.length && searchedItems.length > 0 && (
                  <div className="p-2.5 bg-[var(--red)]/5 border-b-2 border-[var(--red)]/20 text-center text-xs font-bold text-[var(--red)]">
                    تم تحديد كل الـ{searchedItems.length} منتج.{' '}
                    <button
                      onClick={() => setSelectedIds(new Set())}
                      className="underline hover:no-underline"
                    >
                      إلغاء التحديد
                    </button>
                  </div>
                )}
                <table className="w-full text-sm border-collapse">
                  <thead className="bg-[var(--navy)] text-white">
                    <tr>
                      <th className="p-3 text-center font-bold border-2 border-white/20 w-10">
                        <input
                          type="checkbox"
                          checked={paginatedItems.length > 0 && paginatedItems.every((i) => selectedIds.has(i.id!))}
                          onChange={toggleSelectAllOnPage}
                          className="w-4 h-4 cursor-pointer"
                        />
                      </th>
                      <th className="p-3 text-right font-bold border-2 border-white/20">الباركود</th>
                      <th className="p-3 text-right font-bold border-2 border-white/20">اسم المنتج</th>
                      <th className="p-3 text-right font-bold border-2 border-white/20">السعر السابق</th>
                      <th className="p-3 text-right font-bold border-2 border-white/20">سعر العرض</th>
                      <th className="p-3 text-center font-bold border-2 border-white/20">الخصم</th>
                      <th className="p-3 text-center font-bold border-2 border-white/20">الحالة</th>
                      <th className="p-3 text-center font-bold border-2 border-white/20">حذف</th>
                    </tr>
                  </thead>
                  <tbody>
                    {paginatedItems.map((item, i) => {
                      const discount = item.previous_price
                        ? Math.round((1 - item.offer_price / item.previous_price) * 100)
                        : 0
                      const cancelled = item.is_active === false
                      const isEditing = editingId === item.id
                      return (
                        <tr key={item.id} className={`${i % 2 === 0 ? 'bg-white' : 'bg-[var(--navy)]/[0.03]'} hover:bg-[var(--yellow)]/10 transition-colors ${cancelled ? 'opacity-50' : ''} ${selectedIds.has(item.id!) ? 'bg-[var(--navy)]/10' : ''} ${isEditing ? 'bg-[var(--yellow)]/10' : ''}`}>
                          <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                            <input
                              type="checkbox"
                              checked={selectedIds.has(item.id!)}
                              onChange={() => toggleSelectItem(item.id!)}
                              className="w-4 h-4 cursor-pointer"
                            />
                          </td>
                          <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">{item.barcode}</td>
                          <td className="p-3 text-[var(--navy)] font-bold border-2 border-[var(--navy)]/10">
                            {isEditing ? (
                              <input
                                value={editName}
                                onChange={(e) => setEditName(e.target.value)}
                                className="w-full min-w-[160px] bg-white border-2 border-[var(--navy)]/20 rounded-lg p-1.5 text-sm text-[var(--navy)] font-bold focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                              />
                            ) : (
                              item.product_name
                            )}
                          </td>
                          <td className="p-3 text-gray-500 font-bold border-2 border-[var(--navy)]/10">
                            {isEditing ? (
                              <input
                                type="number"
                                value={editPrevPrice}
                                onChange={(e) => setEditPrevPrice(e.target.value)}
                                className="w-24 bg-white border-2 border-[var(--navy)]/20 rounded-lg p-1.5 text-sm text-[var(--navy)] font-bold focus:outline-none focus:ring-2 focus:ring-[var(--navy)]/20"
                              />
                            ) : (
                              <span className="line-through">{item.previous_price.toFixed(2)}</span>
                            )}
                          </td>
                          <td className="p-3 text-[var(--red)] font-black border-2 border-[var(--navy)]/10">
                            {isEditing ? (
                              <input
                                type="number"
                                value={editOfferPrice}
                                onChange={(e) => setEditOfferPrice(e.target.value)}
                                className="w-24 bg-white border-2 border-[var(--red)]/30 rounded-lg p-1.5 text-sm text-[var(--red)] font-black focus:outline-none focus:ring-2 focus:ring-[var(--red)]/20"
                              />
                            ) : (
                              item.offer_price.toFixed(2)
                            )}
                          </td>
                          <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                            <span className={`inline-block text-xs font-bold px-2.5 py-1 rounded-md ${discountBadgeClass(discount)}`}>
                              {discount}%-
                            </span>
                          </td>
                          <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                            {cancelled ? (
                              <button
                                onClick={() => handleReactivate(item.id!)}
                                className="text-xs font-bold text-[var(--red)] hover:text-emerald-600 transition-colors"
                              >
                                ملغى (إعادة تفعيل)
                              </button>
                            ) : (
                              <span className="text-xs font-bold text-emerald-600">نشط</span>
                            )}
                          </td>
                          <td className="p-3 text-center border-2 border-[var(--navy)]/10">
                            <div className="flex items-center justify-center gap-2">
                              {isEditing ? (
                                <>
                                  <button
                                    onClick={() => handleSaveEdit(item.id!)}
                                    disabled={savingEdit}
                                    className="text-emerald-600 hover:text-emerald-700 transition-colors disabled:opacity-50"
                                    title="حفظ"
                                  >
                                    <Check size={16} />
                                  </button>
                                  <button
                                    onClick={handleCancelEdit}
                                    className="text-gray-400 hover:text-gray-600 transition-colors"
                                    title="إلغاء"
                                  >
                                    <X size={16} />
                                  </button>
                                </>
                              ) : (
                                <>
                                  <button
                                    onClick={() => handleStartEdit(item)}
                                    className="text-gray-400 hover:text-[var(--navy)] transition-colors"
                                    title="تعديل السعر"
                                  >
                                    <Pencil size={15} />
                                  </button>
                                  <button onClick={() => handleDelete(item.id!)} className="text-gray-400 hover:text-[var(--red)] transition-colors" title="حذف">
                                    <Trash2 size={16} />
                                  </button>
                                </>
                              )}
                            </div>
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>

              <div className="p-4 border-t-2 border-[var(--navy)]/10 flex flex-col sm:flex-row items-center justify-between gap-3 bg-[var(--navy)]/5">
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setPage((p) => Math.max(1, p - 1))}
                    disabled={page === 1}
                    className="w-8 h-8 rounded-lg bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 disabled:opacity-30 flex items-center justify-center transition-colors"
                  >
                    <ChevronRight size={15} className="text-[var(--navy)]" />
                  </button>
                  {Array.from({ length: totalPages }, (_, i) => i + 1)
                    .filter((p) => p === 1 || p === totalPages || Math.abs(p - page) <= 1)
                    .map((p, idx, arr) => (
                      <span key={p} className="flex items-center gap-2">
                        {idx > 0 && arr[idx - 1] !== p - 1 && <span className="text-gray-400">...</span>}
                        <button
                          onClick={() => setPage(p)}
                          className={`w-8 h-8 rounded-lg text-sm font-semibold transition-colors ${
                            p === page ? 'bg-[var(--navy)] text-white' : 'bg-white border-2 border-[var(--navy)]/15 text-[var(--navy)] hover:bg-[var(--navy)]/10'
                          }`}
                        >
                          {p}
                        </button>
                      </span>
                    ))}
                  <button
                    onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                    disabled={page === totalPages}
                    className="w-8 h-8 rounded-lg bg-white border-2 border-[var(--navy)]/15 hover:bg-[var(--navy)]/10 disabled:opacity-30 flex items-center justify-center transition-colors"
                  >
                    <ChevronLeft size={15} className="text-[var(--navy)]" />
                  </button>
                </div>
                <div className="flex items-center gap-2 text-xs text-gray-600 font-bold">
                  <span>
                    عرض {(page - 1) * pageSize + 1} - {Math.min(page * pageSize, searchedItems.length)} من {searchedItems.length} منتج
                  </span>
                  <select
                    value={pageSize}
                    onChange={(e) => setPageSize(Number(e.target.value))}
                    className="bg-white border-2 border-[var(--navy)]/15 rounded-lg px-2 py-1 text-[var(--navy)] text-xs font-bold focus:outline-none"
                  >
                    {PAGE_SIZE_OPTIONS.map((n) => (
                      <option key={n} value={n}>{n}</option>
                    ))}
                  </select>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      {confirmDialog && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl p-6 max-w-md w-full shadow-2xl">
            <p className="text-sm text-[var(--navy)] font-bold mb-6 whitespace-pre-line leading-relaxed">
              {confirmDialog.message}
            </p>
            <div className="flex gap-3">
              <button
                onClick={() => {
                  confirmDialog.onConfirm()
                  setConfirmDialog(null)
                }}
                className="flex-1 bg-[var(--red)] hover:bg-[#c11a20] text-white text-sm font-bold py-2.5 rounded-lg transition-colors"
              >
                تأكيد الحذف
              </button>
              <button
                onClick={() => setConfirmDialog(null)}
                className="flex-1 bg-gray-100 hover:bg-gray-200 text-gray-600 text-sm font-bold py-2.5 rounded-lg transition-colors"
              >
                إلغاء
              </button>
            </div>
          </div>
        </div>
      )}

      {showFactoryResetModal && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl p-6 max-w-md w-full shadow-2xl border-4 border-[var(--red)]">
            <h3 className="text-[var(--red)] font-black text-base mb-2 flex items-center gap-2">
              <XCircle size={20} />
              تحذير: تهيئة كاملة
            </h3>
            <p className="text-sm text-gray-700 font-medium mb-4 leading-relaxed">
              هذا الإجراء يمسح كل المنتجات، التحديثات، الإلغاءات، تأكيدات الفروع، باركوداتهم، الرسائل، سجل النشاط، بيانات التدقيق، وقالب الملصق — <strong>نهائياً وبدون رجعة</strong>.
              <br /><br />
              اكتب كلمة <strong className="text-[var(--red)]">تهيئة</strong> بالخانة تحت عشان تأكد:
            </p>
            <input
              value={factoryResetInput}
              onChange={(e) => setFactoryResetInput(e.target.value)}
              placeholder="اكتب: تهيئة"
              className="w-full bg-white border-2 border-[var(--red)]/30 rounded-lg p-2.5 text-sm text-[var(--navy)] font-bold mb-4 focus:outline-none focus:ring-2 focus:ring-[var(--red)]/30"
            />
            <div className="flex gap-3">
              <button
                onClick={handleFactoryReset}
                disabled={factoryResetInput !== 'تهيئة' || resettingFactory}
                className="flex-1 bg-[var(--red)] hover:bg-[#c11a20] text-white text-sm font-black py-2.5 rounded-lg transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
              >
                {resettingFactory ? 'جاري التهيئة...' : 'نعم، امسح كل شي'}
              </button>
              <button
                onClick={() => { setShowFactoryResetModal(false); setFactoryResetInput('') }}
                className="flex-1 bg-gray-100 hover:bg-gray-200 text-gray-600 text-sm font-bold py-2.5 rounded-lg transition-colors"
              >
                إلغاء
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
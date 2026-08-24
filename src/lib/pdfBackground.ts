import * as pdfjsLib from 'pdfjs-dist'

// يشغّل الـ worker من CDN عشان ما نحتاج نعدّل إعدادات البناء (webpack/turbopack)
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'

/**
 * يفتح أول صفحة من ملف PDF ويرسمها على Canvas بجودة تناسب المقاس المطلوب.
 * الـ Canvas الناتج يُستخدم مباشرة كخلفية للملصق (ctx.drawImage يقبل Canvas زي ما يقبل Image).
 */
export async function renderPdfToCanvas(
  url: string,
  targetWidth: number,
  targetHeight: number
): Promise<HTMLCanvasElement> {
  const pdf = await pdfjsLib.getDocument({ url }).promise
  const page = await pdf.getPage(1)
  const baseViewport = page.getViewport({ scale: 1 })
  // نكبّر لأكبر بعد مطلوب عشان الجودة تكفي حتى لو الملصق طلع أطول أو أعرض من نسبة الـ PDF
  const scale = Math.max(targetWidth / baseViewport.width, targetHeight / baseViewport.height)
  const viewport = page.getViewport({ scale })

  const canvas = document.createElement('canvas')
  canvas.width = viewport.width
  canvas.height = viewport.height
  const ctx = canvas.getContext('2d')!

  await page.render({ canvas, canvasContext: ctx, viewport }).promise
  return canvas
}
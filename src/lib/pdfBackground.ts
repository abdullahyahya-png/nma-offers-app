import * as pdfjsLib from 'pdfjs-dist'

// نجيب ملف الـ Worker من نفس نسخة pdfjs-dist المثبتة بالمشروع (بدل رابط CDN ثابت)
// عشان نضمن تطابق النسخة دايماً، حتى لو تحدثت المكتبة مستقبلاً
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/build/pdf.worker.min.mjs',
  import.meta.url
).toString()

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
  const scale = Math.max(targetWidth / baseViewport.width, targetHeight / baseViewport.height)
  const viewport = page.getViewport({ scale })

  const canvas = document.createElement('canvas')
  canvas.width = viewport.width
  canvas.height = viewport.height
  const ctx = canvas.getContext('2d')!

  await page.render({ canvas, canvasContext: ctx, viewport }).promise
  return canvas
}
// نأخر استيراد pdfjs-dist لحد ما الدالة تتنفذ فعلياً بالمتصفح (Client)
// عشان ما يصطدم بعملية بناء Next.js على السيرفر (اللي ما فيه DOMMatrix أو أي API متصفح)
let pdfjsLibPromise: Promise<typeof import('pdfjs-dist')> | null = null

async function getPdfjsLib() {
  if (!pdfjsLibPromise) {
    pdfjsLibPromise = import('pdfjs-dist').then((lib) => {
      lib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'
      return lib
    })
  }
  return pdfjsLibPromise
}

/**
 * يفتح أول صفحة من ملف PDF ويرسمها على Canvas بجودة تناسب المقاس المطلوب.
 * الـ Canvas الناتج يُستخدم مباشرة كخلفية للملصق (ctx.drawImage يقبل Canvas زي ما يقبل Image).
 */
export async function renderPdfToCanvas(
  url: string,
  targetWidth: number,
  targetHeight: number
): Promise<HTMLCanvasElement> {
  const pdfjsLib = await getPdfjsLib()
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
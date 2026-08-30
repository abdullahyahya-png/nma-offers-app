import type { Metadata } from 'next'

export const dynamic = 'force-dynamic'

export const metadata: Metadata = {
  title: 'واجهة الطباعة السريعة ومتابعة العروض',
  description: 'نظام إدارة العروض',
  manifest: '/print-manifest.json',
  openGraph: {
    title: 'واجهة الطباعة السريعة ومتابعة العروض',
    description: 'نظام إدارة العروض',
  },
}

export default function PrintLayout({ children }: { children: React.ReactNode }) {
  return children
}
import type { Metadata } from 'next'

export const dynamic = 'force-dynamic'

export const metadata: Metadata = {
  title: 'واجهة الطباعة السريعة - NMA',
}

export default function PrintLayout({ children }: { children: React.ReactNode }) {
  return children
}
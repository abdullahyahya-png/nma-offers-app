import type { Metadata } from 'next'

// هذا يخلي صفحة الفرع تتثبت كتطبيق منفصل عن لوحة الأدمن (أيقونة واسم مختلف)
export const metadata: Metadata = {
  title: 'بوابة الفروع - NMA',
  manifest: '/branch-manifest.json',
}

export default function BranchLayout({ children }: { children: React.ReactNode }) {
  return children
}
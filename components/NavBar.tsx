'use client'

import Link from 'next/link'
import { usePathname } from 'next/navigation'
import { usePermissions } from '@/lib/permissions'
import LockButton from './LockButton'

interface Props {
  onUnlock?: () => void
  onLock?: () => void
}

export default function NavBar({ onUnlock, onLock }: Props) {
  const pathname = usePathname()
  const { canAccess, isAdmin } = usePermissions()

  const cls = (href: string) =>
    pathname === href || (href !== '/' && pathname.startsWith(href + '/') && href.split('/').length === pathname.split('/').length)
      ? 'text-blue-600'
      : 'text-gray-500 hover:text-gray-800 transition-colors'

  // Summary dropdown active when on /summary or /qc/summary
  const summaryActive = pathname === '/summary' || pathname === '/qc/summary'
  const showSummaryDropdown = canAccess('summary') || canAccess('qc')

  return (
    <nav className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-6 text-sm shrink-0 sticky top-0 z-20 shadow-sm flex-wrap">
      <span className="font-bold text-gray-900">Import PO</span>

      {canAccess('po-matching') && (
        <Link href="/" className={pathname === '/' ? 'text-blue-600' : 'text-gray-500 hover:text-gray-800 transition-colors'}>
          PO Matching
        </Link>
      )}
      {canAccess('dashboard') && (
        <Link href="/dashboard" className={cls('/dashboard')}>Dashboard</Link>
      )}
      {canAccess('calendar') && (
        <Link href="/calendar" className={cls('/calendar')}>Calendar</Link>
      )}
      {canAccess('report') && (
        <Link href="/report" className={cls('/report')}>Report</Link>
      )}
      {canAccess('compare') && (
        <Link href="/compare" className={cls('/compare')}>Cost Compare</Link>
      )}
      {canAccess('po-builder') && (
        <Link href="/po-builder" className={cls('/po-builder')}>PO Builder</Link>
      )}
      {canAccess('order-plan') && (
        <Link href="/order-plan" className={cls('/order-plan')}>Order Plan</Link>
      )}

      {showSummaryDropdown && (
        <div className="relative group">
          <span className={`cursor-default ${summaryActive ? 'text-blue-600' : 'text-gray-500 hover:text-gray-800'}`}>
            Summary ▾
          </span>
          <div className="absolute left-0 top-full pt-1 hidden group-hover:block z-50">
            <div className="bg-white border border-gray-200 rounded-lg shadow-lg py-1 min-w-[150px]">
              {canAccess('summary') && (
                <Link href="/summary" className={`block px-4 py-2 text-sm hover:bg-blue-50 ${pathname === '/summary' ? 'text-blue-600' : 'text-gray-700'}`}>
                  Item Summary
                </Link>
              )}
              {canAccess('qc') && (
                <Link href="/qc/summary" className={`block px-4 py-2 text-sm hover:bg-blue-50 ${pathname === '/qc/summary' ? 'text-blue-600' : 'text-gray-700'}`}>
                  QC Summary
                </Link>
              )}
            </div>
          </div>
        </div>
      )}

      {canAccess('qc') && (
        <Link href="/qc" className={pathname === '/qc' ? 'text-blue-600' : 'text-gray-500 hover:text-gray-800 transition-colors'}>
          QC Report
        </Link>
      )}
      {canAccess('guide') && (
        <Link href="/guide" className={cls('/guide')}>Guide</Link>
      )}
      {isAdmin && (
        <Link href="/admin" className={cls('/admin')}>Admin</Link>
      )}

      <div className="ml-auto">
        <LockButton onUnlock={onUnlock} onLock={onLock} />
      </div>
    </nav>
  )
}

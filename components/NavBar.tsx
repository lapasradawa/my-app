'use client'

import Link from 'next/link'
import { usePathname } from 'next/navigation'
import { useEffect, useState } from 'react'
import { usePermissions } from '@/lib/permissions'
import LockButton from './LockButton'
import { supabase } from '@/lib/supabase'

interface Props {
  onUnlock?: () => void
  onLock?: () => void
}

export default function NavBar({ onUnlock, onLock }: Props) {
  const pathname = usePathname()
  const { canAccess, isAdmin } = usePermissions()
  const [pendingCount, setPendingCount] = useState(0)

  useEffect(() => {
    if (!isAdmin) return
    supabase.from('container_hub_requests').select('id', { count: 'exact', head: true }).eq('status', 'pending')
      .then(({ count }) => setPendingCount(count ?? 0))
  }, [isAdmin, pathname])

  const cls = (href: string) =>
    pathname === href || (href !== '/' && pathname.startsWith(href + '/') && href.split('/').length === pathname.split('/').length)
      ? 'text-blue-600'
      : 'text-gray-500 hover:text-gray-800 transition-colors'

  // Summary dropdown active when on /summary or /qc/summary
  const summaryActive = pathname === '/summary' || pathname === '/qc/summary' || pathname === '/po-summary'
  const summaryCount = [canAccess('summary'), canAccess('qc'), canAccess('po-summary')].filter(Boolean).length
  const showSummaryDropdown = summaryCount >= 2

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
        <Link href="/po-builder" className={cls('/po-builder')}>PO Insights</Link>
      )}
      {canAccess('order-plan') && (
        <Link href="/order-plan" className={cls('/order-plan')}>Order Plan</Link>
      )}
      {canAccess('load-plan') && (
        <Link href="/load-plan" className={cls('/load-plan')}>Branch Load</Link>
      )}

      {/* Direct links when user has access to only one summary page */}
      {!showSummaryDropdown && canAccess('summary') && (
        <Link href="/summary" className={cls('/summary')}>Invoice Summary</Link>
      )}
      {!showSummaryDropdown && canAccess('qc') && (
        <Link href="/qc/summary" className={pathname === '/qc/summary' ? 'text-blue-600' : 'text-gray-500 hover:text-gray-800 transition-colors'}>QC Summary</Link>
      )}
      {!showSummaryDropdown && canAccess('po-summary') && (
        <Link href="/po-summary" className={cls('/po-summary')}>PO Summary</Link>
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
                  Invoice Summary
                </Link>
              )}
              {canAccess('qc') && (
                <Link href="/qc/summary" className={`block px-4 py-2 text-sm hover:bg-blue-50 ${pathname === '/qc/summary' ? 'text-blue-600' : 'text-gray-700'}`}>
                  QC Summary
                </Link>
              )}
              {canAccess('po-summary') && (
                <Link href="/po-summary" className={`block px-4 py-2 text-sm hover:bg-blue-50 ${pathname === '/po-summary' ? 'text-blue-600' : 'text-gray-700'}`}>
                  PO Summary
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
        <Link href="/admin" className={`relative ${cls('/admin')}`}>
          Admin
          {pendingCount > 0 && (
            <span className="absolute -top-1.5 -right-3 bg-red-500 text-white text-[9px] font-bold rounded-full min-w-[16px] h-4 flex items-center justify-center px-1 leading-none">
              {pendingCount}
            </span>
          )}
        </Link>
      )}

      <div className="ml-auto">
        <LockButton onUnlock={onUnlock} onLock={onLock} />
      </div>
    </nav>
  )
}

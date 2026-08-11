'use client'

import { useEffect, useState } from 'react'
import { useRouter, usePathname } from 'next/navigation'
import { supabase } from '@/lib/supabase'
import { PermissionsContext, buildPermissions, pathnameToPageKey, type UserPermissions } from '@/lib/permissions'

const DEFAULT_PERMS: UserPermissions = buildPermissions(false, null)

export default function AuthGuard({ children }: { children: React.ReactNode }) {
  const [checking, setChecking] = useState(true)
  const [userEmail, setUserEmail] = useState<string | null>(null)
  const [permissions, setPermissions] = useState<UserPermissions>(DEFAULT_PERMS)
  const router = useRouter()
  const pathname = usePathname()

  async function fetchPermissions(email: string): Promise<UserPermissions> {
    const { data } = await supabase
      .from('page_permissions')
      .select('is_admin, allowed_pages')
      .eq('email', email)
      .single()
    // not in table → restricted (empty allowlist = only home page)
    if (!data) return buildPermissions(false, [])
    return buildPermissions(data.is_admin ?? false, data.allowed_pages ?? null)
  }

  useEffect(() => {
    supabase.auth.getSession().then(async ({ data: { session } }) => {
      if (!session && pathname !== '/login') {
        router.replace('/login')
        return
      }
      if (session?.user?.email) {
        const email = session.user.email
        setUserEmail(email)
        const perms = await fetchPermissions(email)
        setPermissions(perms)

        // Guard: check if current page is allowed
        const pageKey = pathnameToPageKey(pathname)
        if (pathname === '/admin' && !perms.isAdmin) {
          router.replace('/')
          return
        }
        if (pageKey && !perms.canAccess(pageKey)) {
          router.replace('/')
          return
        }
      }
      setChecking(false)
    })

    const { data: { subscription } } = supabase.auth.onAuthStateChange(async (_event, session) => {
      if (!session && pathname !== '/login') {
        router.replace('/login')
      } else if (session?.user?.email) {
        const email = session.user.email
        setUserEmail(email)
        const perms = await fetchPermissions(email)
        setPermissions(perms)
      }
    })

    return () => subscription.unsubscribe()
  }, [pathname, router])

  async function handleLogout() {
    await supabase.auth.signOut()
    router.replace('/login')
  }

  if (checking && pathname !== '/login') {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <p className="text-sm text-gray-400">กำลังตรวจสอบ...</p>
      </div>
    )
  }

  return (
    <PermissionsContext.Provider value={permissions}>
      {children}
      {userEmail && pathname !== '/login' && (
        <div className="fixed bottom-4 right-4 z-50 flex items-center gap-2 bg-white border border-gray-200 rounded-full px-3 py-1.5 shadow-sm text-xs text-gray-500">
          <span className="hidden sm:inline truncate max-w-[160px]">{userEmail}</span>
          <button
            onClick={handleLogout}
            className="text-gray-400 hover:text-red-500 transition-colors font-medium whitespace-nowrap"
          >
            ออกจากระบบ
          </button>
        </div>
      )}
    </PermissionsContext.Provider>
  )
}

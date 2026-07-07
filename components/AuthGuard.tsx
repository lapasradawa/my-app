'use client'

import { useEffect, useState } from 'react'
import { useRouter, usePathname } from 'next/navigation'
import { supabase } from '@/lib/supabase'

export default function AuthGuard({ children }: { children: React.ReactNode }) {
  const [checking, setChecking] = useState(true)
  const [userEmail, setUserEmail] = useState<string | null>(null)
  const router = useRouter()
  const pathname = usePathname()

  useEffect(() => {
    // Check current session
    supabase.auth.getSession().then(({ data: { session } }) => {
      if (!session && pathname !== '/login') {
        router.replace('/login')
      } else {
        setUserEmail(session?.user?.email ?? null)
        setChecking(false)
      }
    })

    // Listen for auth changes (logout from another tab, session expiry)
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      if (!session && pathname !== '/login') {
        router.replace('/login')
      } else {
        setUserEmail(session?.user?.email ?? null)
      }
    })

    return () => subscription.unsubscribe()
  }, [pathname, router])

  async function handleLogout() {
    await supabase.auth.signOut()
    router.replace('/login')
  }

  // Show nothing while checking auth (avoids flash of content)
  if (checking && pathname !== '/login') {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <p className="text-sm text-gray-400">กำลังตรวจสอบ...</p>
      </div>
    )
  }

  return (
    <>
      {children}
      {/* Logout button — only show when logged in and not on login page */}
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
    </>
  )
}

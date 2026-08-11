'use client'

import { createContext, useContext } from 'react'

export type PageKey =
  | 'po-matching'
  | 'dashboard'
  | 'calendar'
  | 'report'
  | 'compare'
  | 'po-builder'
  | 'order-plan'
  | 'summary'
  | 'qc'
  | 'guide'

export interface UserPermissions {
  isAdmin: boolean
  allowedPages: PageKey[] | null // null = no restriction (all pages)
  canAccess: (page: PageKey) => boolean
}

const defaultPermissions: UserPermissions = {
  isAdmin: false,
  allowedPages: null,
  canAccess: () => true,
}

export const PermissionsContext = createContext<UserPermissions>(defaultPermissions)

export function usePermissions(): UserPermissions {
  return useContext(PermissionsContext)
}

export function buildPermissions(isAdmin: boolean, allowedPages: string[] | null): UserPermissions {
  const pages = allowedPages as PageKey[] | null
  return {
    isAdmin,
    allowedPages: pages,
    canAccess: (page: PageKey) => {
      if (page === 'po-matching') return true
      if (isAdmin) return true
      if (pages === null) return true
      return pages.includes(page)
    },
  }
}

// Map pathname → PageKey
export function pathnameToPageKey(pathname: string): PageKey | null {
  const seg = pathname.split('/')[1] ?? ''
  switch (seg) {
    case '':           return 'po-matching'
    case 'dashboard':  return 'dashboard'
    case 'calendar':   return 'calendar'
    case 'report':     return 'report'
    case 'compare':    return 'compare'
    case 'po-builder': return 'po-builder'
    case 'order-plan': return 'order-plan'
    case 'summary':    return 'summary'
    case 'qc':         return 'qc'
    case 'guide':      return 'guide'
    case 'admin':      return null // handled separately
    case 'login':      return null // public
    default:           return null
  }
}

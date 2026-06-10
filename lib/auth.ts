const SESSION_KEY = 'edit_unlocked'
const PASSWORD = 'RBSCH'

export function isUnlocked(): boolean {
  if (typeof window === 'undefined') return false
  return sessionStorage.getItem(SESSION_KEY) === '1'
}

export function tryUnlock(password: string): boolean {
  if (password === PASSWORD) {
    sessionStorage.setItem(SESSION_KEY, '1')
    return true
  }
  return false
}

export function lock(): void {
  sessionStorage.removeItem(SESSION_KEY)
}

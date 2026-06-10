'use client'

import { useEffect, useState } from 'react'
import { isUnlocked, lock } from '@/lib/auth'
import PasswordModal from './PasswordModal'

interface Props {
  onUnlock?: () => void
  onLock?: () => void
}

export default function LockButton({ onUnlock, onLock }: Props) {
  const [unlocked, setUnlocked] = useState(false)
  const [showModal, setShowModal] = useState(false)

  useEffect(() => { setUnlocked(isUnlocked()) }, [])

  function handleClick() {
    if (unlocked) {
      lock()
      setUnlocked(false)
      onLock?.()
    } else {
      setShowModal(true)
    }
  }

  return (
    <>
      <button
        onClick={handleClick}
        className={`flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-full border transition-colors ${
          unlocked
            ? 'border-green-300 bg-green-50 text-green-700 hover:bg-green-100'
            : 'border-gray-300 bg-gray-50 text-gray-600 hover:bg-gray-100'
        }`}
      >
        <span>{unlocked ? '🔓' : '🔒'}</span>
        <span>{unlocked ? 'โหมดแก้ไข (คลิกล็อก)' : 'ล็อก — คลิกเพื่อแก้ไข'}</span>
      </button>
      {showModal && (
        <PasswordModal
          onSuccess={() => { setUnlocked(true); setShowModal(false); onUnlock?.() }}
          onCancel={() => setShowModal(false)}
        />
      )}
    </>
  )
}

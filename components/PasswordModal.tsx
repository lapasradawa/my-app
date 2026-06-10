'use client'

import { useRef, useState } from 'react'

interface Props {
  onSuccess: () => void
  onCancel: () => void
}

export default function PasswordModal({ onSuccess, onCancel }: Props) {
  const [value, setValue] = useState('')
  const [error, setError] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)

  function handleSubmit() {
    const { tryUnlock } = require('@/lib/auth')
    if (tryUnlock(value)) {
      onSuccess()
    } else {
      setError(true)
      setValue('')
      inputRef.current?.focus()
    }
  }

  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50" onClick={onCancel}>
      <div className="bg-white rounded-2xl shadow-xl p-6 w-80" onClick={e => e.stopPropagation()}>
        <h2 className="text-base font-semibold text-gray-800 mb-1">ใส่รหัสผ่านเพื่อแก้ไข</h2>
        <p className="text-xs text-gray-500 mb-4">ต้องใส่รหัสผ่านก่อนจึงจะแก้ไขข้อมูลได้</p>
        <input
          ref={inputRef}
          autoFocus
          type="password"
          value={value}
          onChange={e => { setValue(e.target.value); setError(false) }}
          onKeyDown={e => e.key === 'Enter' && handleSubmit()}
          placeholder="รหัสผ่าน"
          className={`w-full border rounded-lg px-3 py-2 text-sm outline-none mb-1 ${
            error ? 'border-red-400 focus:border-red-400' : 'border-gray-300 focus:border-blue-400'
          }`}
        />
        {error && <p className="text-xs text-red-500 mb-3">รหัสผ่านไม่ถูกต้อง</p>}
        {!error && <div className="mb-3" />}
        <div className="flex gap-2">
          <button
            onClick={onCancel}
            className="flex-1 py-2 text-sm text-gray-600 border border-gray-200 rounded-lg hover:bg-gray-50"
          >
            ยกเลิก
          </button>
          <button
            onClick={handleSubmit}
            className="flex-1 py-2 text-sm text-white bg-blue-600 rounded-lg hover:bg-blue-700"
          >
            ยืนยัน
          </button>
        </div>
      </div>
    </div>
  )
}

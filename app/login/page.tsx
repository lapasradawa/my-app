'use client'

import { useState, useEffect } from 'react'
import { useRouter } from 'next/navigation'
import { supabase } from '@/lib/supabase'

export default function LoginPage() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [newPassword, setNewPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const [mode, setMode] = useState<'login' | 'set-password'>('login')
  const router = useRouter()

  useEffect(() => {
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event) => {
      if (event === 'PASSWORD_RECOVERY') setMode('set-password')
    })
    return () => subscription.unsubscribe()
  }, [])

  async function handleLogin(e: React.FormEvent) {
    e.preventDefault()
    setLoading(true)
    setError('')
    const { error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) {
      setError('อีเมลหรือรหัสผ่านไม่ถูกต้อง')
    } else {
      router.replace('/')
    }
    setLoading(false)
  }

  async function handleSetPassword(e: React.FormEvent) {
    e.preventDefault()
    if (newPassword !== confirmPassword) { setError('รหัสผ่านไม่ตรงกัน'); return }
    if (newPassword.length < 6) { setError('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร'); return }
    setLoading(true)
    setError('')
    const { error } = await supabase.auth.updateUser({ password: newPassword })
    if (error) { setError('ตั้งรหัสผ่านไม่สำเร็จ: ' + error.message) }
    else { router.replace('/') }
    setLoading(false)
  }

  const inputStyle: React.CSSProperties = {
    width: '100%', border: '1.5px solid #e2d8c8', borderRadius: 10,
    padding: '11px 14px', fontSize: 13, outline: 'none',
    background: '#fff', color: '#3a2a1a', transition: 'border-color 0.15s',
    fontFamily: 'system-ui, sans-serif',
  }

  if (mode === 'set-password') {
    return (
      <div style={{ minHeight: '100vh', background: '#ede5d4', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, fontFamily: 'system-ui, sans-serif' }}>
        <div style={{ width: '100%', maxWidth: 380 }}>
          <div style={{ background: '#1e3340', borderRadius: '18px 18px 0 0', padding: '28px 32px 24px' }}>
            <div style={{ fontSize: 10, fontWeight: 800, color: '#3d8b82', textTransform: 'uppercase', letterSpacing: '0.14em', marginBottom: 6 }}>Global Sourcing</div>
            <div style={{ fontSize: 22, fontWeight: 900, color: '#d4962a' }}>ตั้งรหัสผ่าน</div>
            <div style={{ fontSize: 12, color: '#7a9aaa', marginTop: 4 }}>กรุณาตั้งรหัสผ่านสำหรับบัญชีของคุณ</div>
          </div>
          <div style={{ background: '#faf5ee', border: '1.5px solid #e2d8c8', borderTop: 'none', borderRadius: '0 0 18px 18px', padding: '28px 32px 32px' }}>
            <form onSubmit={handleSetPassword} style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <input type="password" value={newPassword} onChange={e => setNewPassword(e.target.value)}
                placeholder="รหัสผ่านใหม่ (อย่างน้อย 6 ตัว)" required autoFocus autoComplete="new-password" style={inputStyle} />
              <input type="password" value={confirmPassword} onChange={e => setConfirmPassword(e.target.value)}
                placeholder="ยืนยันรหัสผ่าน" required autoComplete="new-password" style={inputStyle} />
              {error && <p style={{ fontSize: 12, color: '#c85a3a', margin: 0 }}>{error}</p>}
              <button type="submit" disabled={loading} style={{
                width: '100%', background: loading ? '#a0c4c0' : '#3d8b82', color: '#fff',
                border: 'none', borderRadius: 10, padding: '12px 0', fontSize: 13,
                fontWeight: 700, cursor: loading ? 'not-allowed' : 'pointer', marginTop: 4,
                transition: 'background 0.15s',
              }}>
                {loading ? 'กำลังบันทึก...' : 'บันทึกรหัสผ่าน'}
              </button>
            </form>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div style={{ minHeight: '100vh', background: '#ede5d4', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16, fontFamily: 'system-ui, sans-serif' }}>
      {/* Decorative blobs */}
      <div style={{ position: 'fixed', top: -80, right: -80, width: 320, height: 320, borderRadius: '50%', background: 'rgba(61,139,130,0.10)', pointerEvents: 'none' }} />
      <div style={{ position: 'fixed', bottom: -100, left: -60, width: 280, height: 280, borderRadius: '50%', background: 'rgba(212,150,42,0.09)', pointerEvents: 'none' }} />

      <div style={{ width: '100%', maxWidth: 380, position: 'relative' }}>
        {/* Header card */}
        <div style={{ background: '#1e3340', borderRadius: '18px 18px 0 0', padding: '32px 32px 28px', textAlign: 'center' }}>
          {/* Icon circle */}
          <div style={{ width: 52, height: 52, borderRadius: '50%', background: 'rgba(61,139,130,0.25)', border: '2px solid rgba(61,139,130,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 16px', fontSize: 22 }}>
            🌐
          </div>
          <div style={{ fontSize: 10, fontWeight: 800, color: '#3d8b82', textTransform: 'uppercase', letterSpacing: '0.16em', marginBottom: 6 }}>
            Import PO System
          </div>
          <div style={{ fontSize: 26, fontWeight: 900, color: '#d4962a', lineHeight: 1.15 }}>
            Global Sourcing
          </div>
          <div style={{ fontSize: 12, color: '#7a9aaa', marginTop: 6 }}>
            เข้าสู่ระบบเพื่อใช้งาน
          </div>
        </div>

        {/* Form card */}
        <div style={{ background: '#faf5ee', border: '1.5px solid #e2d8c8', borderTop: 'none', borderRadius: '0 0 18px 18px', padding: '28px 32px 32px' }}>
          <form onSubmit={handleLogin} style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div>
              <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: '#8a7a6a', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>
                Email
              </label>
              <input
                type="email" value={email} onChange={e => setEmail(e.target.value)}
                placeholder="your@email.com" required autoComplete="email"
                style={inputStyle}
                onFocus={e => (e.target.style.borderColor = '#3d8b82')}
                onBlur={e => (e.target.style.borderColor = '#e2d8c8')}
              />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: '#8a7a6a', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>
                Password
              </label>
              <input
                type="password" value={password} onChange={e => setPassword(e.target.value)}
                placeholder="••••••••" required autoComplete="current-password"
                style={inputStyle}
                onFocus={e => (e.target.style.borderColor = '#3d8b82')}
                onBlur={e => (e.target.style.borderColor = '#e2d8c8')}
              />
            </div>

            {error && (
              <div style={{ background: 'rgba(200,90,58,0.08)', border: '1px solid rgba(200,90,58,0.2)', borderRadius: 8, padding: '8px 12px', fontSize: 12, color: '#c85a3a' }}>
                {error}
              </div>
            )}

            <button type="submit" disabled={loading} style={{
              width: '100%', background: loading ? '#a0c4c0' : '#3d8b82',
              color: '#fff', border: 'none', borderRadius: 10,
              padding: '13px 0', fontSize: 14, fontWeight: 700,
              cursor: loading ? 'not-allowed' : 'pointer', marginTop: 6,
              transition: 'background 0.15s', letterSpacing: '0.02em',
            }}>
              {loading ? 'กำลังเข้าสู่ระบบ...' : 'เข้าสู่ระบบ'}
            </button>
          </form>

          <div style={{ marginTop: 20, paddingTop: 16, borderTop: '1px solid #e8dcc8', textAlign: 'center', fontSize: 11, color: '#b0a090' }}>
            หากเข้าไม่ได้ หรือยังไม่ได้ลงทะเบียน · ติดต่อทีมงานเพื่อขอสิทธิ
          </div>
        </div>
      </div>
    </div>
  )
}

import { useState } from 'react'
import { supabase } from '../lib/supabase'

export default function Login() {
  const [email,    setEmail]    = useState('')
  const [password, setPassword] = useState('')
  const [loading,  setLoading]  = useState(false)
  const [error,    setError]    = useState('')

  async function handleLogin(e) {
    e.preventDefault()
    setLoading(true)
    setError('')
    const { error } = await supabase.auth.signInWithPassword({ email, password })
    if (error) setError(
      error.message === 'Invalid login credentials'
        ? 'Email o contraseña incorrectos'
        : error.message
    )
    setLoading(false)
  }

  return (
    <div style={{
      minHeight: '100vh', background: '#0D0D0D',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
    }}>
      <div style={{ width: 380 }}>

        {/* Logo */}
        <div style={{ textAlign: 'center', marginBottom: 36 }}>
          <img
            src="/logo.png"
            alt="Grupo Randazzo"
            style={{ height: 56, width: 'auto', display: 'inline-block' }}
            onError={e => { e.target.style.display = 'none'; e.target.nextSibling.style.display = 'block' }}
          />
          <div style={{ display: 'none', fontSize: '1.4rem', fontWeight: 900, color: '#B5CC2E', letterSpacing: '.05em' }}>
            GRUPO RANDAZZO
          </div>
          <div style={{ fontSize: '.75rem', color: '#555', marginTop: 8, letterSpacing: '.08em', textTransform: 'uppercase' }}>
            Gestoría · Área Finanzas
          </div>
        </div>

        {/* Card */}
        <div style={{
          background: '#1A1A1A', borderRadius: 12,
          border: '1px solid #252525', padding: '32px 28px',
        }}>
          <div style={{ fontSize: '1rem', fontWeight: 700, marginBottom: 4 }}>Iniciar sesión</div>
          <div style={{ fontSize: '.78rem', color: '#666', marginBottom: 24 }}>
            Acceso restringido al equipo de finanzas
          </div>

          <form onSubmit={handleLogin}>
            {error && (
              <div style={{
                background: 'rgba(229,57,53,0.1)', border: '1px solid rgba(229,57,53,0.3)',
                color: '#ff6b6b', borderRadius: 6, padding: '10px 12px',
                fontSize: '.78rem', marginBottom: 16,
              }}>
                {error}
              </div>
            )}

            <label style={{ fontSize: '.72rem', fontWeight: 600, color: '#888', display: 'block', marginBottom: 6, textTransform: 'uppercase', letterSpacing: '.05em' }}>
              Email
            </label>
            <input
              type="email"
              value={email}
              onChange={e => setEmail(e.target.value)}
              placeholder="nombre@gruporandazzo.com.ar"
              required
              style={{
                width: '100%', background: '#111', border: '1px solid #2a2a2a',
                color: '#fff', borderRadius: 7, padding: '10px 12px',
                fontSize: '.85rem', marginBottom: 16, outline: 'none',
              }}
              onFocus={e => e.target.style.borderColor = '#B5CC2E'}
              onBlur={e => e.target.style.borderColor = '#2a2a2a'}
            />

            <label style={{ fontSize: '.72rem', fontWeight: 600, color: '#888', display: 'block', marginBottom: 6, textTransform: 'uppercase', letterSpacing: '.05em' }}>
              Contraseña
            </label>
            <input
              type="password"
              value={password}
              onChange={e => setPassword(e.target.value)}
              placeholder="••••••••"
              required
              style={{
                width: '100%', background: '#111', border: '1px solid #2a2a2a',
                color: '#fff', borderRadius: 7, padding: '10px 12px',
                fontSize: '.85rem', marginBottom: 24, outline: 'none',
              }}
              onFocus={e => e.target.style.borderColor = '#B5CC2E'}
              onBlur={e => e.target.style.borderColor = '#2a2a2a'}
            />

            <button
              type="submit"
              disabled={loading}
              style={{
                width: '100%', background: '#B5CC2E', color: '#000',
                border: 'none', borderRadius: 7, padding: '11px 0',
                fontSize: '.875rem', fontWeight: 700, cursor: 'pointer',
                opacity: loading ? .6 : 1,
              }}
            >
              {loading ? 'Ingresando…' : 'Ingresar'}
            </button>
          </form>
        </div>
      </div>
    </div>
  )
}

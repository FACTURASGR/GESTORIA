import { useEffect, useState } from 'react'
import { Routes, Route, Navigate, NavLink, useNavigate } from 'react-router-dom'
import { supabase } from './lib/supabase'
import Login from './pages/Login'
import Dashboard from './pages/Dashboard'
import Historial from './pages/Historial'
import Conciliacion from './pages/Conciliacion'

export default function App() {
  const [session, setSession] = useState(undefined)

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session))
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => setSession(s))
    return () => subscription.unsubscribe()
  }, [])

  if (session === undefined) return <Loader />

  if (!session) return <Routes><Route path="*" element={<Login />} /></Routes>

  return (
    <div style={{ display: 'flex', minHeight: '100vh' }}>
      <Sidebar session={session} />
      <div style={{ flex: 1, marginLeft: 220, background: 'var(--bg)', minHeight: '100vh' }}>
        <Routes>
          <Route path="/"                 element={<Dashboard />} />
          <Route path="/historial"        element={<Historial />} />
          <Route path="/nueva"            element={<Conciliacion />} />
          <Route path="/conciliacion/:id" element={<Conciliacion />} />
          <Route path="*"                 element={<Navigate to="/" />} />
        </Routes>
      </div>
    </div>
  )
}

function Sidebar({ session }) {
  const navigate = useNavigate()

  async function logout() {
    await supabase.auth.signOut()
    navigate('/')
  }

  const navItems = [
    { to: '/',         label: 'Dashboard',          icon: '▦',  end: true },
    { to: '/historial',label: 'Historial',           icon: '📋', end: true },
    { to: '/nueva',    label: 'Nueva conciliación',  icon: '➕', end: true },
  ]

  return (
    <aside style={{
      width: 220, background: 'var(--sidebar)',
      borderRight: '1px solid var(--bdr)',
      position: 'fixed', top: 0, left: 0, bottom: 0,
      display: 'flex', flexDirection: 'column', zIndex: 100,
    }}>
      {/* LOGO */}
      <div style={{ padding: '22px 20px 18px', borderBottom: '1px solid var(--bdr)' }}>
        <img
          src="/logo.png"
          alt="Grupo Randazzo"
          style={{ width: '100%', maxWidth: 160, height: 'auto', display: 'block' }}
          onError={e => { e.target.style.display = 'none'; e.target.nextSibling.style.display = 'block' }}
        />
        <div style={{ display: 'none', fontSize: '.9rem', fontWeight: 900, color: 'var(--green)', letterSpacing: '.04em' }}>
          GRUPO RANDAZZO
        </div>
        <div style={{ fontSize: '.62rem', color: 'var(--txt3)', marginTop: 6, letterSpacing: '.08em', textTransform: 'uppercase' }}>
          Gestoría · Finanzas
        </div>
      </div>

      {/* NAV */}
      <nav style={{ flex: 1, padding: '14px 0' }}>
        <div style={{ fontSize: '.6rem', color: 'var(--txt3)', letterSpacing: '.1em', textTransform: 'uppercase', padding: '0 18px 8px' }}>
          Principal
        </div>
        {navItems.map(({ to, label, icon, end }) => (
          <NavLink
            key={to} to={to} end={end}
            style={({ isActive }) => ({
              display: 'flex', alignItems: 'center', gap: 10,
              padding: '9px 18px', fontSize: '.82rem',
              fontWeight: isActive ? 600 : 400,
              color: isActive ? 'var(--green)' : 'var(--txt2)',
              background: isActive ? 'var(--green-dim)' : 'transparent',
              borderLeft: `3px solid ${isActive ? 'var(--green)' : 'transparent'}`,
              textDecoration: 'none', transition: 'all .15s',
            })}
          >
            <span style={{ fontSize: '.95rem' }}>{icon}</span>
            {label}
          </NavLink>
        ))}


      </nav>

      {/* USER */}
      <div style={{ borderTop: '1px solid var(--bdr)', padding: '14px 18px' }}>
        <div style={{
          width: 30, height: 30, borderRadius: '50%',
          background: 'var(--green)', color: '#000',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontWeight: 700, fontSize: '.8rem', marginBottom: 8,
        }}>
          {session.user.email[0].toUpperCase()}
        </div>
        <div style={{ fontSize: '.7rem', color: 'var(--txt2)', marginBottom: 2, wordBreak: 'break-all' }}>
          {session.user.email}
        </div>
        <button
          onClick={logout}
          style={{
            marginTop: 10, width: '100%',
            background: 'transparent', border: '1px solid var(--bdr2)',
            color: 'var(--txt2)', borderRadius: 6, padding: '7px 0',
            fontSize: '.73rem', fontWeight: 500, cursor: 'pointer',
          }}
          onMouseEnter={e => { e.target.style.color = '#fff'; e.target.style.borderColor = '#444' }}
          onMouseLeave={e => { e.target.style.color = 'var(--txt2)'; e.target.style.borderColor = 'var(--bdr2)' }}
        >
          Cerrar sesión
        </button>
      </div>
    </aside>
  )
}

function Loader() {
  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100vh', background: '#111' }}>
      <div style={{ textAlign: 'center', color: '#444' }}>
        <div style={{ fontSize: '2rem', marginBottom: 10 }}>📊</div>
        <div style={{ fontSize: '.8rem' }}>Cargando…</div>
      </div>
    </div>
  )
}

import { useEffect, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { supabase } from '../lib/supabase'
import {
  BarChart, Bar, LineChart, Line,
  XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Legend, Cell,
} from 'recharts'

const EMPRESAS  = ['MOVILIS', 'KIARA', 'CIARA', 'PEARA', 'INV4', 'SALRA']
const GR_GREEN  = '#B5CC2E'
const MESES     = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

function fm(n) {
  if (!n && n !== 0) return '—'
  return parseFloat(n).toLocaleString('es-AR', { minimumFractionDigits: 0, maximumFractionDigits: 0 })
}
function fmPesos(n) {
  if (!n && n !== 0) return '—'
  const x = parseFloat(n)
  if (x >= 1_000_000) return `$${(x/1_000_000).toFixed(1)}M`
  if (x >= 1_000)     return `$${(x/1_000).toFixed(0)}K`
  return `$${x.toFixed(0)}`
}

// Tooltip personalizado con estilo oscuro
function DarkTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null
  return (
    <div style={{
      background: '#1A1A1A', border: '1px solid #2E2E2E',
      borderRadius: 8, padding: '10px 14px', fontSize: '.75rem',
    }}>
      <div style={{ color: '#888', marginBottom: 6, fontWeight: 600 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, marginBottom: 2 }}>
          {p.name}: <strong>{typeof p.value === 'number' && p.value > 1000 ? fmPesos(p.value) : p.value}</strong>
        </div>
      ))}
    </div>
  )
}

export default function Dashboard() {
  const navigate  = useNavigate()
  const [data,    setData]    = useState([])
  const [loading, setLoading] = useState(true)
  const [user,    setUser]    = useState(null)

  useEffect(() => {
    supabase.auth.getUser().then(({ data: { user } }) => setUser(user))
    fetchData()
  }, [])

  async function fetchData() {
    setLoading(true)
    const { data: rows, error } = await supabase
      .from('v_resumen')
      .select('*')
      .order('created_at', { ascending: false })
    if (!error) setData(rows || [])
    setLoading(false)
  }

  // ── CÁLCULOS DERIVADOS ──────────────────────────────────
  const totalConcil    = data.length
  const totalDepMonto  = data.reduce((s, r) => s + (parseFloat(r.dep_monto) || 0), 0)
  const totalPagMonto  = data.reduce((s, r) => s + (parseFloat(r.pag_monto) || 0), 0)
  const totalDiffs     = data.reduce((s, r) => s + (r.pag_con_diff || 0), 0)
  const totalPendientes= data.reduce((s, r) => s + (r.pendientes_total || 0), 0)
  const totalSinLote   = data.reduce((s, r) => s + (r.pag_sin_lote || 0), 0)

  // Gráfico por empresa
  const porEmpresa = EMPRESAS.map(emp => {
    const rows = data.filter(r => r.empresa === emp)
    return {
      empresa: emp,
      conciliaciones: rows.length,
      depositos: rows.reduce((s, r) => s + (parseFloat(r.dep_monto) || 0), 0),
      pagos:     rows.reduce((s, r) => s + (parseFloat(r.pag_monto) || 0), 0),
      diferencias: rows.reduce((s, r) => s + (r.pag_con_diff || 0), 0),
    }
  }).filter(r => r.conciliaciones > 0)

  // Gráfico por mes (últimos 6 meses)
  const porMes = (() => {
    const mapa = {}
    data.forEach(r => {
      const d = new Date(r.created_at)
      const key = `${MESES[d.getMonth()]} ${d.getFullYear()}`
      if (!mapa[key]) mapa[key] = { mes: key, conciliaciones: 0, diferencias: 0, pendientes: 0 }
      mapa[key].conciliaciones++
      mapa[key].diferencias  += r.pag_con_diff || 0
      mapa[key].pendientes   += r.pendientes_total || 0
    })
    return Object.values(mapa).slice(-6)
  })()

  // ── ÚLTIMO LOGIN FRIENDLY ────────────────────────────────
  const hora = new Date().getHours()
  const saludo = hora < 12 ? 'Buenos días' : hora < 19 ? 'Buenas tardes' : 'Buenas noches'
  const nombre = user?.email?.split('@')[0] || ''

  // ── RENDER ──────────────────────────────────────────────
  if (loading) return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '80vh' }}>
      <div style={{ color: '#555', textAlign: 'center' }}>
        <div style={{ fontSize: '2rem', marginBottom: 8 }}>📊</div>
        <div style={{ fontSize: '.82rem' }}>Cargando datos…</div>
      </div>
    </div>
  )

  return (
    <div style={{ padding: '28px 28px 56px' }}>

      {/* HEADER */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 28 }}>
        <div>
          <div style={{ fontSize: '1.15rem', fontWeight: 700, marginBottom: 4 }}>
            {saludo}{nombre ? `, ${nombre}` : ''} 👋
          </div>
          <div style={{ fontSize: '.78rem', color: '#888' }}>
            Resumen general del sistema de conciliación TH
          </div>
        </div>
        <button
          onClick={() => navigate('/nueva')}
          style={{ background: GR_GREEN, color: '#000', border: 'none', borderRadius: 8, padding: '10px 20px', fontSize: '.82rem', fontWeight: 700, cursor: 'pointer' }}
        >
          + Nueva conciliación
        </button>
      </div>

      {/* KPI ROW */}
      {totalConcil === 0 ? (
        <EmptyState onNueva={() => navigate('/nueva')} />
      ) : (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 12, marginBottom: 28 }}>
            {[
              { label: 'Conciliaciones',    value: fm(totalConcil),          sub: 'registradas',                color: GR_GREEN },
              { label: 'Total procesado',   value: fmPesos(totalDepMonto + totalPagMonto), sub: 'depósitos + pagos', color: GR_GREEN },
              { label: 'Sin lote asignado', value: fm(totalSinLote),          sub: 'completar con CELER',        color: totalSinLote  ? '#F59E0B' : GR_GREEN },
              { label: 'Diferencias > $50', value: fm(totalDiffs),            sub: 'requieren revisión',         color: totalDiffs    ? '#E53935' : GR_GREEN },
              { label: 'Pendientes activos',value: fm(totalPendientes),       sub: 'sin conciliar',              color: totalPendientes ? '#F59E0B' : GR_GREEN },
            ].map(k => (
              <div key={k.label} style={{ background: '#1A1A1A', borderRadius: 9, padding: '16px 18px', border: '1px solid #252525', borderLeft: `3px solid ${k.color}` }}>
                <div style={{ fontSize: '.65rem', color: '#666', textTransform: 'uppercase', letterSpacing: '.06em', marginBottom: 7 }}>{k.label}</div>
                <div style={{ fontSize: '1.5rem', fontWeight: 900, color: k.color, fontVariantNumeric: 'tabular-nums' }}>{k.value}</div>
                <div style={{ fontSize: '.7rem', color: '#555', marginTop: 4 }}>{k.sub}</div>
              </div>
            ))}
          </div>

          {/* CHARTS ROW */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 20 }}>

            {/* GRÁFICO: Por empresa */}
            <ChartCard title="Monto procesado por empresa" sub="Depósitos + pagos conciliados">
              <ResponsiveContainer width="100%" height={240}>
                <BarChart data={porEmpresa} margin={{ top: 4, right: 8, left: 0, bottom: 4 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#252525" vertical={false} />
                  <XAxis dataKey="empresa" tick={{ fill: '#666', fontSize: 11 }} axisLine={false} tickLine={false} />
                  <YAxis tickFormatter={v => fmPesos(v)} tick={{ fill: '#555', fontSize: 10 }} axisLine={false} tickLine={false} width={50} />
                  <Tooltip content={<DarkTooltip />} />
                  <Legend wrapperStyle={{ fontSize: '.72rem', color: '#666' }} />
                  <Bar dataKey="depositos" name="Depósitos" fill={GR_GREEN} radius={[4,4,0,0]} />
                  <Bar dataKey="pagos"     name="Pagos"     fill="#3B5E0A" radius={[4,4,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            </ChartCard>

            {/* GRÁFICO: Evolución mensual */}
            {porMes.length > 1 ? (
              <ChartCard title="Evolución mensual" sub="Conciliaciones registradas por mes">
                <ResponsiveContainer width="100%" height={240}>
                  <LineChart data={porMes} margin={{ top: 4, right: 8, left: 0, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#252525" vertical={false} />
                    <XAxis dataKey="mes" tick={{ fill: '#666', fontSize: 10 }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fill: '#555', fontSize: 10 }} axisLine={false} tickLine={false} width={28} />
                    <Tooltip content={<DarkTooltip />} />
                    <Legend wrapperStyle={{ fontSize: '.72rem', color: '#666' }} />
                    <Line type="monotone" dataKey="conciliaciones" name="Conciliaciones" stroke={GR_GREEN} strokeWidth={2} dot={{ fill: GR_GREEN, r: 4 }} activeDot={{ r: 6 }} />
                    <Line type="monotone" dataKey="diferencias"    name="Dif. > $50"    stroke="#E53935" strokeWidth={1.5} dot={{ fill: '#E53935', r: 3 }} strokeDasharray="4 2" />
                  </LineChart>
                </ResponsiveContainer>
              </ChartCard>
            ) : (
              <ChartCard title="Evolución mensual" sub="Aparecerá cuando haya más de un período registrado">
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: 240, color: '#444', fontSize: '.8rem' }}>
                  Registrá más conciliaciones para ver la evolución
                </div>
              </ChartCard>
            )}
          </div>

          {/* SEGUNDA FILA: Diferencias por empresa + tabla recientes */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.6fr', gap: 16 }}>

            {/* GRÁFICO: Diferencias por empresa */}
            <ChartCard title="Diferencias > $50 por empresa" sub="Items que requieren revisión">
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={porEmpresa} layout="vertical" margin={{ top: 4, right: 8, left: 0, bottom: 4 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#252525" horizontal={false} />
                  <XAxis type="number" tick={{ fill: '#555', fontSize: 10 }} axisLine={false} tickLine={false} />
                  <YAxis type="category" dataKey="empresa" tick={{ fill: '#666', fontSize: 11 }} axisLine={false} tickLine={false} width={52} />
                  <Tooltip content={<DarkTooltip />} />
                  <Bar dataKey="diferencias" name="Dif. > $50" radius={[0,4,4,0]}>
                    {porEmpresa.map((e, i) => (
                      <Cell key={i} fill={e.diferencias > 0 ? '#E53935' : '#2a2a2a'} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </ChartCard>

            {/* TABLA: Últimas conciliaciones */}
            <div style={{ background: '#1A1A1A', borderRadius: 10, border: '1px solid #252525', overflow: 'hidden' }}>
              <div style={{ padding: '16px 18px', borderBottom: '1px solid #252525' }}>
                <div style={{ fontSize: '.85rem', fontWeight: 700 }}>Últimas conciliaciones</div>
                <div style={{ fontSize: '.72rem', color: '#666', marginTop: 2 }}>Las más recientes del sistema</div>
              </div>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                  <tr>
                    {['Empresa', 'Período', 'Pagos', 'Sin lote', 'Estado'].map(h => (
                      <th key={h} style={{ padding: '9px 14px', textAlign: 'left', fontSize: '.65rem', fontWeight: 600, color: '#555', textTransform: 'uppercase', letterSpacing: '.06em', borderBottom: '1px solid #252525' }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.slice(0, 8).map(row => (
                    <tr
                      key={row.id}
                      onClick={() => navigate(`/conciliacion/${row.id}`)}
                      style={{ borderBottom: '1px solid #1f1f1f', cursor: 'pointer', transition: 'background .12s' }}
                      onMouseEnter={e => e.currentTarget.style.background = '#222'}
                      onMouseLeave={e => e.currentTarget.style.background = 'transparent'}
                    >
                      <td style={{ padding: '10px 14px', fontWeight: 700, fontSize: '.78rem' }}>{row.empresa}</td>
                      <td style={{ padding: '10px 14px', fontSize: '.72rem', color: '#666' }}>
                        {row.periodo_desde ? row.periodo_desde.slice(5) : '—'} → {row.periodo_hasta ? row.periodo_hasta.slice(5) : '—'}
                      </td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', fontSize: '.78rem', fontVariantNumeric: 'tabular-nums' }}>{row.pag_total || 0}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', fontSize: '.78rem', fontWeight: 600, color: (row.pag_sin_lote || 0) > 0 ? '#F59E0B' : '#555' }}>
                        {row.pag_sin_lote || 0}
                      </td>
                      <td style={{ padding: '10px 14px' }}>
                        <span style={{
                          display: 'inline-block', padding: '2px 8px', borderRadius: 20, fontSize: '.65rem', fontWeight: 600,
                          background: (row.pag_sin_lote || 0) === 0 && (row.pendientes_total || 0) === 0 ? 'rgba(181,204,46,.15)' : 'rgba(245,158,11,.15)',
                          color:      (row.pag_sin_lote || 0) === 0 && (row.pendientes_total || 0) === 0 ? GR_GREEN : '#F59E0B',
                        }}>
                          {(row.pag_sin_lote || 0) === 0 && (row.pendientes_total || 0) === 0 ? 'Completa' : 'En proceso'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  )
}

function ChartCard({ title, sub, children }) {
  return (
    <div style={{ background: '#1A1A1A', borderRadius: 10, border: '1px solid #252525', padding: '16px 18px' }}>
      <div style={{ fontSize: '.85rem', fontWeight: 700, marginBottom: 2 }}>{title}</div>
      <div style={{ fontSize: '.72rem', color: '#666', marginBottom: 14 }}>{sub}</div>
      {children}
    </div>
  )
}

function EmptyState({ onNueva }) {
  return (
    <div style={{
      background: '#1A1A1A', borderRadius: 12, border: '1px solid #252525',
      padding: '60px 40px', textAlign: 'center',
    }}>
      <div style={{ fontSize: '3rem', marginBottom: 16 }}>📊</div>
      <div style={{ fontSize: '1rem', fontWeight: 700, marginBottom: 8 }}>Sin datos todavía</div>
      <div style={{ fontSize: '.82rem', color: '#666', marginBottom: 24, maxWidth: 340, margin: '0 auto 24px' }}>
        El dashboard va a mostrar gráficos y métricas en tiempo real cuando empieces a registrar conciliaciones.
      </div>
      <button onClick={onNueva} style={{ background: '#B5CC2E', color: '#000', border: 'none', borderRadius: 8, padding: '11px 28px', fontSize: '.85rem', fontWeight: 700, cursor: 'pointer' }}>
        Crear primera conciliación
      </button>
    </div>
  )
}

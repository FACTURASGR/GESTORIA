import { useEffect, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { supabase } from '../lib/supabase'

const EMPRESAS = ['MOVILIS', 'KIARA', 'CIARA', 'PEARA', 'INV4', 'SALRA']

function fm(n) {
  if (!n && n !== 0) return '—'
  return parseFloat(n).toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

export default function Historial() {
  const [rows,    setRows]    = useState([])
  const [loading, setLoading] = useState(true)
  const [empresa, setEmpresa] = useState('')
  const navigate = useNavigate()

  useEffect(() => { fetchData() }, [empresa])

  async function fetchData() {
    setLoading(true)
    let q = supabase.from('v_resumen').select('*').order('created_at', { ascending: false }).limit(100)
    if (empresa) q = q.eq('empresa', empresa)
    const { data } = await q
    setRows(data || [])
    setLoading(false)
  }

  function estadoBadge(row) {
    const sinLote = row.pag_sin_lote || 0
    const pend    = row.pendientes_total || 0
    if (sinLote === 0 && pend === 0)
      return <span style={{ ...badge, background: 'rgba(181,204,46,.15)', color: '#B5CC2E' }}>Completa</span>
    return <span style={{ ...badge, background: 'rgba(245,158,11,.15)', color: '#F59E0B' }}>En proceso</span>
  }

  const badge = { display: 'inline-block', padding: '3px 10px', borderRadius: 20, fontSize: '.68rem', fontWeight: 600 }

  return (
    <div style={{ padding: '28px 28px 56px' }}>

      {/* HEADER */}
      <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 28 }}>
        <div>
          <div style={{ fontSize: '1.15rem', fontWeight: 700, marginBottom: 4 }}>Historial de conciliaciones</div>
          <div style={{ fontSize: '.78rem', color: 'var(--txt2)' }}>
            Todas las conciliaciones registradas en el sistema
          </div>
        </div>
        <button
          onClick={() => navigate('/nueva')}
          style={{
            background: 'var(--green)', color: '#000', border: 'none',
            borderRadius: 8, padding: '10px 20px', fontSize: '.82rem', fontWeight: 700,
          }}
        >
          + Nueva conciliación
        </button>
      </div>

      {/* FILTROS */}
      <div style={{ display: 'flex', gap: 8, marginBottom: 20 }}>
        {['', ...EMPRESAS].map(e => (
          <button
            key={e}
            onClick={() => setEmpresa(e)}
            style={{
              padding: '6px 14px', borderRadius: 20, fontSize: '.75rem', fontWeight: 600,
              border: '1px solid',
              borderColor: empresa === e ? 'var(--green)' : 'var(--bdr2)',
              background: empresa === e ? 'var(--green-dim)' : 'transparent',
              color: empresa === e ? 'var(--green)' : 'var(--txt2)',
            }}
          >
            {e || 'Todas'}
          </button>
        ))}
      </div>

      {/* TABLA */}
      <div style={{ background: 'var(--surf)', borderRadius: 10, border: '1px solid var(--bdr)', overflow: 'hidden' }}>
        {loading ? (
          <div style={{ textAlign: 'center', padding: '48px 20px', color: 'var(--txt2)', fontSize: '.82rem' }}>
            Cargando…
          </div>
        ) : rows.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '56px 20px' }}>
            <div style={{ fontSize: '2.5rem', marginBottom: 12 }}>📋</div>
            <div style={{ fontSize: '.9rem', fontWeight: 600, marginBottom: 8 }}>Sin conciliaciones todavía</div>
            <div style={{ fontSize: '.78rem', color: 'var(--txt2)', marginBottom: 20 }}>
              Creá la primera para empezar a registrar el historial
            </div>
            <button
              onClick={() => navigate('/nueva')}
              style={{ background: 'var(--green)', color: '#000', border: 'none', borderRadius: 7, padding: '10px 24px', fontSize: '.82rem', fontWeight: 700 }}
            >
              + Nueva conciliación
            </button>
          </div>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr style={{ borderBottom: '1px solid var(--bdr)' }}>
                {['Empresa', 'Período', 'Depósitos', 'Pagos', 'Sin lote', 'Dif. > $50', 'Pendientes', 'Estado', ''].map(h => (
                  <th key={h} style={{
                    padding: '12px 14px', textAlign: 'left',
                    fontSize: '.68rem', fontWeight: 600, color: 'var(--txt2)',
                    textTransform: 'uppercase', letterSpacing: '.06em', whiteSpace: 'nowrap',
                  }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map(row => (
                <tr
                  key={row.id}
                  onClick={() => navigate(`/conciliacion/${row.id}`)}
                  style={{ borderBottom: '1px solid var(--bdr)', cursor: 'pointer', transition: 'background .12s' }}
                  onMouseEnter={e => e.currentTarget.style.background = 'var(--surf2)'}
                  onMouseLeave={e => e.currentTarget.style.background = 'transparent'}
                >
                  <td style={{ padding: '13px 14px', fontWeight: 700, fontSize: '.82rem' }}>{row.empresa}</td>
                  <td style={{ padding: '13px 14px', fontSize: '.78rem', color: 'var(--txt2)', whiteSpace: 'nowrap' }}>
                    {row.periodo_desde} → {row.periodo_hasta}
                  </td>
                  <td style={{ padding: '13px 14px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', fontSize: '.78rem' }}>
                    {row.dep_total || 0}
                  </td>
                  <td style={{ padding: '13px 14px', textAlign: 'right', fontSize: '.78rem' }}>
                    {row.pag_total || 0}
                  </td>
                  <td style={{ padding: '13px 14px', textAlign: 'right', fontWeight: 600, fontSize: '.78rem',
                    color: (row.pag_sin_lote || 0) > 0 ? 'var(--amber)' : 'var(--green)' }}>
                    {row.pag_sin_lote || 0}
                  </td>
                  <td style={{ padding: '13px 14px', textAlign: 'right', fontWeight: 600, fontSize: '.78rem',
                    color: (row.pag_con_diff || 0) > 0 ? 'var(--red)' : 'var(--green)' }}>
                    {row.pag_con_diff || 0}
                  </td>
                  <td style={{ padding: '13px 14px', textAlign: 'right', fontSize: '.78rem',
                    color: (row.pendientes_total || 0) > 0 ? 'var(--amber)' : 'var(--txt2)' }}>
                    {row.pendientes_total || 0}
                  </td>
                  <td style={{ padding: '13px 14px' }}>{estadoBadge(row)}</td>
                  <td style={{ padding: '13px 14px', fontSize: '.72rem', color: 'var(--green)' }}>Ver →</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  )
}

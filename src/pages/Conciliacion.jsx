import { useState, useEffect, useMemo, useRef } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'
import {
  detectFileType, parseMayor, parseMov, parseOperaciones, parseRendiciones,
  conciliarTH, recalcularMov, armarLotes,
  formatDate, formatMonto, parseDate, isoDate,
} from '../lib/conciliacion'

const EMPRESAS = ['MOVILIS', 'KIARA', 'CIARA', 'PEARA', 'INV4', 'SALRA']

// ── ESTILOS BASE ────────────────────────────────────────
const th = {
  padding: '10px 12px', textAlign: 'left',
  fontSize: '.68rem', fontWeight: 600, color: 'var(--txt2)',
  textTransform: 'uppercase', letterSpacing: '.06em',
  whiteSpace: 'nowrap', borderBottom: '1px solid var(--bdr)',
  background: 'var(--surf)',
}
const td = {
  padding: '9px 12px', fontSize: '.76rem',
  borderBottom: '1px solid var(--bdr)', verticalAlign: 'middle',
}
const pill = (bg, txt) => ({
  display: 'inline-block', padding: '2px 8px', borderRadius: 4,
  fontSize: '.68rem', fontWeight: 700, background: bg, color: txt,
})
const btnGreen = {
  background: 'var(--green)', color: '#000', border: 'none',
  borderRadius: 7, padding: '9px 20px', fontSize: '.82rem', fontWeight: 700, cursor: 'pointer',
}
const btnOutline = {
  background: 'transparent', color: 'var(--txt2)',
  border: '1px solid var(--bdr2)', borderRadius: 7,
  padding: '9px 20px', fontSize: '.82rem', fontWeight: 500, cursor: 'pointer',
}

// ── DEFINICIÓN DE LOS 4 ARCHIVOS ────────────────────────
const SLOT_DEFS = [
  { slot: 0, tipo: 'mayor',       grupo: 'empresa', icon: '📊', label: 'Mayor Autodealer (K1)',    hint: 'K1 → Finanzas → Resumen de Cuenta → Export Excel 5.0' },
  { slot: 1, tipo: 'mov',         grupo: 'empresa', icon: '💳', label: 'Movimientos de Cuenta TH', hint: 'Web TH → Movimientos de Cuenta → Exportar' },
  { slot: 2, tipo: 'operaciones', grupo: 'empresa', icon: '📋', label: 'Operaciones de Pago TH',   hint: 'Web TH → Operaciones de Pago → Exportar' },
  { slot: 3, tipo: 'rendiciones', grupo: 'periodo', icon: '🗂️', label: 'Rendiciones CRM',          hint: 'CRM → Rendiciones → Reporte del período (todas las empresas)' },
]
const TIPO_A_SLOT = { mayor: 0, mov: 1, operaciones: 2, rendiciones: 3 }

// Reconoce qué reporte es por sus columnas (rendiciones primero, el resto lo resuelve la librería)
export function detectTipo(wb) {
  try {
    const ws = wb.Sheets[wb.SheetNames[0]]
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const s = (rows[i] || []).join('|').toLowerCase()
      if (s.includes('lote rendicion') || s.includes('num_minuta')) return 'rendiciones'
    }
  } catch (e) { /* si falla, deja que decida la librería */ }
  return detectFileType(wb)
}

// Lee TODAS las filas del reporte de Rendiciones (sin filtrar por empresa), solo para previsualizar
function leerRendiciones(wb) {
  const ws  = wb.Sheets[wb.SheetNames[0]]
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null })
  let h = 0
  for (let i = 0; i < Math.min(5, raw.length); i++) {
    const s = (raw[i] || []).join('|').toUpperCase()
    if (s.includes('NUM_MINUTA') || s.includes('LOTE')) { h = i; break }
  }
  const hdr = (raw[h] || []).map(x => String(x || '').trim().toUpperCase())
  const iS = hdr.findIndex(x => x.includes('SISTEMA'))
  const iM = hdr.findIndex(x => x.includes('NUM_MINUTA') || x.includes('MINUTA'))
  const iL = hdr.findIndex(x => x.includes('LOTE'))
  const iR = hdr.findIndex(x => x.includes('REAL') || x.includes('RENDIDO'))
  const out = []
  for (let i = h + 1; i < raw.length; i++) {
    const r = raw[i]
    if (!r || !r[iM]) continue
    out.push({
      sistema: String(r[iS] || '').toUpperCase().trim(),
      minuta:  String(r[iM]).trim(),
      lote:    String(r[iL] || '').trim(),
      monto:   parseFloat(r[iR]) || 0,
    })
  }
  return out
}

function rangoFechas(valores) {
  let min = null, max = null
  for (const v of valores) {
    const d = parseDate(v)
    if (!d) continue
    if (min === null || d < min) min = d
    if (max === null || d > max) max = d
  }
  return min ? `${formatDate(min)} → ${formatDate(max)}` : '—'
}
const sumar = (arr, f) => arr.reduce((s, x) => s + (f(x) || 0), 0)

// Arma lo que se muestra al previsualizar un archivo: tarjetas de resumen, avisos y tabla completa
export function buildPreview(slot, wb, ctx = {}) {
  const avisos = []
  const emp = (ctx.empresa || '').toUpperCase()

  if (slot === 0) {
    const items = parseMayor(wb)
    const op = items.filter(i => i.tipo === 'OP'), rm = items.filter(i => i.tipo === 'RM')
    if (!items.length) avisos.push({ tone: 'bad', text: 'No encontré movimientos OP ni RM. ¿Es el Mayor correcto?' })
    const sinRef = rm.filter(i => !i.referencia).length
    if (sinRef) avisos.push({ tone: 'warn', text: `${sinRef} línea(s) RM sin número después de "TARJETA HABITUALISTA N".` })
    avisos.push({ tone: 'info', text: 'Las líneas RC (recibos) se ignoran: no se concilian.' })
    return {
      resumen: `${items.length} movimientos · ${rangoFechas(items.map(i => i.fecha))}`,
      stats: [
        { label: 'OP (depósitos)', value: op.length, sub: '$ ' + formatMonto(sumar(op, i => i.monto)) },
        { label: 'RM (pagos)',     value: rm.length, sub: '$ ' + formatMonto(sumar(rm, i => i.monto)) },
        { label: 'Período',        value: rangoFechas(items.map(i => i.fecha)), small: true },
      ],
      columns: [
        { k: 'tipo', label: 'Tipo' }, { k: 'fecha', label: 'Fecha', fmt: 'fecha' },
        { k: 'descripcion', label: 'Comprobante' }, { k: 'referencia', label: 'Ref.' },
        { k: 'monto', label: 'Monto', fmt: 'monto', right: true },
      ],
      rows: items, avisos,
    }
  }

  if (slot === 1) {
    const items = parseMov(wb)
    const cr = items.filter(i => i.tipo === 'CREDIT'), db = items.filter(i => i.tipo === 'DEBIT')
    if (!items.length) avisos.push({ tone: 'bad', text: 'No encontré movimientos CREDIT ni DEBIT. ¿Es el reporte correcto?' })
    avisos.push({ tone: 'info', text: 'Se omiten los movimientos de monto 0 (por ejemplo comisiones bonificadas).' })
    return {
      resumen: `${items.length} movimientos · ${rangoFechas(items.map(i => i.fecha))}`,
      stats: [
        { label: 'CREDIT (ingresos)', value: cr.length, sub: '$ ' + formatMonto(sumar(cr, i => i.monto)) },
        { label: 'DEBIT (pagos)',     value: db.length, sub: '$ ' + formatMonto(sumar(db, i => i.monto)) },
        { label: 'Período',           value: rangoFechas(items.map(i => i.fecha)), small: true },
      ],
      columns: [
        { k: 'fecha', label: 'Fecha', fmt: 'fecha' }, { k: 'tipo', label: 'Tipo' },
        { k: 'descripcion', label: 'Descripción' }, { k: 'pagoNum', label: 'Pago N°' },
        { k: 'monto', label: 'Monto', fmt: 'monto', right: true },
      ],
      rows: items, avisos,
    }
  }

  if (slot === 2) {
    const items = parseOperaciones(wb)
    const con = items.filter(o => o.minuta), sin = items.filter(o => !o.minuta)
    const epagos   = sin.filter(o => /epagos/i.test(o.motivoPago)).length
    const ambiguos = sin.filter(o => o.minutaEstado === 'ambigua').length
    const otros    = sin.length - epagos - ambiguos
    if (!items.length) avisos.push({ tone: 'bad', text: 'No encontré operaciones de pago. ¿Es el reporte correcto?' })
    if (sin.length) avisos.push({ tone: 'warn', text: `${sin.length} pago(s) sin minuta identificable: ${epagos} EPAGOS, ${ambiguos} con varios números de 6 dígitos y ${otros} sin número. Esos hay que completarlos a mano después de conciliar.` })
    if (ctx.rendSet && ctx.opsMinutas) {
      const unicas = [...new Set(ctx.opsMinutas)]
      const hay = unicas.filter(m => ctx.rendSet.has(m)).length
      avisos.push({
        tone: unicas.length && hay / unicas.length < 0.5 ? 'warn' : 'ok',
        text: `Cruce con Rendiciones: ${hay} de ${unicas.length} minutas de estos pagos aparecen en el reporte de Rendiciones${emp ? ' (' + emp + ')' : ''}.` +
              (unicas.length && hay / unicas.length < 0.5 ? ' Si es poco, revisá que el reporte de Rendiciones cubra las fechas de rendición de estos pagos y que la empresa elegida sea la correcta.' : ''),
      })
    }
    return {
      resumen: `${items.length} pagos · ${con.length} con minuta`,
      stats: [
        { label: 'Pagos',        value: items.length, sub: '$ ' + formatMonto(sumar(items, o => o.total)) },
        { label: 'Con minuta',   value: con.length,   tone: 'ok' },
        { label: 'Sin minuta',   value: sin.length,   tone: sin.length ? 'warn' : 'ok' },
        { label: 'Período',      value: rangoFechas(items.map(o => o.fecha)), small: true },
      ],
      columns: [
        { k: 'nroPago', label: 'Nro pago' }, { k: 'fecha', label: 'Fecha', fmt: 'fecha' },
        { k: 'total', label: 'Total', fmt: 'monto', right: true }, { k: 'seccRPA', label: 'Seccional' },
        { k: 'tarjetahabiente', label: 'Tarjetahabiente' }, { k: 'minuta', label: 'Minuta' },
        { k: 'motivoPago', label: 'Motivo pago' },
      ],
      rows: items, avisos,
    }
  }

  // slot 3 — Rendiciones CRM
  const todas = leerRendiciones(wb)
  const filas = emp ? todas.filter(r => r.sistema === emp) : todas
  const empresas = [...new Set(todas.map(r => r.sistema))].sort()
  const lotes = new Set(filas.map(r => r.lote))
  const nums = filas.map(r => Number(r.minuta)).filter(n => !isNaN(n))
  if (!todas.length) avisos.push({ tone: 'bad', text: 'No encontré minutas. ¿Es el reporte de Rendiciones del CRM?' })
  if (!emp) avisos.push({ tone: 'info', text: 'Elegí la empresa arriba para ver y usar solo sus minutas.' })
  if (emp && todas.length && !filas.length) avisos.push({ tone: 'bad', text: `Este archivo no trae minutas de ${emp}. Trae: ${empresas.join(', ')}.` })
  if (ctx.opsMinutas) {
    const unicas = [...new Set(ctx.opsMinutas)]
    const set = new Set(filas.map(r => r.minuta))
    const hay = unicas.filter(m => set.has(m)).length
    avisos.push({
      tone: unicas.length && hay / unicas.length < 0.5 ? 'warn' : 'ok',
      text: `Cruce con Operaciones: ${hay} de ${unicas.length} minutas de los pagos cargados están en este reporte.`,
    })
  }
  return {
    resumen: `${filas.length} minutas · ${lotes.size} lotes` + (emp ? ` (${emp})` : ` · ${empresas.length} empresas`),
    stats: [
      { label: emp ? `Minutas ${emp}` : 'Minutas (todas)', value: filas.length, sub: '$ ' + formatMonto(sumar(filas, r => r.monto)) },
      { label: 'Lotes de rendición', value: lotes.size },
      { label: 'Empresas en el archivo', value: empresas.join(' · ') || '—', small: true },
      { label: 'Rango de minutas', value: nums.length ? `${Math.min(...nums)} → ${Math.max(...nums)}` : '—', small: true },
    ],
    columns: [
      { k: 'sistema', label: 'Empresa' }, { k: 'minuta', label: 'Minuta' },
      { k: 'lote', label: 'Lote' }, { k: 'monto', label: 'Rendido', fmt: 'monto', right: true },
    ],
    rows: filas, avisos,
  }
}

function safePreview(slot, wb, ctx) {
  try { return buildPreview(slot, wb, ctx) }
  catch (err) { return { error: err.message || 'No pude leer el archivo', resumen: '', stats: [], columns: [], rows: [], avisos: [] } }
}

const TONOS = {
  ok:   { c: 'var(--green)', bg: 'var(--green-dim)',  b: 'rgba(181,204,46,.35)' },
  warn: { c: 'var(--amber)', bg: 'var(--amber-dim)',  b: 'rgba(245,158,11,.35)' },
  bad:  { c: 'var(--red)',   bg: 'var(--red-dim)',    b: 'rgba(229,57,53,.35)' },
  info: { c: 'var(--txt2)',  bg: '#161616',           b: 'var(--bdr2)' },
}

function btnMini(activo, peligro) {
  return {
    background: activo ? 'var(--green)' : 'transparent',
    color: activo ? '#000' : (peligro ? 'var(--red)' : 'var(--txt)'),
    border: `1px solid ${activo ? 'var(--green)' : (peligro ? 'rgba(229,57,53,.5)' : 'var(--bdr2)')}`,
    borderRadius: 6, padding: '5px 10px', fontSize: '.7rem', fontWeight: 600, cursor: 'pointer',
  }
}

// Tarjeta de un archivo: zona para soltarlo si está vacío; si ya está cargado, resumen + ver/quitar
function SlotCard({ def, fileName, loaded, resumen, error, activo, onFiles, onRemove, onView }) {
  const [over, setOver] = useState(false)

  if (loaded) {
    return (
      <div style={{ border: `1.5px solid ${activo ? 'var(--green)' : 'var(--bdr2)'}`, background: 'var(--green-dim)', borderRadius: 10, padding: '14px 16px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
          <span style={{ fontSize: '1.3rem' }}>{def.icon}</span>
          <div style={{ fontSize: '.8rem', fontWeight: 700 }}>{def.label}</div>
        </div>
        <div title={fileName} style={{ fontSize: '.7rem', color: 'var(--green)', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>✓ {fileName}</div>
        <div style={{ fontSize: '.7rem', color: error ? 'var(--red)' : 'var(--txt2)', margin: '4px 0 10px' }}>{error || resumen}</div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={onView} style={btnMini(activo)}>{activo ? '▲ Ocultar' : '👁 Ver contenido'}</button>
          <button onClick={onRemove} style={btnMini(false, true)}>✕ Quitar</button>
        </div>
      </div>
    )
  }

  return (
    <div
      onDragOver={e => { e.preventDefault(); setOver(true) }}
      onDragLeave={() => setOver(false)}
      onDrop={e => { e.preventDefault(); setOver(false); onFiles([...e.dataTransfer.files]) }}
      style={{
        border: `1.5px dashed ${over ? 'var(--green)' : 'var(--bdr2)'}`,
        background: over ? 'var(--green-dim)' : '#111',
        borderRadius: 10, padding: '20px 16px', textAlign: 'center', cursor: 'pointer', position: 'relative',
      }}
    >
      <input type="file" accept=".xlsx,.xls,.csv" multiple
        style={{ position: 'absolute', inset: 0, opacity: 0, cursor: 'pointer', width: '100%', height: '100%' }}
        onChange={e => { onFiles([...e.target.files]); e.target.value = '' }} />
      <div style={{ fontSize: '1.5rem', marginBottom: 7 }}>{def.icon}</div>
      <div style={{ fontSize: '.8rem', fontWeight: 700, marginBottom: 4 }}>{def.label}</div>
      <div style={{ fontSize: '.67rem', color: 'var(--txt3)', marginBottom: 8 }}>{def.hint}</div>
      <div style={{ fontSize: '.7rem', color: 'var(--txt2)' }}>Arrastrá o hacé click</div>
    </div>
  )
}

// Panel con todo lo que contiene el archivo: resumen, avisos y tabla (con buscador)
function PreviewPanel({ def, fileName, preview, onClose, onRemove }) {
  const LIMITE = 200
  const [filtro, setFiltro] = useState('')
  const [verTodo, setVerTodo] = useState(false)

  const coincidentes = useMemo(() => {
    const q = filtro.trim().toLowerCase()
    if (!q) return preview.rows
    return preview.rows.filter(r => preview.columns.some(c => String(r[c.k] ?? '').toLowerCase().includes(q)))
  }, [preview, filtro])
  const visibles = verTodo ? coincidentes : coincidentes.slice(0, LIMITE)

  const celda = (c, v) => c.fmt === 'fecha' ? formatDate(v) : c.fmt === 'monto' ? formatMonto(v) : (v == null ? '' : String(v))

  return (
    <div style={{ border: '1px solid var(--green)', borderRadius: 10, background: '#0e0e0e', padding: '16px 18px', marginBottom: 20 }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, marginBottom: 14, flexWrap: 'wrap' }}>
        <div>
          <div style={{ fontSize: '.85rem', fontWeight: 700 }}>{def.icon} Previsualización · {def.label}</div>
          <div style={{ fontSize: '.7rem', color: 'var(--txt2)' }}>{fileName}</div>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={onRemove} style={btnMini(false, true)}>✕ Quitar archivo</button>
          <button onClick={onClose} style={btnMini(false)}>Cerrar</button>
        </div>
      </div>

      {preview.error ? (
        <div style={{ background: TONOS.bad.bg, border: `1px solid ${TONOS.bad.b}`, color: TONOS.bad.c, borderRadius: 8, padding: '12px 14px', fontSize: '.78rem' }}>
          No pude interpretar este archivo: {preview.error}
        </div>
      ) : (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 10, marginBottom: 12 }}>
            {preview.stats.map(s => (
              <div key={s.label} style={{ background: 'var(--surf)', border: '1px solid var(--bdr)', borderLeft: `3px solid ${(TONOS[s.tone] || TONOS.ok).c}`, borderRadius: 8, padding: '10px 12px' }}>
                <div style={{ fontSize: '.62rem', color: 'var(--txt2)', textTransform: 'uppercase', letterSpacing: '.05em', marginBottom: 4 }}>{s.label}</div>
                <div style={{ fontSize: s.small ? '.78rem' : '1.2rem', fontWeight: 800, lineHeight: 1.3 }}>{s.value}</div>
                {s.sub && <div style={{ fontSize: '.68rem', color: 'var(--txt2)', marginTop: 2 }}>{s.sub}</div>}
              </div>
            ))}
          </div>

          {preview.avisos.map((a, i) => (
            <div key={i} style={{ background: TONOS[a.tone].bg, border: `1px solid ${TONOS[a.tone].b}`, color: TONOS[a.tone].c, borderRadius: 8, padding: '8px 12px', fontSize: '.74rem', marginBottom: 8 }}>
              {a.text}
            </div>
          ))}

          <div style={{ display: 'flex', alignItems: 'center', gap: 12, margin: '12px 0 8px', flexWrap: 'wrap' }}>
            <input
              value={filtro} onChange={e => { setFiltro(e.target.value); setVerTodo(false) }}
              placeholder="Buscar en el archivo (minuta, monto, texto…)"
              style={{ flex: '1 1 260px', background: '#111', border: '1px solid var(--bdr2)', color: '#fff', borderRadius: 7, padding: '8px 10px', fontSize: '.78rem', outline: 'none' }}
            />
            <span style={{ fontSize: '.72rem', color: 'var(--txt2)' }}>
              Mostrando {visibles.length} de {coincidentes.length}{filtro ? ` (${preview.rows.length} en total)` : ''}
            </span>
            {!verTodo && coincidentes.length > LIMITE && (
              <button onClick={() => setVerTodo(true)} style={btnMini(false)}>Mostrar las {coincidentes.length}</button>
            )}
          </div>

          <div style={{ maxHeight: 360, overflow: 'auto', border: '1px solid var(--bdr)', borderRadius: 8, background: 'var(--surf)' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>{preview.columns.map(c => <th key={c.k} style={{ ...th, position: 'sticky', top: 0, textAlign: c.right ? 'right' : 'left' }}>{c.label}</th>)}</tr>
              </thead>
              <tbody>
                {visibles.map((r, i) => (
                  <tr key={i}>
                    {preview.columns.map(c => (
                      <td key={c.k} style={{ ...td, textAlign: c.right ? 'right' : 'left', fontVariantNumeric: c.right ? 'tabular-nums' : 'normal', maxWidth: 380, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={celda(c, r[c.k])}>
                        {celda(c, r[c.k])}
                      </td>
                    ))}
                  </tr>
                ))}
                {!visibles.length && (
                  <tr><td colSpan={preview.columns.length} style={{ ...td, textAlign: 'center', color: 'var(--txt2)', padding: 24 }}>Sin filas para mostrar</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  )
}

export default function Conciliacion() {
  const { id } = useParams()
  const navigate = useNavigate()

  const [workbooks,    setWorkbooks]    = useState([null, null, null, null])
  const [fileNames,    setFileNames]    = useState(['', '', '', ''])
  const [empresa,      setEmpresa]      = useState('')
  const [periodoDesde, setPeriodoDesde] = useState('')
  const [periodoHasta, setPeriodoHasta] = useState('')
  const [results,      setResults]      = useState(null)
  const [concilId,     setConcilId]     = useState(id || null)
  const [processing,   setProcessing]   = useState(false)
  const [saving,       setSaving]       = useState(false)
  const [activeTab,    setActiveTab]    = useState(0)
  const [previewSlot,  setPreviewSlot]  = useState(null)
  const [notice,       setNotice]       = useState(null)
  const cargadoRef = useRef(null)   // evita recargar desde Supabase la conciliación que acabamos de guardar

  useEffect(() => {
    if (id && id !== cargadoRef.current) { cargadoRef.current = id; loadFromSupabase(id) }
  }, [id])

  // Reabre una conciliación guardada. Los totales de Mayor y Rendiciones por lote no se guardan:
  // para verlos hay que volver a conciliar con los archivos.
  async function loadFromSupabase(cId) {
    setProcessing(true)
    const [{ data: concil }, { data: deps }, { data: pags }, { data: pends }] = await Promise.all([
      supabase.from('conciliaciones').select('*').eq('id', cId).single(),
      supabase.from('depositos').select('*').eq('conciliacion_id', cId),
      supabase.from('pagos').select('*').eq('conciliacion_id', cId),
      supabase.from('pendientes').select('*').eq('conciliacion_id', cId),
    ])
    if (concil) {
      setEmpresa(concil.empresa)
      setPeriodoDesde(concil.periodo_desde || '')
      setPeriodoHasta(concil.periodo_hasta || '')
      const depositos = (deps || []).map(d => ({
        mayor: { referencia: d.op_ref, fecha: d.fecha_mayor, monto: d.monto_mayor, descripcion: d.desc_mayor },
        th: { codRef: d.cod_ref_th, fecha: d.fecha_th, monto: d.monto_th, descripcion: d.desc_th },
        dias: d.dias_diferencia, estado: d.estado,
      }))
      const pendCredits = (pends || []).filter(p => p.origen === 'th' && p.tipo_movimiento === 'CREDIT')
        .map(p => ({ codRef: p.referencia, fecha: p.fecha, monto: p.monto, descripcion: p.descripcion, posibleDevolucion: !/dep[oó]sito/i.test(p.descripcion || '') }))
      const pendOP = (pends || []).filter(p => p.origen === 'mayor')
        .map(p => ({ item: { tipo: p.tipo_movimiento, fecha: p.fecha, monto: p.monto, descripcion: p.descripcion, referencia: p.referencia }, razon: p.motivo }))
      const filas = (pags || []).map(p => ({
        _dbId: p.id, nroOp: p.nro_operacion, nroPago: p.nro_pago_th, fecha: p.fecha_th, total: parseFloat(p.monto_th) || 0,
        seccRPA: p.secc_rpa || '', tarjetahabiente: p.tarjetahabiente || '', tarjeta: p.tarjeta || '', motivoPago: p.motivo_pago || '',
        minuta: p.minuta || '', minutaEstado: p.minuta ? 'ok' : 'manual', minutaCandidatos: [], loteManual: p.lote || '',
      }))
      const dbDif = (pags || []).map(p => (p.diferencia === null || p.diferencia === undefined) ? null : parseFloat(p.diferencia))
      const mov = recalcularMov(filas, null).map((r, i) => ({
        ...r, diferencia: dbDif[i],
        estadoFila: (pags[i].estado === 'revisar') ? 'revisar' : r.estadoFila,
        alerta: (pags[i].estado === 'revisar') ? 'Marcado para revisar al guardar' : r.alerta,
      }))
      setResults({
        depositos, pendCredits, pendOP,
        creditos: [
          ...depositos.map(d => ({ codRef: d.th.codRef, fecha: d.th.fecha, descripcion: d.th.descripcion, monto: d.th.monto, movimiento: d.mayor.referencia })),
          ...pendCredits.map(c => ({ codRef: c.codRef, fecha: c.fecha, descripcion: c.descripcion, monto: c.monto, movimiento: 'OP A CONCILIAR' })),
        ],
        mov, lotes: armarLotes(mov, [], {}), mayorRM: [], rendMap: null, rendLotes: {}, mayorLineas: [], avisos: [], deBase: true,
      })
      setConcilId(cId)
    }
    setProcessing(false)
  }

  async function readFile(file, hint) {
    try {
      const buf  = await file.arrayBuffer()
      const wb   = XLSX.read(buf, { type: 'array', cellDates: true })
      const tipo = detectTipo(wb)
      let slot = TIPO_A_SLOT[tipo]
      let aviso = null
      if (slot === undefined) {
        slot = hint
        aviso = { kind: 'warn', text: `No pude identificar el formato de «${file.name}». Lo cargué como ${SLOT_DEFS[slot].label}: revisá la previsualización.` }
      } else if (slot !== hint) {
        aviso = { kind: 'info', text: `«${file.name}» es ${SLOT_DEFS[slot].label}: lo cargué en su lugar.` }
      }
      setWorkbooks(prev => { const n = [...prev]; n[slot] = wb; return n })
      setFileNames(prev => { const n = [...prev]; n[slot] = file.name; return n })
      setPreviewSlot(slot)
      setNotice(aviso)
    } catch (err) {
      setNotice({ kind: 'error', text: `No pude leer «${file.name}»: ${err.message}` })
    }
  }

  function readFiles(files, hint) { files.forEach(f => readFile(f, hint)) }

  function removeFile(slot) {
    setWorkbooks(prev => prev.map((w, i) => i === slot ? null : w))
    setFileNames(prev => prev.map((n, i) => i === slot ? '' : n))
    setPreviewSlot(p => p === slot ? null : p)
    setNotice(null)
  }

  // Previsualización de cada archivo cargado (se recalcula si cambia un archivo o la empresa)
  const previews = useMemo(() => {
    const emp = (empresa || '').toUpperCase()
    let rendSet = null, opsMinutas = null
    if (workbooks[3]) {
      try { rendSet = new Set(leerRendiciones(workbooks[3]).filter(r => !emp || r.sistema === emp).map(r => r.minuta)) } catch (e) { /* se informa en su propia previsualización */ }
    }
    if (workbooks[2]) {
      try { opsMinutas = parseOperaciones(workbooks[2]).map(o => o.minuta).filter(Boolean) } catch (e) { /* idem */ }
    }
    return [0, 1, 2, 3].map(s => workbooks[s] ? safePreview(s, workbooks[s], { empresa, rendSet, opsMinutas }) : null)
  }, [workbooks, empresa])

  async function procesar() {
    setProcessing(true)
    try {
      const rendMap = workbooks[3] ? parseRendiciones(workbooks[3], empresa) : null
      const res = conciliarTH({
        mayorItems: parseMayor(workbooks[0]),
        movItems: parseMov(workbooks[1]),
        operaciones: parseOperaciones(workbooks[2]),
        rendMap, empresa,
      })
      setResults(res)
      setActiveTab(0)
    } catch (err) { alert('Error: ' + err.message) }
    setProcessing(false)
  }

  // Edición manual de un pago (minuta o lote): recalcula lote, diferencia, estado y el control por lote
  function editarMov(idx, cambios) {
    setResults(prev => {
      const base = prev.mov.map((r, i) => i === idx ? { ...r, ...cambios } : r)
      const mov = recalcularMov(base, prev.rendMap)
      return { ...prev, mov, lotes: armarLotes(mov, prev.mayorRM, prev.rendLotes) }
    })
  }

  // Si la conciliación ya está guardada, el cambio manual se guarda al salir del campo
  async function persistirMov(r) {
    if (!r || !r._dbId) return
    const estado = r.estadoFila === 'ok' ? 'ok' : r.estadoFila === 'revisar' ? 'revisar' : 'sin_lote'
    await supabase.from('pagos').update({
      minuta: r.minuta || null, lote: r.lote || null, diferencia: r.diferencia ?? 0, estado,
      updated_at: new Date().toISOString(),
    }).eq('id', r._dbId)
  }

  async function guardar() {
    if (!results || !empresa) return
    setSaving(true)
    const ok = async q => { const { error } = await q; if (error) throw error }
    try {
      let cId = concilId
      if (!cId) {
        const { data, error } = await supabase.from('conciliaciones').insert({
          empresa, periodo_desde: periodoDesde || null, periodo_hasta: periodoHasta || null, estado: 'borrador',
        }).select('id').single()
        if (error) throw error
        cId = data.id
        setConcilId(cId)
        cargadoRef.current = cId
        navigate(`/conciliacion/${cId}`, { replace: true })
      }
      await ok(supabase.from('depositos').delete().eq('conciliacion_id', cId))
      await ok(supabase.from('pagos').delete().eq('conciliacion_id', cId))
      await ok(supabase.from('pendientes').delete().eq('conciliacion_id', cId))

      if (results.depositos.length) {
        await ok(supabase.from('depositos').insert(results.depositos.map(d => ({
          conciliacion_id: cId, op_ref: d.mayor.referencia,
          fecha_mayor: isoDate(d.mayor.fecha), monto_mayor: d.mayor.monto, desc_mayor: d.mayor.descripcion,
          cod_ref_th: d.th.codRef || null, fecha_th: isoDate(d.th.fecha), monto_th: d.th.monto, desc_th: d.th.descripcion,
          estado: d.estado || 'conciliado', dias_diferencia: d.dias || 0,
        }))))
      }

      if (results.mov.length) {
        const { data: ins, error } = await supabase.from('pagos').insert(results.mov.map(r => ({
          conciliacion_id: cId, fecha_th: isoDate(r.fecha), monto_th: r.total,
          nro_pago_th: r.nroPago || null, nro_operacion: r.nroOp != null ? String(r.nroOp) : null,
          minuta: r.minuta || null, secc_rpa: r.seccRPA || null, tarjetahabiente: r.tarjetahabiente || null,
          tarjeta: r.tarjeta || null, motivo_pago: r.motivoPago || null,
          lote: r.lote || null, diferencia: r.diferencia ?? 0,
          estado: r.estadoFila === 'ok' ? 'ok' : r.estadoFila === 'revisar' ? 'revisar' : 'sin_lote',
        }))).select('id')
        if (error) throw error
        if (ins) setResults(prev => ({ ...prev, mov: prev.mov.map((r, i) => ({ ...r, _dbId: ins[i]?.id })) }))
      }

      const pend = [
        ...results.pendOP.map(p => ({ conciliacion_id: cId, origen: 'mayor', tipo_movimiento: p.item.tipo, fecha: isoDate(p.item.fecha), monto: p.item.monto, descripcion: p.item.descripcion, referencia: p.item.referencia, motivo: p.razon })),
        ...results.pendCredits.map(p => ({ conciliacion_id: cId, origen: 'th', tipo_movimiento: 'CREDIT', fecha: isoDate(p.fecha), monto: p.monto, descripcion: p.descripcion, referencia: p.codRef, motivo: p.posibleDevolucion ? 'Posible devolución' : 'OP a conciliar' })),
      ]
      if (pend.length) await ok(supabase.from('pendientes').insert(pend))
      alert('✓ Guardado correctamente')
    } catch (err) { alert('Error al guardar: ' + err.message); console.error(err) }
    setSaving(false)
  }

  function exportExcel() {
    if (!results) return
    const { creditos, mov, lotes, pendCredits, pendOP } = results
    const wb = XLSX.utils.book_new()
    // Las fechas del reporte vienen como texto (31/05/26): se pasan a fecha real, como el paso «Convertir XX en 20XX» del instructivo
    const f = v => parseDate(v) || v
    const hoja = aoa => {
      const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true })
      Object.keys(ws).forEach(k => { if (k[0] !== '!' && ws[k].t === 'd') ws[k].z = 'dd/mm/yyyy' })
      return ws
    }

    // DEB-CRED: solo CREDIT, con la columna MOVIMIENTO (G) completa
    const deb = hoja([
      ['Cod referencia', 'Fecha acreditacion', 'Descripcion', 'Tipo', 'Monto', 'Disponible despues de movimiento', 'MOVIMIENTO'],
      ...creditos.map(c => [c.codRef || '', f(c.fecha), c.descripcion, 'CREDIT', c.monto, c.disponible ?? '', c.movimiento]),
    ])
    deb['!cols'] = [{ wch: 16 }, { wch: 14 }, { wch: 70 }, { wch: 8 }, { wch: 16 }, { wch: 18 }, { wch: 16 }]
    XLSX.utils.book_append_sheet(wb, deb, 'DEB-CRED')

    // Mov: mismo orden de columnas que la hoja Mov de la conciliación (N Minuta · O Diferencias · P Lote)
    const movHoja = hoja([
      ['Nro. operacion', 'Nro de pago', 'Canal', 'Acreditacion', 'Nro cupon', 'Fecha', 'Total', 'Secc. RPA', 'Tarjetahabiente', 'Tarjeta', 'Estado', 'Motivo pago', 'Observaciones', 'Minuta', 'Diferencias', 'Lote', 'OBSERVACIONES'],
      ...mov.map(r => [r.nroOp, r.nroPago, r.canal || '', r.acreditacion || '', r.nroCupon ?? '', f(r.fecha), r.total, r.seccRPA, r.tarjetahabiente, r.tarjeta, r.estado || '', r.motivoPago, r.observaciones || '',
        r.minuta ? Number(r.minuta) : '', r.diferencia ?? '', r.lote ? (/^\d+$/.test(r.lote) ? Number(r.lote) : r.lote) : '', r.alerta || '']),
    ])
    movHoja['!cols'] = [{ wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 11 }, { wch: 14 }, { wch: 26 }, { wch: 26 }, { wch: 20 }, { wch: 11 }, { wch: 46 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 9 }, { wch: 60 }]
    XLSX.utils.book_append_sheet(wb, movHoja, 'Mov')

    // Mayor: la hoja descargada de Autodealer + columna L (OP-xxxx para depósitos, N° de lote para RM). Las RC se ignoran.
    if (workbooks[0]) {
      const raw = XLSX.utils.sheet_to_json(workbooks[0].Sheets[workbooks[0].SheetNames[0]], { header: 1, raw: true, defval: null })
      const filas = raw.map(r => [...r])
      for (const it of results.mayorLineas || []) {
        if (it.fila == null || !filas[it.fila] || !it.referencia) continue
        while (filas[it.fila].length < 12) filas[it.fila].push(null)
        filas[it.fila][11] = it.tipo === 'RM' && /^\d+$/.test(it.referencia) ? Number(it.referencia) : it.referencia
      }
      XLSX.utils.book_append_sheet(wb, hoja(filas), 'Mayor')
    }

    // Control por lote
    const ctl = hoja([
      ['Lote', 'Mayor (RM)', 'Rendiciones CRM', 'Pagos Mov', 'Cant. pagos', 'Dif. Mov - CRM', 'Dif. Mov - Mayor', 'Dif. CRM - Mayor', 'Estado'],
      ...lotes.map(l => [Number(l.lote), l.mayor ?? '', l.crm ?? '', l.mov, l.movN, l.difMovCrm ?? '', l.difMovMayor ?? '', l.difCrmMayor ?? '',
        { ok: 'Cierra', revisar: 'Revisar', sin_pagos: 'Sin pagos en este Mov', sin_mayor: 'Sin RM en el Mayor' }[l.estado]]),
    ])
    ctl['!cols'] = [{ wch: 9 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 11 }, { wch: 15 }, { wch: 16 }, { wch: 16 }, { wch: 24 }]
    XLSX.utils.book_append_sheet(wb, ctl, 'Control por lote')

    // Pendientes / alertas
    const pend = hoja([
      ['Tipo', 'Fecha', 'Monto', 'Detalle', 'Observación'],
      ...pendCredits.map(c => ['OP a conciliar (CREDIT sin OP)', f(c.fecha), c.monto, c.descripcion, c.posibleDevolucion ? 'Posible devolución' : '']),
      ...pendOP.map(p => ['OP del Mayor sin CREDIT', f(p.item.fecha), p.item.monto, p.item.descripcion, p.razon]),
      ...mov.filter(r => r.estadoFila !== 'ok').map(r => ['Pago a completar o revisar', f(r.fecha), r.total, r.motivoPago, r.alerta]),
    ])
    pend['!cols'] = [{ wch: 32 }, { wch: 12 }, { wch: 16 }, { wch: 70 }, { wch: 70 }]
    XLSX.utils.book_append_sheet(wb, pend, 'Pendientes')

    XLSX.writeFile(wb, `Conciliacion_${empresa}_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  if (processing && id) return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '80vh' }}>
      <div style={{ textAlign: 'center', color: 'var(--txt2)' }}>
        <div style={{ fontSize: '2rem', marginBottom: 10 }}>⏳</div>
        <div>Cargando conciliación…</div>
      </div>
    </div>
  )

  return (
    <div style={{ padding: '28px 28px 56px' }}>

      {/* HEADER */}
      <div style={{ marginBottom: 24 }}>
        <div style={{ fontSize: '1.1rem', fontWeight: 700, marginBottom: 4 }}>
          {concilId ? `Conciliación · ${empresa}` : 'Nueva conciliación'}
        </div>
        <div style={{ fontSize: '.78rem', color: 'var(--txt2)' }}>
          {concilId ? `Período: ${periodoDesde} → ${periodoHasta}` : 'Completá los datos y cargá los tres archivos'}
        </div>
      </div>

      {/* UPLOAD */}
      {!results && (
        <div style={{ background: 'var(--surf)', borderRadius: 10, border: '1px solid var(--bdr)', padding: 24, marginBottom: 20 }}>
          {/* INSTRUCCIONES DE ORIGEN */}
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: '.72rem', fontWeight: 700, color: 'var(--txt2)', textTransform: 'uppercase', letterSpacing: '.07em', marginBottom: 12 }}>
              ¿De dónde viene cada archivo?
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 10 }}>
              {[
                {
                  num: '1', icon: '📊', titulo: 'Mayor (K1)',
                  origen: 'Autodealer',
                  pasos: ['K1 → Sistema de Finanzas', 'Resumen de Cuenta', 'Filtrar: Tarjeta Habitualista', 'Exportar Excel 5.0'],
                },
                {
                  num: '2', icon: '💳', titulo: 'Movimientos de Cuenta',
                  origen: 'Web Tarjeta Habitualista',
                  pasos: ['Ingresar a la web de TH', 'Movimientos de cuenta', 'Seleccionar período', 'Exportar / Descargar'],
                },
                {
                  num: '3', icon: '📋', titulo: 'Operaciones de Pago',
                  origen: 'Web Tarjeta Habitualista',
                  pasos: ['Ingresar a la web de TH', 'Operaciones de Pago', 'Seleccionar período', 'Exportar / Descargar'],
                },
              ].map(item => (
                <div key={item.num} style={{ background: '#111', border: '1px solid var(--bdr)', borderRadius: 8, padding: '14px 16px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
                    <div style={{ width: 22, height: 22, borderRadius: '50%', background: 'var(--green)', color: '#000', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '.72rem', fontWeight: 900, flexShrink: 0 }}>
                      {item.num}
                    </div>
                    <div>
                      <div style={{ fontSize: '.8rem', fontWeight: 700 }}>{item.icon} {item.titulo}</div>
                      <div style={{ fontSize: '.68rem', color: 'var(--txt3)' }}>Origen: {item.origen}</div>
                    </div>
                  </div>
                  {item.pasos.map((p, i) => (
                    <div key={i} style={{ fontSize: '.7rem', color: 'var(--txt2)', padding: '2px 0 2px 8px', borderLeft: '2px solid var(--bdr2)' }}>
                      {p}
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>

          {/* SEPARADOR */}
          <div style={{ borderTop: '1px solid var(--bdr)', margin: '0 0 20px', position: 'relative' }}>
            <span style={{ position: 'absolute', top: -9, left: '50%', transform: 'translateX(-50%)', background: 'var(--surf)', padding: '0 12px', fontSize: '.68rem', color: 'var(--txt3)', textTransform: 'uppercase', letterSpacing: '.07em' }}>
              Subí los archivos acá abajo
            </span>
          </div>

          {/* Config empresa + período */}
          <div style={{ display: 'grid', gridTemplateColumns: '180px 1fr 1fr', gap: 16, marginBottom: 20 }}>
            <div>
              <div style={{ fontSize: '.7rem', fontWeight: 600, color: 'var(--txt2)', marginBottom: 6, textTransform: 'uppercase', letterSpacing: '.05em' }}>Empresa *</div>
              <select
                value={empresa} onChange={e => setEmpresa(e.target.value)}
                style={{ width: '100%', background: '#111', border: '1px solid var(--bdr2)', color: '#fff', borderRadius: 7, padding: '9px 10px', fontSize: '.82rem', outline: 'none' }}
              >
                <option value="">Seleccioná</option>
                {EMPRESAS.map(e => <option key={e}>{e}</option>)}
              </select>
            </div>
            {[['periodoDesde','Desde', periodoDesde, setPeriodoDesde], ['periodoHasta','Hasta', periodoHasta, setPeriodoHasta]].map(([k, label, val, set]) => (
              <div key={k}>
                <div style={{ fontSize: '.7rem', fontWeight: 600, color: 'var(--txt2)', marginBottom: 6, textTransform: 'uppercase', letterSpacing: '.05em' }}>{label}</div>
                <input type="date" value={val} onChange={e => set(e.target.value)}
                  style={{ width: '100%', background: '#111', border: '1px solid var(--bdr2)', color: '#fff', borderRadius: 7, padding: '9px 10px', fontSize: '.82rem', outline: 'none' }} />
              </div>
            ))}
          </div>

          {/* AVISO (archivo reubicado / no reconocido / error) */}
          {notice && (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, background: TONOS[notice.kind === 'error' ? 'bad' : notice.kind === 'warn' ? 'warn' : 'info'].bg, border: `1px solid ${TONOS[notice.kind === 'error' ? 'bad' : notice.kind === 'warn' ? 'warn' : 'info'].b}`, color: TONOS[notice.kind === 'error' ? 'bad' : notice.kind === 'warn' ? 'warn' : 'info'].c, borderRadius: 8, padding: '9px 12px', fontSize: '.76rem', marginBottom: 14 }}>
              <span>{notice.text}</span>
              <button onClick={() => setNotice(null)} style={{ background: 'none', border: 'none', color: 'inherit', cursor: 'pointer', fontSize: '.9rem' }}>✕</button>
            </div>
          )}

          {/* ARCHIVOS POR EMPRESA */}
          <div style={{ fontSize: '.7rem', fontWeight: 700, color: 'var(--txt2)', textTransform: 'uppercase', letterSpacing: '.07em', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ background: 'var(--green)', color: '#000', borderRadius: 4, padding: '1px 7px', fontSize: '.65rem' }}>1 por empresa</span>
            Se descargan individualmente para la empresa seleccionada
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 12, marginBottom: 20 }}>
            {SLOT_DEFS.filter(d => d.grupo === 'empresa').map(def => (
              <SlotCard key={def.slot} def={def}
                fileName={fileNames[def.slot]} loaded={!!workbooks[def.slot]}
                resumen={previews[def.slot]?.resumen} error={previews[def.slot]?.error}
                activo={previewSlot === def.slot}
                onFiles={files => readFiles(files, def.slot)}
                onRemove={() => removeFile(def.slot)}
                onView={() => setPreviewSlot(previewSlot === def.slot ? null : def.slot)} />
            ))}
          </div>

          {/* RENDICIONES CRM — un solo archivo para todo el período */}
          <div style={{ border: '1px solid var(--bdr2)', borderRadius: 10, padding: '16px 18px', marginBottom: 20, background: '#0e0e0e' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
              <span style={{ background: '#1A73E8', color: '#fff', borderRadius: 4, padding: '1px 7px', fontSize: '.65rem', fontWeight: 700 }}>Todo el período</span>
              <div>
                <div style={{ fontSize: '.78rem', fontWeight: 700 }}>Rendiciones CRM <span style={{ color: 'var(--txt3)', fontWeight: 500 }}>· opcional</span></div>
                <div style={{ fontSize: '.68rem', color: 'var(--txt2)' }}>Un solo archivo con todas las empresas. Si no lo subís, los lotes se cargan a mano.</div>
              </div>
            </div>
            <SlotCard def={SLOT_DEFS[3]}
              fileName={fileNames[3]} loaded={!!workbooks[3]}
              resumen={previews[3]?.resumen} error={previews[3]?.error}
              activo={previewSlot === 3}
              onFiles={files => readFiles(files, 3)}
              onRemove={() => removeFile(3)}
              onView={() => setPreviewSlot(previewSlot === 3 ? null : 3)} />
          </div>

          {/* PREVISUALIZACIÓN DEL ARCHIVO SELECCIONADO */}
          {previewSlot !== null && workbooks[previewSlot] && previews[previewSlot] && (
            <PreviewPanel key={previewSlot + '|' + fileNames[previewSlot]}
              def={SLOT_DEFS[previewSlot]} fileName={fileNames[previewSlot]} preview={previews[previewSlot]}
              onClose={() => setPreviewSlot(null)} onRemove={() => removeFile(previewSlot)} />
          )}

          {/* CHECKLIST */}
          <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', fontSize: '.72rem', margin: '0 0 14px' }}>
            {SLOT_DEFS.map(d => (
              <span key={d.slot} style={{ color: workbooks[d.slot] ? 'var(--green)' : 'var(--txt3)', fontWeight: workbooks[d.slot] ? 600 : 400 }}>
                {workbooks[d.slot] ? '✓' : '○'} {d.label}{d.slot === 3 ? ' (opcional)' : ''}
              </span>
            ))}
          </div>

          <button
            onClick={procesar}
            disabled={!workbooks[0] || !workbooks[1] || !workbooks[2] || !empresa || processing}
            style={{ ...btnGreen, opacity: (!workbooks[0] || !workbooks[1] || !workbooks[2] || !empresa || processing) ? .4 : 1 }}
          >
            {processing ? '⏳ Procesando…' : '⚡ Conciliar'}
          </button>
          {(!empresa || !workbooks[0] || !workbooks[1] || !workbooks[2]) && (
            <span style={{ marginLeft: 12, fontSize: '.75rem', color: 'var(--txt2)' }}>
              {!empresa ? 'Seleccioná la empresa primero' : 'Cargá los 3 archivos obligatorios'}
            </span>
          )}
        </div>
      )}

      {/* RESULTADOS */}
      {results && (
        <ResultsSection
          results={results}
          empresa={empresa}
          concilId={concilId}
          saving={saving}
          activeTab={activeTab}
          setActiveTab={setActiveTab}
          onGuardar={guardar}
          onExport={exportExcel}
          onNueva={() => { setResults(null); setWorkbooks([null, null, null, null]); setFileNames(['', '', '', '']); setConcilId(null); cargadoRef.current = null; navigate('/nueva') }}
          onEditar={() => setResults(null)}
          puedeEditar={!!(workbooks[0] && workbooks[1] && workbooks[2])}
          onEditMov={editarMov}
          onPersistMov={persistirMov}
        />
      )}
    </div>
  )
}

const ESTADO_FILA = {
  ok:         { txt: '✓ Listo',    bg: 'var(--green-dim)', c: 'var(--green)' },
  revisar:    { txt: 'Revisar',    bg: 'var(--red-dim)',   c: 'var(--red)' },
  sin_lote:   { txt: 'Sin lote',   bg: 'var(--amber-dim)', c: 'var(--amber)' },
  sin_minuta: { txt: 'Sin minuta', bg: 'var(--amber-dim)', c: 'var(--amber)' },
}
const ESTADO_LOTE = {
  ok:        { txt: 'Cierra',                 bg: 'var(--green-dim)', c: 'var(--green)', tip: 'Mayor, Rendiciones y pagos de Mov coinciden (tolerancia $50)' },
  revisar:   { txt: 'Revisar',                bg: 'var(--red-dim)',   c: 'var(--red)',   tip: 'Hay una diferencia mayor a $50 entre Mayor, Rendiciones y Mov' },
  sin_pagos: { txt: 'Sin pagos en este Mov',  bg: 'var(--amber-dim)', c: 'var(--amber)', tip: 'Ningún pago de este Mov tiene este lote: pueden ser pagos de un período anterior' },
  sin_mayor: { txt: 'Sin RM en el Mayor',     bg: 'var(--amber-dim)', c: 'var(--amber)', tip: 'Hay pagos con este lote pero no hay una línea RM en el Mayor del período' },
}

function ResultsSection({ results, empresa, concilId, saving, activeTab, setActiveTab, onGuardar, onExport, onNueva, onEditar, puedeEditar, onEditMov, onPersistMov }) {
  const { depositos, pendCredits, pendOP, mov, lotes, mayorRM } = results
  const totDep    = depositos.reduce((s, r) => s + r.mayor.monto, 0)
  const totMov    = mov.reduce((s, r) => s + r.total, 0)
  const conLote   = mov.filter(r => r.lote).length
  const sinMinuta = mov.filter(r => r.estadoFila === 'sin_minuta').length
  const sinLote   = mov.filter(r => r.estadoFila === 'sin_lote').length
  const revisar   = mov.filter(r => r.estadoFila === 'revisar').length
  const lotesRev  = lotes.filter(l => l.estado === 'revisar').length
  const rmSinNum  = (mayorRM || []).filter(m => !m.referencia)
  const pendTotal = pendCredits.length + pendOP.length + rmSinNum.length

  const kpis = [
    { label: 'Depósitos conciliados', value: depositos.length, sub: pendCredits.length ? `${pendCredits.length} OP a conciliar` : `$ ${formatMonto(totDep)}`, color: pendCredits.length ? 'var(--amber)' : 'var(--green)' },
    { label: 'Pagos (Mov)',           value: mov.length,       sub: `$ ${formatMonto(totMov)}`, color: 'var(--green)' },
    { label: 'Con lote asignado',     value: `${conLote}/${mov.length}`, sub: conLote === mov.length ? 'Todos asignados' : `${mov.length - conLote} sin lote`, color: conLote === mov.length ? 'var(--green)' : 'var(--amber)' },
    { label: 'A completar a mano',    value: sinMinuta + sinLote, sub: `${sinMinuta} sin minuta · ${sinLote} sin lote`, color: (sinMinuta + sinLote) ? 'var(--amber)' : 'var(--green)' },
    { label: 'A revisar',             value: revisar + lotesRev, sub: `${revisar} pagos · ${lotesRev} lotes con diferencia`, color: (revisar + lotesRev) ? 'var(--red)' : 'var(--green)' },
  ]

  const tabs = [
    `💰 Depósitos (${depositos.length + pendCredits.length})`,
    `📋 Pagos · Mov (${mov.length})`,
    `🧮 Lotes (${lotes.length})`,
    `⚠ Pendientes (${pendTotal})`,
  ]

  return (
    <>
      <div style={{ display: 'flex', gap: 10, marginBottom: 20, flexWrap: 'wrap', alignItems: 'center' }}>
        <button onClick={onGuardar} disabled={saving} style={btnGreen}>
          {saving ? '⏳ Guardando…' : concilId ? '💾 Actualizar' : '💾 Guardar'}
        </button>
        <button onClick={onExport} style={btnOutline}>⬇ Descargar Excel</button>
        {puedeEditar && <button onClick={onEditar} style={btnOutline}>✎ Editar archivos</button>}
        <button onClick={onNueva} style={btnOutline}>↺ Nueva</button>
        {concilId && <span style={{ fontSize: '.72rem', color: 'var(--txt2)' }}>Los cambios de minuta y lote se guardan al salir del campo</span>}
      </div>

      {(results.avisos || []).map((a, i) => (
        <div key={i} style={{ background: TONOS.warn.bg, border: `1px solid ${TONOS.warn.b}`, color: TONOS.warn.c, borderRadius: 8, padding: '9px 12px', fontSize: '.76rem', marginBottom: 12 }}>⚠ {a}</div>
      ))}
      {results.deBase && (
        <div style={{ background: TONOS.info.bg, border: `1px solid ${TONOS.info.b}`, color: TONOS.info.c, borderRadius: 8, padding: '9px 12px', fontSize: '.76rem', marginBottom: 12 }}>
          Conciliación guardada. Los totales de Mayor y Rendiciones por lote no se guardan: para verlos, volvé a conciliar con los archivos.
        </div>
      )}

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5,1fr)', gap: 12, marginBottom: 22 }}>
        {kpis.map(k => (
          <div key={k.label} style={{ background: 'var(--surf)', borderRadius: 9, padding: '16px', border: '1px solid var(--bdr)', borderLeft: `3px solid ${k.color}` }}>
            <div style={{ fontSize: '.65rem', color: 'var(--txt2)', textTransform: 'uppercase', letterSpacing: '.05em', marginBottom: 6 }}>{k.label}</div>
            <div style={{ fontSize: '1.4rem', fontWeight: 900, color: k.color, fontVariantNumeric: 'tabular-nums' }}>{k.value}</div>
            <div style={{ fontSize: '.7rem', color: 'var(--txt2)', marginTop: 3 }}>{k.sub}</div>
          </div>
        ))}
      </div>

      <div style={{ display: 'flex', gap: 2, borderBottom: '1px solid var(--bdr)', marginBottom: 18, flexWrap: 'wrap' }}>
        {tabs.map((label, i) => (
          <button key={i} onClick={() => setActiveTab(i)} style={{
            padding: '9px 18px', background: 'none', border: 'none', cursor: 'pointer', fontSize: '.8rem', fontWeight: 600,
            color: activeTab === i ? 'var(--green)' : 'var(--txt2)',
            borderBottom: `2px solid ${activeTab === i ? 'var(--green)' : 'transparent'}`, marginBottom: -1,
          }}>{label}</button>
        ))}
      </div>

      {/* DEPÓSITOS: OP del Mayor ↔ CREDIT de Movimientos → columna MOVIMIENTO de DEB-CRED */}
      {activeTab === 0 && (
        <>
          {pendCredits.length > 0 && (
            <div style={{ background: 'var(--amber-dim)', border: '1px solid rgba(245,158,11,.3)', borderRadius: 8, padding: '12px 16px', marginBottom: 14, fontSize: '.78rem', color: 'var(--amber)' }}>
              <strong>⚠ {pendCredits.length} CREDIT sin OP en el Mayor</strong> — quedan como «OP A CONCILIAR» en DEB-CRED (anotarlos en B7 de la pestaña Conciliación). Si no es un depósito, puede ser una devolución.
            </div>
          )}
          <TableBox>
            <thead><tr>
              {['OP (Mayor)', 'Fecha Mayor', 'Monto Mayor', 'Fecha TH', 'Descripción TH (CREDIT)', 'Monto TH', 'Días', 'MOVIMIENTO'].map(h => <th key={h} style={th}>{h}</th>)}
            </tr></thead>
            <tbody>
              {depositos.map((r, i) => (
                <tr key={i}>
                  <td style={td}><strong style={{ color: 'var(--green)' }}>{r.mayor.referencia}</strong></td>
                  <td style={td}>{formatDate(r.mayor.fecha)}</td>
                  <td style={{ ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{formatMonto(r.mayor.monto)}</td>
                  <td style={td}>{formatDate(r.th.fecha)}</td>
                  <td style={{ ...td, maxWidth: 300, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={r.th.descripcion}>{r.th.descripcion}</td>
                  <td style={{ ...td, textAlign: 'right' }}>{formatMonto(r.th.monto)}</td>
                  <td style={{ ...td, textAlign: 'right', color: r.dias > 3 ? 'var(--amber)' : 'var(--green)', fontWeight: 600 }}>{r.dias}</td>
                  <td style={td}><span style={pill('var(--green-dim)', 'var(--green)')}>{r.mayor.referencia}</span></td>
                </tr>
              ))}
              {pendCredits.map((r, i) => (
                <tr key={'c' + i}>
                  <td style={{ ...td, color: 'var(--txt2)' }}>—</td>
                  <td style={{ ...td, color: 'var(--txt2)' }}>—</td>
                  <td style={{ ...td, color: 'var(--txt2)', textAlign: 'right' }}>—</td>
                  <td style={td}>{formatDate(r.fecha)}</td>
                  <td style={{ ...td, maxWidth: 300, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={r.descripcion}>{r.descripcion}</td>
                  <td style={{ ...td, textAlign: 'right', color: 'var(--red)', fontWeight: 600 }}>{formatMonto(r.monto)}</td>
                  <td style={{ ...td, color: 'var(--txt2)' }}>—</td>
                  <td style={td}><span style={pill('var(--amber-dim)', 'var(--amber)')}>{r.posibleDevolucion ? '¿Devolución?' : 'OP A CONCILIAR'}</span></td>
                </tr>
              ))}
              {!depositos.length && !pendCredits.length && (
                <tr><td colSpan={8} style={{ ...td, textAlign: 'center', color: 'var(--txt2)', padding: 24 }}>No hay depósitos en este período</td></tr>
              )}
            </tbody>
          </TableBox>
        </>
      )}

      {activeTab === 1 && <PagosTab mov={mov} hayCRM={!!results.rendMap} onEdit={onEditMov} onPersist={onPersistMov} />}

      {/* LOTES: cada RM del Mayor es un lote */}
      {activeTab === 2 && (
        <>
          <div style={{ background: TONOS.info.bg, border: `1px solid ${TONOS.info.b}`, color: 'var(--txt2)', borderRadius: 8, padding: '10px 14px', fontSize: '.74rem', marginBottom: 14 }}>
            Cada línea <strong>RM</strong> del Mayor es un lote de rendición. La suma de los pagos de Mov con ese lote tiene que coincidir con el Mayor y con el total del lote en Rendiciones (se revisa toda diferencia mayor a $50).
          </div>
          <TableBox>
            <thead><tr>
              {['Lote', 'Mayor (RM)', 'Rendiciones CRM', 'Pagos Mov', 'Dif. Mov − CRM', 'Dif. Mov − Mayor', 'Estado'].map((h, i) => <th key={h} style={{ ...th, textAlign: i >= 1 && i <= 5 ? 'right' : 'left' }}>{h}</th>)}
            </tr></thead>
            <tbody>
              {lotes.map(l => {
                const e = ESTADO_LOTE[l.estado]
                const dif = d => d === null ? '—' : (Math.abs(d) < 0.01 ? '0,00' : formatMonto(d))
                const colorDif = d => d === null ? 'var(--txt3)' : Math.abs(d) > 50 ? 'var(--red)' : Math.abs(d) > 0.01 ? 'var(--amber)' : 'var(--green)'
                return (
                  <tr key={l.lote}>
                    <td style={td}><strong>{l.lote}</strong></td>
                    <td style={{ ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{l.mayor === null ? '—' : formatMonto(l.mayor)}</td>
                    <td style={{ ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{l.crm === null ? '—' : formatMonto(l.crm)}</td>
                    <td style={{ ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{l.movN ? `${formatMonto(l.mov)} (${l.movN})` : '—'}</td>
                    <td style={{ ...td, textAlign: 'right', fontWeight: 600, color: colorDif(l.difMovCrm) }}>{dif(l.difMovCrm)}</td>
                    <td style={{ ...td, textAlign: 'right', fontWeight: 600, color: colorDif(l.difMovMayor) }}>{dif(l.difMovMayor)}</td>
                    <td style={td}><span title={e.tip} style={{ ...pill(e.bg, e.c), cursor: 'help' }}>{e.txt}</span></td>
                  </tr>
                )
              })}
              {!lotes.length && <tr><td colSpan={7} style={{ ...td, textAlign: 'center', color: 'var(--txt2)', padding: 24 }}>Todavía no hay lotes para controlar: asigná lotes a los pagos o cargá el Mayor del período</td></tr>}
            </tbody>
          </TableBox>
        </>
      )}

      {/* PENDIENTES */}
      {activeTab === 3 && (
        pendTotal === 0
          ? <div style={{ textAlign: 'center', padding: '48px 20px', color: 'var(--txt2)' }}><div style={{ fontSize: '2rem', marginBottom: 10 }}>✅</div>Sin pendientes de depósitos ni de Mayor.</div>
          : <>
            {pendOP.length > 0 && (
              <Section label={`📌 OP del Mayor sin depósito en TH (${pendOP.length})`}>
                <TableBox>
                  <thead><tr>{['OP', 'Fecha', 'Descripción', 'Monto', 'Motivo'].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
                  <tbody>{pendOP.map((r, i) => (
                    <tr key={i}>
                      <td style={td}><strong>{r.item.referencia}</strong></td>
                      <td style={td}>{formatDate(r.item.fecha)}</td>
                      <td style={{ ...td, maxWidth: 320 }}>{r.item.descripcion}</td>
                      <td style={{ ...td, textAlign: 'right', color: 'var(--amber)', fontWeight: 600 }}>{formatMonto(r.item.monto)}</td>
                      <td style={{ ...td, fontSize: '.72rem', color: 'var(--txt2)' }}>{r.razon}</td>
                    </tr>
                  ))}</tbody>
                </TableBox>
              </Section>
            )}
            {rmSinNum.length > 0 && (
              <Section label={`📌 Líneas RM del Mayor sin número de lote (${rmSinNum.length})`}>
                <TableBox>
                  <thead><tr>{['Fecha', 'Descripción', 'Monto'].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
                  <tbody>{rmSinNum.map((r, i) => (
                    <tr key={i}>
                      <td style={td}>{formatDate(r.fecha)}</td>
                      <td style={{ ...td, maxWidth: 420 }}>{r.descripcion}</td>
                      <td style={{ ...td, textAlign: 'right', color: 'var(--amber)', fontWeight: 600 }}>{formatMonto(r.monto)}</td>
                    </tr>
                  ))}</tbody>
                </TableBox>
              </Section>
            )}
            {pendCredits.length > 0 && (
              <Section label={`💳 CREDIT de TH sin OP en el Mayor (${pendCredits.length})`}>
                <TableBox>
                  <thead><tr>{['Fecha', 'Descripción', 'Monto', 'Observación'].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
                  <tbody>{pendCredits.map((r, i) => (
                    <tr key={i}>
                      <td style={td}>{formatDate(r.fecha)}</td>
                      <td style={{ ...td, maxWidth: 420 }}>{r.descripcion}</td>
                      <td style={{ ...td, textAlign: 'right', color: 'var(--red)', fontWeight: 600 }}>{formatMonto(r.monto)}</td>
                      <td style={{ ...td, fontSize: '.72rem', color: 'var(--txt2)' }}>{r.posibleDevolucion ? 'No parece un depósito: ¿devolución?' : 'OP a conciliar'}</td>
                    </tr>
                  ))}</tbody>
                </TableBox>
              </Section>
            )}
          </>
      )}
    </>
  )
}

// Pagos de "Operaciones de pago" (hoja Mov): minuta, lote, total y diferencia contra Rendiciones.
// Los que no se pueden resolver solos se completan acá a mano.
function PagosTab({ mov, hayCRM, onEdit, onPersist }) {
  const [filtro, setFiltro] = useState('pendientes')
  const sinMinuta = mov.filter(r => r.estadoFila === 'sin_minuta').length
  const sinLote   = mov.filter(r => r.estadoFila === 'sin_lote').length
  const revisar   = mov.filter(r => r.estadoFila === 'revisar').length
  const necesitan = mov.map((r, i) => ({ r, i })).filter(({ r }) => r.estadoFila !== 'ok')
  const filas = (filtro === 'pendientes' ? necesitan : mov.map((r, i) => ({ r, i })))

  const inputBase = { borderRadius: 5, padding: '4px 7px', fontSize: '.75rem', fontWeight: 600, color: '#fff', outline: 'none', width: 84 }

  return (
    <>
      {sinMinuta > 0 && (
        <div style={{ background: 'var(--amber-dim)', border: '1px solid rgba(245,158,11,.35)', borderRadius: 8, padding: '12px 16px', marginBottom: 12, fontSize: '.78rem', color: 'var(--amber)' }}>
          <strong>⚠ {sinMinuta} pago(s) sin minuta identificable.</strong> No tienen un número de 6 dígitos en «Motivo pago» (por ejemplo AUTORIZACION EPAGOS) o tienen varios. Cargá la minuta a mano en la columna Minuta, o escribí directamente el lote.
        </div>
      )}
      {!hayCRM && (
        <div style={{ background: TONOS.info.bg, border: `1px solid ${TONOS.info.b}`, borderRadius: 8, padding: '10px 14px', marginBottom: 12, fontSize: '.76rem', color: 'var(--txt2)' }}>
          No cargaste el reporte de Rendiciones CRM: los lotes hay que escribirlos a mano. (Para asignarlos solos, usá «Editar archivos» y subilo.)
        </div>
      )}
      {sinLote > 0 && hayCRM && (
        <div style={{ background: 'var(--amber-dim)', border: '1px solid rgba(245,158,11,.35)', borderRadius: 8, padding: '10px 14px', marginBottom: 12, fontSize: '.76rem', color: 'var(--amber)' }}>
          {sinLote} pago(s) tienen minuta pero esa minuta no figura en el reporte de Rendiciones. Revisá que el reporte cubra la fecha en que se rindieron, o asigná el lote a mano.
        </div>
      )}
      {revisar > 0 && (
        <div style={{ background: 'var(--red-dim)', border: '1px solid rgba(229,57,53,.35)', borderRadius: 8, padding: '10px 14px', marginBottom: 12, fontSize: '.76rem', color: 'var(--red)' }}>
          {revisar} pago(s) para revisar: diferencia mayor a $50 contra Rendiciones o posible pago duplicado. Buscá el lote físico y estudiá la minuta.
        </div>
      )}

      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10, flexWrap: 'wrap' }}>
        {[['pendientes', `Necesitan acción (${necesitan.length})`], ['todos', `Todos (${mov.length})`]].map(([k, label]) => (
          <button key={k} onClick={() => setFiltro(k)} style={btnMini(filtro === k)}>{label}</button>
        ))}
        <span style={{ fontSize: '.7rem', color: 'var(--txt3)' }}>Lote en verde = asignado desde Rendiciones · azul = cargado a mano</span>
      </div>

      <TableBox>
        <thead><tr>
          {['Estado', 'Fecha', 'Nro pago', 'Total', 'Motivo pago', 'Minuta', 'Lote', 'Rendido CRM', 'Diferencia'].map((h, i) => <th key={h} style={{ ...th, textAlign: i === 3 || i >= 7 ? 'right' : 'left' }}>{h}</th>)}
        </tr></thead>
        <tbody>
          {filas.map(({ r, i }) => {
            const e = ESTADO_FILA[r.estadoFila]
            const manualLote = r.loteManual && String(r.loteManual).trim() !== ''
            return (
              <tr key={i}>
                <td style={td}><span title={r.alerta} style={{ ...pill(e.bg, e.c), cursor: r.alerta ? 'help' : 'default', whiteSpace: 'nowrap' }}>{e.txt}</span></td>
                <td style={{ ...td, whiteSpace: 'nowrap' }}>{formatDate(r.fecha)}</td>
                <td style={td}>{r.nroPago}</td>
                <td style={{ ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{formatMonto(r.total)}</td>
                <td style={{ ...td, maxWidth: 260, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={r.motivoPago}>{r.motivoPago}</td>
                <td style={td}>
                  <input value={r.minuta} maxLength={6} inputMode="numeric" placeholder="6 dígitos"
                    onChange={e => onEdit(i, { minuta: e.target.value.replace(/\D/g, ''), minutaEstado: 'editada' })}
                    onBlur={() => onPersist(r)}
                    style={{ ...inputBase, background: r.minuta ? '#111' : 'var(--amber-dim)', border: `1px solid ${r.minuta ? 'var(--bdr2)' : 'var(--amber)'}` }} />
                  {!r.minuta && r.minutaCandidatos && r.minutaCandidatos.length > 1 && (
                    <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap', marginTop: 4, maxWidth: 160 }}>
                      {r.minutaCandidatos.slice(0, 6).map(c => (
                        <button key={c} onClick={() => { onEdit(i, { minuta: c, minutaEstado: 'editada' }) }} style={{ ...btnMini(false), padding: '1px 6px' }}>{c}</button>
                      ))}
                    </div>
                  )}
                </td>
                <td style={td}>
                  <input value={r.lote} placeholder="Lote"
                    onChange={e => onEdit(i, { loteManual: e.target.value })}
                    onBlur={() => onPersist(r)}
                    style={{ ...inputBase, background: r.lote ? (manualLote ? 'var(--blue-dim)' : 'var(--green-dim)') : '#111', border: `1px solid ${r.lote ? (manualLote ? 'var(--blue)' : 'var(--green)') : 'var(--amber)'}` }} />
                </td>
                <td style={{ ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: r.rendido === null ? 'var(--txt3)' : 'inherit' }}>
                  {r.rendido === null ? '—' : formatMonto(r.rendido)}
                  {r.pagoMultiple && <span title="Esta minuta se pagó en más de una vez: la diferencia está calculada sobre la suma de sus pagos" style={{ color: 'var(--txt2)', marginLeft: 4, cursor: 'help' }}>Σ</span>}
                </td>
                <td style={{ ...td, textAlign: 'right', fontWeight: 600, color: r.diferencia === null ? 'var(--txt3)' : Math.abs(r.diferencia) > 50 ? 'var(--red)' : Math.abs(r.diferencia) > 0.01 ? 'var(--amber)' : 'var(--green)' }}>
                  {r.diferencia === null ? '—' : (Math.abs(r.diferencia) < 0.01 ? '0,00' : formatMonto(r.diferencia))}
                </td>
              </tr>
            )
          })}
          {!filas.length && (
            <tr><td colSpan={9} style={{ ...td, textAlign: 'center', color: 'var(--txt2)', padding: 28 }}>✅ No hay pagos que necesiten acción</td></tr>
          )}
        </tbody>
      </TableBox>
    </>
  )
}

function TableBox({ children }) {
  return (
    <div style={{ overflow: 'auto', borderRadius: 8, border: '1px solid var(--bdr)', background: 'var(--surf)' }}>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>{children}</table>
    </div>
  )
}

function Section({ label, children }) {
  return (
    <div style={{ marginBottom: 24 }}>
      <div style={{ fontSize: '.8rem', fontWeight: 700, marginBottom: 10, color: 'var(--txt)' }}>{label}</div>
      {children}
    </div>
  )
}

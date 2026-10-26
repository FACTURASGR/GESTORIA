import { useState, useEffect } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'
import {
  detectFileType, parseMayor, parseMov, parseOperaciones, parseRendiciones, buildMatchesConLotes,
  formatDate, formatMonto,
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

  useEffect(() => { if (id) loadFromSupabase(id) }, [id])

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
      setResults({
        depositos:   (deps  || []).map(d => ({ mayor: { referencia: d.op_ref, fecha: d.fecha_mayor, monto: d.monto_mayor, descripcion: d.desc_mayor }, th: { fecha: d.fecha_th, monto: d.monto_th, descripcion: d.desc_th }, dias: d.dias_diferencia, estado: d.estado })),
        pagos:       (pags  || []).map(p => ({ _dbId: p.id, mayor: { referencia: p.rm_ref, fecha: p.fecha_mayor, monto: p.monto_mayor, descripcion: p.desc_mayor }, th: { fecha: p.fecha_th, monto: p.monto_th }, op: p.minuta ? { minuta: p.minuta, seccRPA: p.secc_rpa, tarjetahabiente: p.tarjetahabiente, motivoPago: p.motivo_pago } : null, diferencia: parseFloat(p.diferencia || 0), lote: p.lote || '', estado: p.estado })),
        pendMayor:   (pends || []).filter(p => p.origen === 'mayor').map(p => ({ item: { tipo: p.tipo_movimiento, fecha: p.fecha, monto: p.monto, descripcion: p.descripcion, referencia: p.referencia }, razon: p.motivo })),
        pendCredits: (pends || []).filter(p => p.origen === 'th' && p.tipo_movimiento === 'CREDIT').map(p => ({ codRef: p.referencia, fecha: p.fecha, monto: p.monto, descripcion: p.descripcion })),
        pendDebits:  (pends || []).filter(p => p.origen === 'th' && p.tipo_movimiento === 'DEBIT').map(p => ({ codRef: p.referencia, fecha: p.fecha, monto: p.monto, descripcion: p.descripcion })),
      })
      setConcilId(cId)
    }
    setProcessing(false)
  }

  async function readFile(file, hint) {
    const buf = await file.arrayBuffer()
    const wb  = XLSX.read(buf, { type: 'array', cellDates: true })
    const tipo = detectFileType(wb)
    const slot = tipo === 'mayor' ? 0 : tipo === 'mov' ? 1 : tipo === 'operaciones' ? 2 : hint
    setWorkbooks(prev => { const n = [...prev]; n[slot] = wb; return n })
    setFileNames(prev => { const n = [...prev]; n[slot] = file.name; return n })
  }

  async function procesar() {
    setProcessing(true)
    try {
      const rendiciones = workbooks[3] ? parseRendiciones(workbooks[3], empresa) : {}
      const res = buildMatchesConLotes(parseMayor(workbooks[0]), parseMov(workbooks[1]), parseOperaciones(workbooks[2]), rendiciones)
      setResults(res)
      setActiveTab(0)
    } catch (err) { alert('Error: ' + err.message) }
    setProcessing(false)
  }

  async function guardar() {
    if (!results || !empresa) return
    setSaving(true)
    try {
      let cId = concilId
      if (!cId) {
        const { data, error } = await supabase.from('conciliaciones').insert({ empresa, periodo_desde: periodoDesde || null, periodo_hasta: periodoHasta || null, estado: 'borrador' }).select('id').single()
        if (error) throw error
        cId = data.id
        setConcilId(cId)
        navigate(`/conciliacion/${cId}`, { replace: true })
      }
      await Promise.all([
        supabase.from('depositos').delete().eq('conciliacion_id', cId),
        supabase.from('pagos').delete().eq('conciliacion_id', cId),
        supabase.from('pendientes').delete().eq('conciliacion_id', cId),
      ])
      if (results.depositos.length) {
        await supabase.from('depositos').insert(results.depositos.map(d => ({
          conciliacion_id: cId, op_ref: d.mayor.referencia, fecha_mayor: formatDate(d.mayor.fecha) || null,
          monto_mayor: d.mayor.monto, desc_mayor: d.mayor.descripcion,
          fecha_th: formatDate(d.th.fecha) || null, monto_th: d.th.monto, desc_th: d.th.descripcion,
          estado: d.estado || 'conciliado', dias_diferencia: d.dias || 0,
        })))
      }
      if (results.pagos.length) {
        const { data: pi } = await supabase.from('pagos').insert(results.pagos.map(p => ({
          conciliacion_id: cId, rm_ref: p.mayor.referencia, fecha_mayor: formatDate(p.mayor.fecha) || null,
          monto_mayor: p.mayor.monto, desc_mayor: p.mayor.descripcion,
          fecha_th: formatDate(p.th.fecha) || null, monto_th: p.th.monto,
          minuta: p.op?.minuta || null, secc_rpa: p.op?.seccRPA || null,
          tarjetahabiente: p.op?.tarjetahabiente || null, motivo_pago: p.op?.motivoPago || null,
          lote: p.lote || null, diferencia: p.diferencia || 0, estado: p.estado || 'sin_lote',
        }))).select('id')
        if (pi) setResults(prev => ({ ...prev, pagos: prev.pagos.map((p, i) => ({ ...p, _dbId: pi[i]?.id })) }))
      }
      const pendAll = [
        ...results.pendMayor.map(p => ({ conciliacion_id: cId, origen: 'mayor', tipo_movimiento: p.item.tipo, fecha: formatDate(p.item.fecha) || null, monto: p.item.monto, descripcion: p.item.descripcion, referencia: p.item.referencia, motivo: p.razon })),
        ...results.pendCredits.map(p => ({ conciliacion_id: cId, origen: 'th', tipo_movimiento: 'CREDIT', fecha: formatDate(p.fecha) || null, monto: p.monto, descripcion: p.descripcion, referencia: p.codRef })),
        ...results.pendDebits.map(p => ({ conciliacion_id: cId, origen: 'th', tipo_movimiento: 'DEBIT', fecha: formatDate(p.fecha) || null, monto: p.monto, descripcion: p.descripcion, referencia: p.codRef })),
      ]
      if (pendAll.length) await supabase.from('pendientes').insert(pendAll)
      alert('✓ Guardado correctamente')
    } catch (err) { alert('Error al guardar: ' + err.message); console.error(err) }
    setSaving(false)
  }

  async function updateLote(idx, value) {
    const pago = results.pagos[idx]
    const estado = value.trim() ? (Math.abs(pago.diferencia) > 50 ? 'revisar' : 'ok') : 'sin_lote'
    setResults(prev => ({ ...prev, pagos: prev.pagos.map((p, i) => i === idx ? { ...p, lote: value, estado } : p) }))
    if (pago._dbId) {
      await supabase.from('pagos').update({ lote: value || null, estado, updated_at: new Date().toISOString() }).eq('id', pago._dbId)
    }
  }

  function exportExcel() {
    if (!results) return
    const { depositos, pagos, pendMayor, pendCredits, pendDebits } = results
    const wb = XLSX.utils.book_new()

    // ── HOJA 1: DEB-CRED ─────────────────────────────────────
    // Replica la hoja DEB-CRED del archivo de conciliación acumulada.
    // Columna G = MOVIMIENTO: para CREDIT que matchearon con un OP se pone el número de OP.
    // Para CREDIT sin match: "OP A CONCILIAR". Para DEBIT: vacío.
    const debCredRows = [
      ['Cod Referencia', 'Fecha Acreditación', 'Descripción', 'Tipo', 'Monto', 'Disponible', 'MOVIMIENTO (OP)'],
      // CREDITs conciliados → columna MOVIMIENTO = número de OP del Mayor
      ...depositos.map(r => [
        r.th.codRef || '',
        formatDate(r.th.fecha),
        r.th.descripcion || '',
        'CREDIT',
        r.th.monto,
        '',
        r.mayor.referencia,   // ← acá va el OP-XXXXXX en la columna MOVIMIENTO
      ]),
      // DEBITs conciliados → columna MOVIMIENTO vacía
      ...pagos.map(r => [
        r.th.codRef || '',
        formatDate(r.th.fecha),
        r.th.descripcion || '',
        'DEBIT',
        r.th.monto,
        '',
        '',
      ]),
      // DEBITs sin match
      ...pendDebits.map(r => [
        r.codRef || '',
        formatDate(r.fecha),
        r.descripcion || '',
        'DEBIT',
        r.monto,
        '',
        '',
      ]),
      // CREDITs sin match → "OP A CONCILIAR"
      ...pendCredits.map(r => [
        r.codRef || '',
        formatDate(r.fecha),
        r.descripcion || '',
        'CREDIT',
        r.monto,
        '',
        'OP A CONCILIAR',
      ]),
    ]
    const ws1 = XLSX.utils.aoa_to_sheet(debCredRows)
    ws1['!cols'] = [{wch:16},{wch:14},{wch:60},{wch:8},{wch:18},{wch:16},{wch:18}]
    XLSX.utils.book_append_sheet(wb, ws1, 'DEB-CRED')

    // ── HOJA 2: MOV (Operaciones de Pago) ────────────────────
    // Replica la hoja MOV. Columna P = LOTE asignado por el usuario.
    const movRows = [
      ['Nro Operación', 'Nro Pago', 'Fecha', 'Total', 'Secc. RPA', 'Tarjetahabiente',
       'Estado', 'Motivo Pago', 'Minuta', 'Diferencia', 'LOTE'],
      ...pagos.map(r => [
        r.op?.nroOp   || '',
        r.op?.nroPago || '',
        formatDate(r.th.fecha),
        r.th.monto,
        r.op?.seccRPA          || '',
        r.op?.tarjetahabiente  || '',
        r.op?.estado           || '',
        r.op?.motivoPago       || '',
        r.op?.minuta           || '',
        r.diferencia,
        r.lote || '',           // ← columna LOTE (P) que se va completando desde CELER
      ]),
      ...pendDebits.map(r => [
        '', r.pagoNum || '',
        formatDate(r.fecha), r.monto,
        '', '', '', '', '', '', '',
      ]),
    ]
    const ws2 = XLSX.utils.aoa_to_sheet(movRows)
    ws2['!cols'] = [{wch:14},{wch:12},{wch:12},{wch:16},{wch:28},{wch:30},{wch:12},{wch:55},{wch:12},{wch:14},{wch:12}]
    XLSX.utils.book_append_sheet(wb, ws2, 'MOV')

    // ── HOJA 3: MAYOR ─────────────────────────────────────────
    // Replica las filas relevantes del Mayor con columna L completada:
    //   OP → columna L = "OP-XXXXXX"
    //   RM → columna L = número de lote (cuando esté asignado) o número TH
    const mayorRows = [
      ['Fecha', 'Conciliado', 'Comprobante', 'Pagos', 'Depósitos', 'Balance', '',
       '', '', '', '', 'COL L (Ref / Lote)'],
      ...depositos.map(r => [
        formatDate(r.mayor.fecha), '', r.mayor.descripcion || '',
        '', r.mayor.monto, '', '', '', '', '', '',
        r.mayor.referencia,       // ← OP-XXXXXX en columna L
      ]),
      ...pagos.map(r => [
        formatDate(r.mayor.fecha), '', r.mayor.descripcion || '',
        r.mayor.monto, '', '', '', '', '', '', '',
        r.lote || r.mayor.referencia || '',  // ← Lote si está asignado, sino ref TH
      ]),
      ...pendMayor.map(r => [
        formatDate(r.item.fecha), '', r.item.descripcion || '',
        r.item.tipo === 'RM' ? r.item.monto : '',
        r.item.tipo === 'OP' ? r.item.monto : '',
        '', '', '', '', '', '',
        '⚠ SIN MATCH EN TH',
      ]),
    ]
    const ws3 = XLSX.utils.aoa_to_sheet(mayorRows)
    ws3['!cols'] = [{wch:14},{wch:14},{wch:65},{wch:18},{wch:18},{wch:18},{wch:4},{wch:4},{wch:4},{wch:4},{wch:4},{wch:20}]
    XLSX.utils.book_append_sheet(wb, ws3, 'Mayor')

    // ── HOJA 4: Pendientes ────────────────────────────────────
    const pendRows = [
      ['Origen', 'Tipo', 'Fecha', 'Monto', 'Descripción', 'Ref', 'Motivo'],
      ...pendMayor.map(p => ['Mayor', p.item.tipo, formatDate(p.item.fecha), p.item.monto, p.item.descripcion, p.item.referencia, p.razon]),
      ...pendCredits.map(p => ['TH', 'CREDIT', formatDate(p.fecha), p.monto, p.descripcion, p.codRef, 'OP a conciliar']),
      ...pendDebits.map(p => ['TH', 'DEBIT', formatDate(p.fecha), p.monto, p.descripcion, p.codRef, '']),
    ]
    const ws4 = XLSX.utils.aoa_to_sheet(pendRows)
    XLSX.utils.book_append_sheet(wb, ws4, 'Pendientes')

    XLSX.writeFile(wb, `Conciliacion_${empresa}_${new Date().toISOString().slice(0,10)}.xlsx`)
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

          {/* ARCHIVOS POR EMPRESA */}
          <div style={{ fontSize: '.7rem', fontWeight: 700, color: 'var(--txt2)', textTransform: 'uppercase', letterSpacing: '.07em', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ background: 'var(--green)', color: '#000', borderRadius: 4, padding: '1px 7px', fontSize: '.65rem' }}>1 por empresa</span>
            Se descargan individualmente para la empresa seleccionada
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 12, marginBottom: 20 }}>
            {[
              { label: 'Mayor Autodealer (K1)', icon: '📊', slot: 0, hint: 'K1 → Finanzas → Resumen de Cuenta → Export Excel 5.0' },
              { label: 'Movimientos de Cuenta TH', icon: '💳', slot: 1, hint: 'Web TH → Movimientos de Cuenta → Exportar' },
              { label: 'Operaciones de Pago TH', icon: '📋', slot: 2, hint: 'Web TH → Operaciones de Pago → Exportar' },
            ].map(({ label, icon, slot, hint }) => (
              <div
                key={slot}
                onDragOver={e => e.preventDefault()}
                onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) readFile(f, slot) }}
                style={{
                  border: `1.5px dashed ${workbooks[slot] ? 'var(--green)' : 'var(--bdr2)'}`,
                  background: workbooks[slot] ? 'var(--green-dim)' : '#111',
                  borderRadius: 10, padding: '20px 16px', textAlign: 'center',
                  cursor: 'pointer', position: 'relative',
                }}
              >
                <input type="file" accept=".xlsx,.xls"
                  style={{ position: 'absolute', inset: 0, opacity: 0, cursor: 'pointer', width: '100%', height: '100%' }}
                  onChange={e => e.target.files[0] && readFile(e.target.files[0], slot)} />
                <div style={{ fontSize: '1.5rem', marginBottom: 7 }}>{icon}</div>
                <div style={{ fontSize: '.8rem', fontWeight: 700, marginBottom: 4 }}>{label}</div>
                <div style={{ fontSize: '.67rem', color: 'var(--txt3)', marginBottom: 8 }}>{hint}</div>
                {workbooks[slot]
                  ? <div style={{ fontSize: '.7rem', color: 'var(--green)', fontWeight: 600 }}>✓ {fileNames[slot]}</div>
                  : <div style={{ fontSize: '.7rem', color: 'var(--txt2)' }}>Arrastrá o hacé click</div>
                }
              </div>
            ))}
          </div>

          {/* RENDICIONES CRM — un solo archivo para todo el período */}
          <div style={{ border: '1px solid var(--bdr2)', borderRadius: 10, padding: '16px 18px', marginBottom: 20, background: '#0e0e0e' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
              <span style={{ background: '#1A73E8', color: '#fff', borderRadius: 4, padding: '1px 7px', fontSize: '.65rem', fontWeight: 700 }}>Todo el período</span>
              <div>
                <div style={{ fontSize: '.78rem', fontWeight: 700 }}>Rendiciones CRM</div>
                <div style={{ fontSize: '.68rem', color: 'var(--txt2)' }}>Un solo archivo · Incluye todas las empresas · CRM → Rendiciones → Reporte del período</div>
              </div>
            </div>
            <div
              onDragOver={e => e.preventDefault()}
              onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) readFile(f, 3) }}
              style={{
                border: `1.5px dashed ${workbooks[3] ? 'var(--green)' : '#333'}`,
                background: workbooks[3] ? 'var(--green-dim)' : '#111',
                borderRadius: 8, padding: '16px', textAlign: 'center',
                cursor: 'pointer', position: 'relative',
              }}
            >
              <input type="file" accept=".xlsx,.xls,.csv"
                style={{ position: 'absolute', inset: 0, opacity: 0, cursor: 'pointer', width: '100%', height: '100%' }}
                onChange={e => e.target.files[0] && readFile(e.target.files[0], 3)} />
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10 }}>
                <span style={{ fontSize: '1.4rem' }}>🗂️</span>
                <div style={{ textAlign: 'left' }}>
                  {workbooks[3]
                    ? <div style={{ fontSize: '.78rem', color: 'var(--green)', fontWeight: 700 }}>✓ {fileNames[3]} — Lotes asignados automáticamente</div>
                    : <>
                        <div style={{ fontSize: '.78rem', fontWeight: 600 }}>Arrastrá el archivo o hacé click para subir</div>
                        <div style={{ fontSize: '.68rem', color: 'var(--txt2)', marginTop: 2 }}>Opcional · Si no lo subís, podés asignar los lotes a mano</div>
                      </>
                  }
                </div>
              </div>
            </div>
          </div>

          <button
            onClick={procesar}
            disabled={!workbooks[0] || !workbooks[1] || !workbooks[2] || !empresa || processing}
            style={{ ...btnGreen, opacity: (!workbooks.every(Boolean) || !empresa || processing) ? .4 : 1 }}
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
          onNueva={() => { setResults(null); setWorkbooks([null,null,null,null]); setFileNames(['','','','']); setConcilId(null); navigate('/nueva') }}
          onUpdateLote={updateLote}
        />
      )}
    </div>
  )
}

function ResultsSection({ results, empresa, concilId, saving, activeTab, setActiveTab, onGuardar, onExport, onNueva, onUpdateLote }) {
  const { depositos, pagos, pendMayor, pendCredits, pendDebits } = results
  const totDep   = depositos.reduce((s, r) => s + r.mayor.monto, 0)
  const totPag   = pagos.reduce((s, r) => s + r.mayor.monto, 0)
  const sinLote  = pagos.filter(p => !p.lote).length
  const conDiff  = pagos.filter(p => Math.abs(p.diferencia) > 50).length
  const pendTotal = pendMayor.length + pendCredits.length + pendDebits.length

  const kpis = [
    { label: 'Depósitos conciliados',  value: depositos.length,  sub: `$ ${formatMonto(totDep)}`,  color: 'var(--green)' },
    { label: 'Pagos conciliados',      value: pagos.length,      sub: `$ ${formatMonto(totPag)}`,  color: 'var(--green)' },
    { label: 'Sin lote asignado',      value: sinLote,           sub: sinLote ? 'Completar con CELER' : 'Todos asignados', color: sinLote ? 'var(--amber)' : 'var(--green)' },
    { label: 'Diferencias > $50',      value: conDiff,           sub: conDiff ? 'Revisar urgente' : 'Sin diferencias',    color: conDiff ? 'var(--red)' : 'var(--green)' },
    { label: 'Pendientes',             value: pendTotal,         sub: pendCredits.length ? `${pendCredits.length} OP a conciliar` : 'Sin pendientes', color: pendTotal ? 'var(--amber)' : 'var(--green)' },
  ]

  const tabs = [
    `💰 Depósitos (${depositos.length})`,
    `📋 Pagos y Lotes (${pagos.length})`,
    `⚠ Pendientes (${pendTotal})`,
  ]

  return (
    <>
      {/* ACCIONES */}
      <div style={{ display: 'flex', gap: 10, marginBottom: 24, flexWrap: 'wrap', alignItems: 'center' }}>
        <button onClick={onGuardar} disabled={saving} style={btnGreen}>
          {saving ? '⏳ Guardando…' : concilId ? '💾 Actualizar' : '💾 Guardar'}
        </button>
        <button onClick={onExport} style={btnOutline}>⬇ Descargar Excel</button>
        <button onClick={onNueva} style={btnOutline}>↺ Nueva</button>
        {concilId && <span style={{ fontSize: '.72rem', color: 'var(--txt2)' }}>Los lotes se guardan automáticamente</span>}
      </div>

      {/* KPIs */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5,1fr)', gap: 12, marginBottom: 24 }}>
        {kpis.map(k => (
          <div key={k.label} style={{ background: 'var(--surf)', borderRadius: 9, padding: '16px', border: '1px solid var(--bdr)', borderLeft: `3px solid ${k.color}` }}>
            <div style={{ fontSize: '.65rem', color: 'var(--txt2)', textTransform: 'uppercase', letterSpacing: '.05em', marginBottom: 6 }}>{k.label}</div>
            <div style={{ fontSize: '1.4rem', fontWeight: 900, color: k.color, fontVariantNumeric: 'tabular-nums' }}>{k.value}</div>
            <div style={{ fontSize: '.7rem', color: 'var(--txt2)', marginTop: 3 }}>{k.sub}</div>
          </div>
        ))}
      </div>

      {/* TABS */}
      <div style={{ display: 'flex', gap: 2, borderBottom: '1px solid var(--bdr)', marginBottom: 20 }}>
        {tabs.map((label, i) => (
          <button key={i} onClick={() => setActiveTab(i)} style={{
            padding: '9px 18px', background: 'none', border: 'none', cursor: 'pointer',
            fontSize: '.8rem', fontWeight: 600,
            color: activeTab === i ? 'var(--green)' : 'var(--txt2)',
            borderBottom: `2px solid ${activeTab === i ? 'var(--green)' : 'transparent'}`,
            marginBottom: -1,
          }}>{label}</button>
        ))}
      </div>

      {/* TAB 0: DEPÓSITOS */}
      {activeTab === 0 && (
        <>
          {pendCredits.length > 0 && (
            <div style={{ background: 'var(--amber-dim)', border: '1px solid rgba(245,158,11,.3)', borderRadius: 8, padding: '12px 16px', marginBottom: 16, fontSize: '.78rem', color: 'var(--amber)' }}>
              <strong>⚠ {pendCredits.length} OP a conciliar</strong> — CREDIT en TH sin OP en el Mayor. Registrarlos en celda B7 de la pestaña Conciliación.
            </div>
          )}
          <TableBox>
            <thead><tr>
              {['Ref. OP','Fecha Mayor','Monto Mayor','Fecha TH','Descripción TH','Monto TH','Días','Estado'].map(h => <th key={h} style={th}>{h}</th>)}
            </tr></thead>
            <tbody>
              {depositos.map((r, i) => (
                <tr key={i} style={{ borderBottom: '1px solid var(--bdr)' }}>
                  <td style={td}><strong style={{ color: 'var(--green)' }}>{r.mayor.referencia}</strong></td>
                  <td style={td}>{formatDate(r.mayor.fecha)}</td>
                  <td style={{ ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{formatMonto(r.mayor.monto)}</td>
                  <td style={td}>{formatDate(r.th.fecha)}</td>
                  <td style={{ ...td, maxWidth: 260, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.th.descripcion}</td>
                  <td style={{ ...td, textAlign: 'right' }}>{formatMonto(r.th.monto)}</td>
                  <td style={{ ...td, textAlign: 'right', color: r.dias > 3 ? 'var(--amber)' : 'var(--green)', fontWeight: 600 }}>{r.dias}</td>
                  <td style={td}><span style={pill('var(--green-dim)', 'var(--green)')}>Conciliado</span></td>
                </tr>
              ))}
              {pendCredits.map((r, i) => (
                <tr key={'c'+i} style={{ borderBottom: '1px solid var(--bdr)' }}>
                  <td style={{ ...td, color: 'var(--txt2)' }}>—</td>
                  <td style={{ ...td, color: 'var(--txt2)' }}>—</td>
                  <td style={{ ...td, color: 'var(--txt2)', textAlign: 'right' }}>—</td>
                  <td style={td}>{formatDate(r.fecha)}</td>
                  <td style={{ ...td, maxWidth: 260, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.descripcion}</td>
                  <td style={{ ...td, textAlign: 'right', color: 'var(--red)', fontWeight: 600 }}>{formatMonto(r.monto)}</td>
                  <td style={{ ...td, color: 'var(--txt2)' }}>—</td>
                  <td style={td}><span style={pill('var(--amber-dim)', 'var(--amber)')}>OP a conciliar</span></td>
                </tr>
              ))}
            </tbody>
          </TableBox>
        </>
      )}

      {/* TAB 1: PAGOS Y LOTES */}
      {activeTab === 1 && (
        <>
          {sinLote > 0 && (
            <div style={{ background: 'var(--amber-dim)', border: '1px solid rgba(245,158,11,.3)', borderRadius: 8, padding: '12px 16px', marginBottom: 16, fontSize: '.78rem', color: 'var(--amber)' }}>
              <strong>ℹ {sinLote} pago{sinLote > 1 ? 's' : ''} sin lote.</strong> Completá la columna Lote con el número de rendición de CELER.
              {!concilId && <strong> Guardá primero para que los lotes se auto-guarden.</strong>}
            </div>
          )}
          <TableBox>
            <thead><tr>
              {['Fecha Mayor','RM Ref.','Monto Mayor','Minuta','Seccional','Tarjetahabiente','Fecha TH','Monto TH','Diferencia','Lote'].map(h => <th key={h} style={th}>{h}</th>)}
            </tr></thead>
            <tbody>
              {pagos.map((r, i) => (
                <tr key={i} style={{ borderBottom: '1px solid var(--bdr)' }}>
                  <td style={td}>{formatDate(r.mayor.fecha)}</td>
                  <td style={td}>{r.mayor.referencia}</td>
                  <td style={{ ...td, textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{formatMonto(r.mayor.monto)}</td>
                  <td style={td}><strong style={{ color: 'var(--green)' }}>{r.op?.minuta || '—'}</strong></td>
                  <td style={{ ...td, fontSize: '.7rem', color: 'var(--txt2)' }}>{r.op?.seccRPA || ''}</td>
                  <td style={{ ...td, fontSize: '.7rem', whiteSpace: 'nowrap', color: 'var(--txt2)' }}>{r.op?.tarjetahabiente || ''}</td>
                  <td style={td}>{formatDate(r.th.fecha)}</td>
                  <td style={{ ...td, textAlign: 'right' }}>{formatMonto(r.th.monto)}</td>
                  <td style={{ ...td, textAlign: 'right', fontWeight: 600, color: Math.abs(r.diferencia) > 50 ? 'var(--red)' : Math.abs(r.diferencia) > 0 ? 'var(--amber)' : 'var(--green)' }}>
                    {Math.abs(r.diferencia) < 0.01 ? '—' : formatMonto(r.diferencia)}
                  </td>
                  <td style={td}>
                    <input
                      value={r.lote}
                      onChange={e => onUpdateLote(i, e.target.value)}
                      placeholder="Lote"
                      style={{
                        width: 80, background: r.lote ? 'var(--green-dim)' : '#111',
                        border: `1px solid ${r.lote ? 'var(--green)' : 'var(--amber)'}`,
                        color: '#fff', borderRadius: 5, padding: '4px 7px',
                        fontSize: '.75rem', fontWeight: 600, outline: 'none',
                      }}
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </TableBox>
        </>
      )}

      {/* TAB 2: PENDIENTES */}
      {activeTab === 2 && (
        pendTotal === 0
          ? <div style={{ textAlign: 'center', padding: '48px 20px', color: 'var(--txt2)' }}>
              <div style={{ fontSize: '2rem', marginBottom: 10 }}>✅</div>
              Sin pendientes — todo conciliado.
            </div>
          : <>
            {pendMayor.length > 0 && (
              <Section label={`📌 En Mayor sin match en TH (${pendMayor.length})`}>
                <TableBox>
                  <thead><tr>{['Tipo','Fecha','Descripción','Ref.','Monto','Motivo'].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
                  <tbody>{pendMayor.map((r, i) => (
                    <tr key={i} style={{ borderBottom: '1px solid var(--bdr)' }}>
                      <td style={td}><span style={pill('var(--amber-dim)', 'var(--amber)')}>{r.item.tipo}</span></td>
                      <td style={td}>{formatDate(r.item.fecha)}</td>
                      <td style={{ ...td, maxWidth: 280 }}>{r.item.descripcion}</td>
                      <td style={td}>{r.item.referencia}</td>
                      <td style={{ ...td, textAlign: 'right', color: 'var(--amber)', fontWeight: 600 }}>{formatMonto(r.item.monto)}</td>
                      <td style={{ ...td, fontSize: '.72rem', color: 'var(--txt2)' }}>{r.razon}</td>
                    </tr>
                  ))}</tbody>
                </TableBox>
              </Section>
            )}
            {pendDebits.length > 0 && (
              <Section label={`🔴 DEBIT en TH sin match en Mayor (${pendDebits.length})`}>
                <TableBox>
                  <thead><tr>{['Cod Ref','Fecha','Descripción','Monto'].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
                  <tbody>{pendDebits.map((r, i) => (
                    <tr key={i} style={{ borderBottom: '1px solid var(--bdr)' }}>
                      <td style={td}>{r.codRef}</td>
                      <td style={td}>{formatDate(r.fecha)}</td>
                      <td style={{ ...td, maxWidth: 320 }}>{r.descripcion}</td>
                      <td style={{ ...td, textAlign: 'right', color: 'var(--red)', fontWeight: 600 }}>{formatMonto(r.monto)}</td>
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

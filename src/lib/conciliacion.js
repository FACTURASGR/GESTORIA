import * as XLSX from 'xlsx'

// ── UTILIDADES ─────────────────────────────────────────────

export function parseDate(v) {
  if (!v) return null
  if (v instanceof Date) return isNaN(v) ? null : v
  const s = String(v).trim()
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/)
  if (m) {
    const y = m[3].length === 2 ? '20' + m[3] : m[3]
    return new Date(+y, +m[2] - 1, +m[1])
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (m) return new Date(+m[1], +m[2] - 1, +m[3])
  const d = new Date(v)
  return isNaN(d) ? null : d
}

// Fecha en formato AAAA-MM-DD (lo que espera una columna DATE de Supabase)
export function isoDate(v) {
  const d = parseDate(v)
  if (!d) return null
  const z = n => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`
}

export function formatDate(v) {
  const d = parseDate(v)
  if (!d) return ''
  return d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' })
}

export function formatMonto(n) {
  if (n === null || n === undefined || n === '') return ''
  const x = parseFloat(n)
  if (isNaN(x)) return ''
  return x.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

export function daysDiff(a, b) {
  const da = parseDate(a), db = parseDate(b)
  if (!da || !db) return 999
  return Math.abs((da - db) / 86400000)
}

function sheetToRows(wb, idx = 0) {
  const ws = wb.Sheets[wb.SheetNames[idx]]
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null })
}

// ── DETECCIÓN AUTOMÁTICA DE TIPO DE ARCHIVO ────────────────

export function detectFileType(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const s = rows[i].join('|').toLowerCase()
    if (s.includes('lote rendicion') || s.includes('num_minuta')) return 'rendiciones'
    if (s.includes('nro. operacion') || (s.includes('nro de pago') && s.includes('canal'))) return 'operaciones'
    if (s.includes('cod referencia') && s.includes('tipo')) return 'mov'
    if (s.includes('comprobante') && (s.includes('pagos') || s.includes('dep'))) return 'mayor'
  }
  return 'desconocido'
}

// ── PARSER: MAYOR (K1 — Autodealer) ───────────────────────

export function parseMayor(wb) {
  const rows = sheetToRows(wb)
  let headerIdx = -1
  let iF = -1, iD = -1, iP = -1, iDep = -1

  for (let i = 0; i < Math.min(25, rows.length); i++) {
    const row = rows[i]
    let hasComp = false, hasPagos = false
    row.forEach((c, j) => {
      const s = String(c || '').toLowerCase().trim()
      if (s === 'fecha' && iF < 0) iF = j
      if (s === 'comprobante') { iD = j; hasComp = true }
      if (s === 'pagos') { iP = j; hasPagos = true }
      if (s === 'depósitos' || s === 'depositos') iDep = j
    })
    if (hasComp && hasPagos) { headerIdx = i; break }
  }

  if (headerIdx < 0) throw new Error('No se encontró el encabezado del Mayor. Verificá que sea el archivo correcto.')

  // El header "Comprobante" puede estar desalineado por celdas fusionadas en el Excel.
  // Escaneamos las primeras filas de datos para encontrar la columna real
  // donde aparecen RM-, OP-, RC- o SALDO (patron confiable del Mayor de Autodealer).
  let iD_real = -1
  for (let i = headerIdx + 1; i < Math.min(headerIdx + 30, rows.length); i++) {
    const row = rows[i]
    for (let j = 0; j < row.length; j++) {
      const v = String(row[j] || '')
      if (v.match(/^(RM|OP|RC|SALDO)/i)) { iD_real = j; break }
    }
    if (iD_real >= 0) break
  }
  if (iD_real >= 0) iD = iD_real  // datos reales ganan sobre header desalineado

  const items = []
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i]
    const comp = row[iD]
    if (!comp || typeof comp !== 'string') continue

    const desc = comp.trim()
    if (!desc.match(/^(RM|OP)-/i)) continue

    const tipo = desc.toUpperCase().startsWith('OP') ? 'OP' : 'RM'
    const monto = parseFloat(tipo === 'OP' ? row[iDep] : row[iP]) || 0
    if (!monto) continue

    let ref = ''
    if (tipo === 'OP') {
      const m = desc.match(/^(OP-\d+)/i)
      ref = m ? m[1].toUpperCase() : ''
    } else {
      const m = desc.match(/TARJETA HABITUALISTA N\s*(\d+)/i)
      ref = m ? m[1] : ''
    }

    items.push({ tipo, descripcion: desc, referencia: ref, fecha: row[iF], monto, fila: i })
  }
  return items
}

// ── PARSER: MOVIMIENTOS DE CUENTA TH ──────────────────────

export function parseMov(wb) {
  const rows = sheetToRows(wb)
  let headerIdx = 0

  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const s = rows[i].join('|').toLowerCase()
    if (s.includes('cod referencia') || (s.includes('tipo') && s.includes('monto'))) {
      headerIdx = i; break
    }
  }

  const hdr = rows[headerIdx].map(h => String(h || '').toLowerCase().trim())
  const iCod  = hdr.findIndex(h => h.includes('cod'))
  const iFecha = hdr.findIndex(h => h.includes('fecha'))
  const iDesc  = hdr.findIndex(h => h.includes('descripci'))
  const iTipo  = hdr.findIndex(h => h === 'tipo')
  const iMonto = hdr.findIndex(h => h === 'monto')
  const iDisp  = hdr.findIndex(h => h.includes('disponible'))

  const items = []
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i]
    const tipo = String(row[iTipo] || '').toUpperCase()
    if (tipo !== 'CREDIT' && tipo !== 'DEBIT') continue

    const desc = String(row[iDesc] || '')
    let pagoNum = null
    if (tipo === 'DEBIT') {
      const m = desc.match(/Pago:\s*0*(\d+)/i)
      if (m) pagoNum = parseInt(m[1])
    }

    const monto = parseFloat(row[iMonto]) || 0
    if (!monto) continue

    items.push({
      codRef: String(row[iCod] || ''),
      fecha: row[iFecha],
      descripcion: desc,
      tipo,
      monto,
      pagoNum,
      disponible: isNaN(parseFloat(row[iDisp])) ? '' : parseFloat(row[iDisp]),
      _usado: false,
    })
  }
  return items
}

// ── PARSER: OPERACIONES DE PAGO TH ────────────────────────

// Analiza el campo "Motivo pago". La minuta es un numero de 6 digitos, escrito de formas muy distintas:
//   "410019 okm bruno"  |  "OKM AI356RH RAVAROTTO MINUTA 410680"  |  "Minuta: 406648 Tramite..."
// Devuelve { minuta, estado, candidatos }
//   'ok'      -> hay un unico numero de 6 digitos
//   'manual'  -> no hay ninguno (ej. "AUTORIZACION EPAGOS - Seccional: 99997"): la persona lo carga a mano
//   'ambigua' -> hay varios numeros de 6 digitos distintos: la persona elige cual es
export function analizarMotivo(motivo) {
  const t = String(motivo || '')
  const candidatos = []
  const re = /(^|\D)(\d{6})(?!\d)/g
  let m
  while ((m = re.exec(t)) !== null) {
    if (!candidatos.includes(m[2])) candidatos.push(m[2])
  }
  if (candidatos.length === 1) return { minuta: candidatos[0], estado: 'ok', candidatos }
  if (candidatos.length === 0) return { minuta: '', estado: 'manual', candidatos }
  return { minuta: '', estado: 'ambigua', candidatos }
}

export function extraerMinuta(motivo) { return analizarMotivo(motivo).minuta }

// Devuelve las operaciones en el MISMO ORDEN del archivo (lo mas nuevo arriba), una por fila
export function parseOperaciones(wb) {
  const rows = sheetToRows(wb)
  let headerIdx = 0

  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const s = rows[i].join('|').toLowerCase()
    if (s.includes('nro. operacion') || s.includes('nro de pago')) { headerIdx = i; break }
  }

  const hdr = rows[headerIdx].map(h => String(h || '').toLowerCase().trim())
  const iOp     = hdr.findIndex(h => h.includes('operacion'))
  const iPago   = hdr.findIndex(h => h.includes('nro de pago'))
  const iCanal  = hdr.findIndex(h => h === 'canal')
  const iAcred  = hdr.findIndex(h => h.includes('acreditacion'))
  const iCupon  = hdr.findIndex(h => h.includes('cupon'))
  const iFecha  = hdr.findIndex(h => h === 'fecha')
  const iTotal  = hdr.findIndex(h => h === 'total')
  const iSecc   = hdr.findIndex(h => h.includes('secc') || h.includes('rpa'))
  const iTH     = hdr.findIndex(h => h.includes('tarjetahabiente'))
  const iTarj   = hdr.findIndex(h => h === 'tarjeta')
  const iEst    = hdr.findIndex(h => h === 'estado')
  const iMot    = hdr.findIndex(h => h.includes('motivo'))
  const iObs    = hdr.findIndex(h => h.includes('observaci'))

  const lista = []
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i]
    if (!row[iOp]) continue
    const motivo = String(row[iMot] || '')
    const a = analizarMotivo(motivo)
    lista.push({
      nroOp:           row[iOp],
      nroPago:         row[iPago] ? parseInt(row[iPago]) : null,
      canal:           String(row[iCanal] || ''),
      acreditacion:    String(row[iAcred] || ''),
      nroCupon:        row[iCupon] ?? '',
      fecha:           row[iFecha],
      total:           parseFloat(row[iTotal]) || 0,
      seccRPA:         String(row[iSecc] || ''),
      tarjetahabiente: String(row[iTH]   || ''),
      tarjeta:         String(row[iTarj] || ''),
      estado:          String(row[iEst]  || ''),
      motivoPago:      motivo,
      observaciones:   String(row[iObs]  || ''),
      minuta:          a.minuta,
      minutaEstado:    a.estado,
      minutaCandidatos: a.candidatos,
    })
  }
  return lista
}

// ── PARSER: RENDICIONES (Autodealer) ──────────────────────
// Archivo CSV disfrazado de .xls con separador ;
// Columnas: SISTEMA | NUM_MINUTA | GESTOR | FECHA ALTA MINUTA | FECHA RENDICION | LOTE RENDICION | REAL RENDIDO
// Retorna un mapa { "NUM_MINUTA": "LOTE RENDICION" } filtrado por empresa

export function parseRendiciones(wb, empresa) {
  const ws  = wb.Sheets[wb.SheetNames[0]]
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null })

  // Buscar la fila de headers (contiene LOTE RENDICION o NUM_MINUTA)
  let headerIdx = 0
  for (let i = 0; i < Math.min(5, raw.length); i++) {
    const s = (raw[i] || []).join('|').toUpperCase()
    if (s.includes('NUM_MINUTA') || s.includes('LOTE')) { headerIdx = i; break }
  }

  const hdr      = (raw[headerIdx] || []).map(h => String(h || '').trim().toUpperCase())
  const iSistema = hdr.findIndex(h => h.includes('SISTEMA'))
  const iMinuta  = hdr.findIndex(h => h.includes('NUM_MINUTA') || h.includes('MINUTA'))
  const iLote    = hdr.findIndex(h => h.includes('LOTE'))
  const iMonto   = hdr.findIndex(h => h.includes('REAL') || h.includes('RENDIDO'))

  // mapa minuta → { lote, monto, sistema }
  const mapa = {}
  const empresaNorm = (empresa || '').toUpperCase().trim()

  for (let i = headerIdx + 1; i < raw.length; i++) {
    const row     = raw[i]
    if (!row || !row[iMinuta]) continue

    const sistema = String(row[iSistema] || '').toUpperCase().trim()
    const minuta  = String(row[iMinuta]  || '').trim()
    const lote    = String(row[iLote]    || '').trim()
    const monto   = parseFloat(row[iMonto]) || 0

    if (!minuta || !lote) continue

    // Si se pasa empresa, filtrar; si no se pasa, incluir todo
    if (empresaNorm && sistema !== empresaNorm) continue

    mapa[minuta] = { lote, monto, sistema }
  }

  return mapa   // { "413713": { lote: "29801", monto: 344834.60, sistema: "CIARA" } }
}

// Total y cantidad de minutas por lote, segun el reporte de Rendiciones (CRM)
export function totalesPorLote(rendMap) {
  const out = {}
  for (const v of Object.values(rendMap || {})) {
    if (!out[v.lote]) out[v.lote] = { total: 0, n: 0 }
    out[v.lote].total += v.monto
    out[v.lote].n += 1
  }
  for (const k of Object.keys(out)) out[k].total = round2(out[k].total)
  return out
}

// ── CRUCE MOV (Operaciones) ↔ RENDICIONES (CRM) ────────────
const round2 = n => Math.round(n * 100) / 100
const TOLERANCIA = 50   // diferencias mayores a $50 se revisan (instructivo)

// Completa para cada pago: lote (segun la minuta), importe rendido, diferencia y estado.
// Si la persona cargo a mano la minuta (row.minuta) o el lote (row.loteManual), se respeta y se recalcula todo.
export function recalcularMov(rows, rendMap) {
  const hayCRM = !!rendMap
  const out = rows.map(r => {
    const minuta = String(r.minuta || '').trim()
    const rend = hayCRM && minuta ? rendMap[minuta] : null
    const manual = r.loteManual !== undefined && r.loteManual !== null && String(r.loteManual).trim() !== ''
    const loteAuto = rend ? rend.lote : ''
    return {
      ...r, minuta, loteAuto,
      lote: manual ? String(r.loteManual).trim() : loteAuto,
      rendido: rend ? rend.monto : null,
      diferencia: null, pagoMultiple: false, duplicado: false,
    }
  })

  // agrupar los pagos de una misma minuta
  const grupos = new Map()
  out.forEach((r, i) => {
    if (!r.minuta) return
    if (!grupos.has(r.minuta)) grupos.set(r.minuta, [])
    grupos.get(r.minuta).push(i)
  })
  for (const idxs of grupos.values()) {
    const suma = idxs.reduce((s, i) => s + out[i].total, 0)
    const rendido = out[idxs[0]].rendido
    idxs.forEach(i => { out[i].pagoMultiple = idxs.length > 1 })
    // la diferencia es de la minuta completa (puede estar paga en varias veces, ej. 0KM): va en el primer pago
    if (rendido !== null) {
      out[idxs[0]].diferencia = round2(suma - rendido)
      idxs.slice(1).forEach(i => { out[i].diferencia = 0 })
    }
    // mismo monto dos veces para la misma minuta: posible pago duplicado
    for (let a = 0; a < idxs.length; a++) {
      for (let b = a + 1; b < idxs.length; b++) {
        if (Math.abs(out[idxs[a]].total - out[idxs[b]].total) < 0.05) { out[idxs[a]].duplicado = true; out[idxs[b]].duplicado = true }
      }
    }
  }

  out.forEach(r => {
    let estadoFila = 'ok', alerta = ''
    if (!r.minuta && !r.lote) {
      estadoFila = 'sin_minuta'
      alerta = r.minutaEstado === 'ambigua'
        ? `Hay varios números de 6 dígitos (${(r.minutaCandidatos || []).join(', ')}): elegí cuál es la minuta`
        : 'No encontré la minuta en "Motivo pago": cargala a mano'
    } else if (!r.lote) {
      estadoFila = 'sin_lote'
      alerta = hayCRM ? 'La minuta no figura en el reporte de Rendiciones: revisá el número o asigná el lote a mano'
                      : 'No cargaste Rendiciones: asigná el lote a mano'
    }
    if (r.duplicado) { estadoFila = 'revisar'; alerta = 'Posible pago duplicado (misma minuta y mismo monto)' }
    else if (r.diferencia !== null && Math.abs(r.diferencia) > TOLERANCIA) { estadoFila = 'revisar'; alerta = 'Diferencia mayor a $50 contra Rendiciones: revisar el lote físico' }
    r.estadoFila = estadoFila
    r.alerta = alerta
  })
  return out
}

// ── CONTROL POR LOTE ───────────────────────────────────────
// Cada linea RM del Mayor es un LOTE de rendicion. La suma de los pagos de Mov con ese lote
// debe coincidir con el Mayor (y con el total del lote en Rendiciones).
export function armarLotes(movRows, mayorRM = [], rendLotes = {}) {
  const mapa = new Map()
  const get = l => {
    if (!mapa.has(l)) mapa.set(l, { lote: l, mayor: null, mayorN: 0, crm: null, crmN: 0, mov: 0, movN: 0 })
    return mapa.get(l)
  }
  for (const r of movRows) if (r.lote) { const x = get(r.lote); x.mov += r.total; x.movN++ }
  for (const m of mayorRM) if (m.referencia) { const x = get(m.referencia); x.mayor = (x.mayor || 0) + m.monto; x.mayorN++ }
  for (const x of mapa.values()) {
    const c = rendLotes[x.lote]
    if (c) { x.crm = c.total; x.crmN = c.n }
  }
  const dif = (a, b) => (a === null || b === null) ? null : round2(a - b)
  const lotes = [...mapa.values()].map(x => {
    x.mov = round2(x.mov)
    if (x.mayor !== null) x.mayor = round2(x.mayor)
    x.difMovCrm   = x.movN ? dif(x.mov, x.crm) : null
    x.difMovMayor = x.movN ? dif(x.mov, x.mayor) : null
    x.difCrmMayor = dif(x.crm, x.mayor)
    const difs = [x.difMovCrm, x.difMovMayor, x.difCrmMayor].filter(d => d !== null)
    if (!x.movN) x.estado = 'sin_pagos'
    else if (x.mayor === null) x.estado = 'sin_mayor'
    else x.estado = difs.every(d => Math.abs(d) <= TOLERANCIA) ? 'ok' : 'revisar'
    return x
  })
  lotes.sort((a, b) => Number(b.lote) - Number(a.lote))
  return lotes
}

// ── CONCILIACION COMPLETA ──────────────────────────────────
//  1) OP del Mayor  ↔ CREDIT de Movimientos (por monto)  → columna MOVIMIENTO de DEB-CRED
//  2) Pagos de Operaciones → minuta (Motivo pago) → lote (Rendiciones CRM) → columna LOTE de Mov
//  3) Cada RM del Mayor es un lote: se controla contra los pagos de Mov y contra Rendiciones
export function conciliarTH({ mayorItems, movItems, operaciones, rendMap = null, empresa = '', fechaBuffer = 10 }) {
  const avisos = []

  // periodo que cubre el extracto de TH
  let dMin = null, dMax = null
  for (const v of [...movItems.map(m => m.fecha), ...operaciones.map(o => o.fecha)]) {
    const d = parseDate(v)
    if (!d) continue
    if (dMin === null || d < dMin) dMin = d
    if (dMax === null || d > dMax) dMax = d
  }
  let mayorPeriodo = mayorItems
  if (dMin) {
    const lo = new Date(dMin); lo.setDate(lo.getDate() - fechaBuffer)
    const hi = new Date(dMax); hi.setDate(hi.getDate() + fechaBuffer)
    mayorPeriodo = mayorItems.filter(m => { const d = parseDate(m.fecha); return d && d >= lo && d <= hi })
  }

  // 1) OP ↔ CREDIT
  const credits = movItems.filter(m => m.tipo === 'CREDIT').map(m => ({ ...m, _usado: false }))
  const mejor = (monto, fecha) => {
    const c = credits.filter(x => !x._usado && Math.abs(x.monto - monto) < 0.05)
    if (!c.length) return null
    c.sort((a, b) => daysDiff(a.fecha, fecha) - daysDiff(b.fecha, fecha))
    return c[0]
  }
  const depositos = [], pendOP = []
  for (const item of mayorPeriodo.filter(m => m.tipo === 'OP')) {
    const th = mejor(item.monto, item.fecha)
    if (th) {
      th._usado = true
      th._op = item.referencia
      depositos.push({ mayor: item, th, dias: Math.round(daysDiff(item.fecha, th.fecha)), estado: 'conciliado' })
    } else {
      const d = parseDate(item.fecha)
      const fuera = dMin && d && (d < dMin || d > dMax)
      pendOP.push({ item, razon: fuera ? 'Depósito fuera del período del extracto de TH' : 'No hay un CREDIT con el mismo monto en TH' })
    }
  }
  const creditos = credits.map(c => ({
    codRef: c.codRef, fecha: c.fecha, descripcion: c.descripcion, monto: c.monto, disponible: c.disponible,
    movimiento: c._usado ? c._op : 'OP A CONCILIAR',
  }))
  const pendCredits = credits.filter(c => !c._usado).map(c => ({
    ...c, posibleDevolucion: !/dep[oó]sito/i.test(c.descripcion),
  }))

  // 2) pagos → minuta → lote
  if (rendMap && Object.keys(rendMap).length === 0) {
    avisos.push(`El reporte de Rendiciones no trae minutas de ${empresa || 'la empresa elegida'}: ningún lote se pudo asignar solo.`)
  }
  const base = operaciones.map(o => ({ ...o, loteManual: '' }))
  const mov = recalcularMov(base, rendMap)

  // 3) lotes (cada RM del Mayor)
  const mayorRM = mayorPeriodo.filter(m => m.tipo === 'RM')
  const rendLotes = rendMap ? totalesPorLote(rendMap) : {}
  const lotes = armarLotes(mov, mayorRM, rendLotes)

  return {
    depositos, pendOP, creditos, pendCredits,
    mov, lotes, mayorRM, rendMap, rendLotes,
    mayorLineas: mayorItems,
    periodo: { desde: dMin, hasta: dMax },
    avisos,
  }
}

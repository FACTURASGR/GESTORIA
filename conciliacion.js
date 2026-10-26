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

    items.push({ tipo, descripcion: desc, referencia: ref, fecha: row[iF], monto })
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
      _usado: false,
    })
  }
  return items
}

// ── PARSER: OPERACIONES DE PAGO TH ────────────────────────

export function parseOperaciones(wb) {
  const rows = sheetToRows(wb)
  let headerIdx = 0

  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const s = rows[i].join('|').toLowerCase()
    if (s.includes('nro. operacion') || s.includes('nro de pago')) { headerIdx = i; break }
  }

  const hdr = rows[headerIdx].map(h => String(h || '').toLowerCase().trim())
  const iOp   = hdr.findIndex(h => h.includes('operacion'))
  const iPago  = hdr.findIndex(h => h.includes('nro de pago') || h === 'nro de pago')
  const iFecha = hdr.findIndex(h => h === 'fecha')
  const iTotal = hdr.findIndex(h => h === 'total')
  const iSecc  = hdr.findIndex(h => h.includes('secc') || h.includes('rpa'))
  const iTH    = hdr.findIndex(h => h.includes('tarjetahabiente'))
  const iTarj  = hdr.findIndex(h => h === 'tarjeta')
  const iEst   = hdr.findIndex(h => h === 'estado')
  const iMot   = hdr.findIndex(h => h.includes('motivo'))
  const iObs   = hdr.findIndex(h => h.includes('observaci'))

  const lookup = {}
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i]
    if (!row[iOp]) continue
    const nroPago = row[iPago] ? parseInt(row[iPago]) : null
    if (!nroPago) continue

    const motivo = String(row[iMot] || '')
    const minutaMatch = motivo.match(/Minuta\s+(\d+)/i)

    lookup[nroPago] = {
      nroOp:          row[iOp],
      nroPago,
      fecha:          row[iFecha],
      total:          parseFloat(row[iTotal]) || 0,
      seccRPA:        String(row[iSecc]  || ''),
      tarjetahabiente: String(row[iTH]   || ''),
      tarjeta:        String(row[iTarj]  || ''),
      estado:         String(row[iEst]   || ''),
      motivoPago:     motivo,
      minuta:         minutaMatch ? minutaMatch[1] : '',
      observaciones:  String(row[iObs]   || ''),
    }
  }
  return lookup
}

// ── ALGORITMO DE MATCHING ──────────────────────────────────

export function buildMatches(mayorItems, movItems, operaciones, fechaBuffer = 10) {
  // Calcular rango de fechas TH para filtrar el Mayor
  const thFechas = movItems.map(m => parseDate(m.fecha)).filter(Boolean)
  let mayorFiltrado = mayorItems

  if (thFechas.length) {
    const minFecha = new Date(Math.min(...thFechas))
    const maxFecha = new Date(Math.max(...thFechas))
    minFecha.setDate(minFecha.getDate() - fechaBuffer)
    maxFecha.setDate(maxFecha.getDate() + fechaBuffer)
    mayorFiltrado = mayorItems.filter(m => {
      const d = parseDate(m.fecha)
      return d && d >= minFecha && d <= maxFecha
    })
  }

  const credits = movItems.filter(m => m.tipo === 'CREDIT').map(m => ({ ...m, _usado: false }))
  const debits  = movItems.filter(m => m.tipo === 'DEBIT' ).map(m => ({ ...m, _usado: false }))

  function mejorMatch(pool, monto, fecha) {
    const candidatos = pool.filter(c => !c._usado && Math.abs(c.monto - monto) < 0.05)
    if (!candidatos.length) return null
    candidatos.sort((a, b) => daysDiff(a.fecha, fecha) - daysDiff(b.fecha, fecha))
    return candidatos[0]
  }

  const depositos = []   // OP matcheados
  const pagos = []       // RM matcheados
  const pendMayor = []   // items del Mayor sin match
  const pendCredits = [] // CREDITs sin match (OP a conciliar)
  const pendDebits = []  // DEBITs sin match

  // OP → CREDIT
  for (const item of mayorFiltrado.filter(m => m.tipo === 'OP')) {
    const th = mejorMatch(credits, item.monto, item.fecha)
    if (th) {
      th._usado = true
      depositos.push({
        mayor: item,
        th,
        dias: Math.round(daysDiff(item.fecha, th.fecha)),
        estado: 'conciliado',
      })
    } else {
      pendMayor.push({ item, razon: 'Sin CREDIT en TH con mismo monto' })
    }
  }

  // RM → DEBIT → Operaciones
  for (const item of mayorFiltrado.filter(m => m.tipo === 'RM')) {
    const th = mejorMatch(debits, item.monto, item.fecha)
    if (th) {
      th._usado = true
      const op = th.pagoNum ? operaciones[th.pagoNum] : null
      const diferencia = parseFloat((item.monto - (op ? op.total : th.monto)).toFixed(2))
      pagos.push({
        mayor: item,
        th,
        op,
        diferencia,
        lote: '',
        estado: Math.abs(diferencia) > 50 ? 'revisar' : 'sin_lote',
      })
    } else {
      pendMayor.push({ item, razon: 'Sin DEBIT en TH con mismo monto' })
    }
  }

  credits.filter(c => !c._usado).forEach(c => pendCredits.push(c))
  debits.filter(d => !d._usado).forEach(d => pendDebits.push(d))

  return {
    depositos,
    pagos,
    pendMayor,
    pendCredits,  // OP a conciliar
    pendDebits,
    mayorFiltrado,
  }
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

// ── MATCHING CON RENDICIONES ───────────────────────────────
// Versión extendida de buildMatches que recibe el mapa de rendiciones
// y asigna el lote automáticamente a cada pago usando el número de minuta

export function buildMatchesConLotes(mayorItems, movItems, operaciones, rendiciones = {}, fechaBuffer = 10) {
  const resultado = buildMatches(mayorItems, movItems, operaciones, fechaBuffer)

  let loteados = 0

  resultado.pagos = resultado.pagos.map(pago => {
    const minuta = pago.op?.minuta || ''
    if (!minuta) return pago

    const rendicion = rendiciones[minuta]
    if (!rendicion) return pago

    loteados++
    return {
      ...pago,
      lote:  rendicion.lote,
      estado: Math.abs(pago.diferencia) > 50 ? 'revisar' : 'ok',
      _loteAutoAsignado: true,
    }
  })

  resultado._loteados    = loteados
  resultado._sinLote     = resultado.pagos.filter(p => !p.lote).length
  resultado._conRendicion = Object.keys(rendiciones).length

  return resultado
}

// Nota: detectFileType actualizado para reconocer Rendiciones
// El archivo de rendiciones tiene header: SISTEMA;NUM_MINUTA;GESTOR;FECHA ALTA MINUTA;FECHA RENDICION;LOTE RENDICION;REAL RENDIDO
// SheetJS lo lee como una fila con esas columnas. Se detecta por 'lote rendicion' o 'num_minuta'.

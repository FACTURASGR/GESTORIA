-- ══════════════════════════════════════════════════
-- CONCILIACIÓN TH — GRUPO RANDAZZO
-- Ejecutar en Supabase → SQL Editor → New query
-- ══════════════════════════════════════════════════

-- Tabla principal: una fila por cada conciliación
CREATE TABLE conciliaciones (
  id              UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  empresa         VARCHAR(20) NOT NULL,   -- CIARA | MOVILIS | KIARA | PEARA | INV4 | SALRA
  periodo_desde   DATE NOT NULL,
  periodo_hasta   DATE NOT NULL,
  estado          VARCHAR(20) DEFAULT 'borrador',  -- borrador | completa | revisión
  -- Paths en Supabase Storage (opcionales)
  archivo_mayor   TEXT,
  archivo_mov     TEXT,
  archivo_ops     TEXT,
  -- Saldos para verificación final
  saldo_mayor     NUMERIC(18,2),
  saldo_th        NUMERIC(18,2),
  created_by      UUID REFERENCES auth.users(id),
  created_at      TIMESTAMPTZ DEFAULT NOW(),
  updated_at      TIMESTAMPTZ DEFAULT NOW()
);

-- Depósitos: OP en Mayor ↔ CREDIT en TH
CREATE TABLE depositos (
  id               UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  conciliacion_id  UUID REFERENCES conciliaciones(id) ON DELETE CASCADE,
  -- Lado Mayor (K1)
  op_ref           VARCHAR(30),        -- "OP-411979"
  fecha_mayor      DATE,
  monto_mayor      NUMERIC(18,2),
  desc_mayor       TEXT,
  -- Lado TH
  cod_ref_th       VARCHAR(100),
  fecha_th         DATE,
  monto_th         NUMERIC(18,2),
  desc_th          TEXT,
  -- Estado
  estado           VARCHAR(20) DEFAULT 'conciliado',  -- conciliado | a_conciliar
  dias_diferencia  INT DEFAULT 0,
  created_at       TIMESTAMPTZ DEFAULT NOW()
);

-- Pagos de minuta: RM en Mayor ↔ DEBIT en TH ↔ Operaciones
CREATE TABLE pagos (
  id               UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  conciliacion_id  UUID REFERENCES conciliaciones(id) ON DELETE CASCADE,
  -- Lado Mayor (K1)
  rm_ref           VARCHAR(30),        -- "RM-405025"
  th_num           VARCHAR(10),        -- "29252" (número TH del Mayor)
  fecha_mayor      DATE,
  monto_mayor      NUMERIC(18,2),
  desc_mayor       TEXT,
  -- Lado TH MOV
  cod_ref_th       VARCHAR(100),
  fecha_th         DATE,
  monto_th         NUMERIC(18,2),
  nro_pago_th      BIGINT,
  -- Datos de Operaciones de Pago
  nro_operacion    VARCHAR(50),
  minuta           VARCHAR(50),        -- número de minuta extraído del motivo pago
  secc_rpa         VARCHAR(100),
  tarjetahabiente  VARCHAR(200),
  tarjeta          VARCHAR(50),
  motivo_pago      TEXT,
  -- Conciliación con CELER (se completa manualmente)
  lote             VARCHAR(30),
  diferencia       NUMERIC(18,2) DEFAULT 0,
  estado           VARCHAR(20) DEFAULT 'sin_lote', -- sin_lote | ok | diferencia | revisar
  updated_at       TIMESTAMPTZ DEFAULT NOW(),
  created_at       TIMESTAMPTZ DEFAULT NOW()
);

-- Pendientes: items sin match de ambos lados
CREATE TABLE pendientes (
  id               UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  conciliacion_id  UUID REFERENCES conciliaciones(id) ON DELETE CASCADE,
  origen           VARCHAR(10) NOT NULL,   -- "mayor" | "th"
  tipo_movimiento  VARCHAR(10) NOT NULL,   -- "OP" | "RM" | "CREDIT" | "DEBIT"
  fecha            DATE,
  monto            NUMERIC(18,2),
  descripcion      TEXT,
  referencia       VARCHAR(100),
  motivo           VARCHAR(200),
  created_at       TIMESTAMPTZ DEFAULT NOW()
);

-- ── ÍNDICES ─────────────────────────────────────────
CREATE INDEX idx_concil_empresa  ON conciliaciones(empresa);
CREATE INDEX idx_concil_periodo  ON conciliaciones(periodo_desde, periodo_hasta);
CREATE INDEX idx_dep_concil      ON depositos(conciliacion_id);
CREATE INDEX idx_pag_concil      ON pagos(conciliacion_id);
CREATE INDEX idx_pag_lote        ON pagos(lote);
CREATE INDEX idx_pag_estado      ON pagos(estado);
CREATE INDEX idx_pend_concil     ON pendientes(conciliacion_id);

-- ── FUNCIÓN: actualizar updated_at automáticamente ──
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_conciliaciones_updated
  BEFORE UPDATE ON conciliaciones
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_pagos_updated
  BEFORE UPDATE ON pagos
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- ── VISTA: resumen para la página de historial ──────
-- Acá aprendés GROUP BY, COUNT, SUM, FILTER
CREATE VIEW v_resumen AS
SELECT
  c.id,
  c.empresa,
  c.periodo_desde,
  c.periodo_hasta,
  c.estado,
  c.created_at,
  -- Depósitos
  COUNT(DISTINCT d.id)                                        AS dep_total,
  COALESCE(SUM(d.monto_mayor), 0)                            AS dep_monto,
  COUNT(DISTINCT d.id) FILTER (WHERE d.estado = 'a_conciliar') AS dep_a_conciliar,
  -- Pagos
  COUNT(DISTINCT p.id)                                        AS pag_total,
  COALESCE(SUM(p.monto_mayor), 0)                            AS pag_monto,
  COUNT(DISTINCT p.id) FILTER (WHERE p.lote IS NOT NULL AND p.lote != '') AS pag_con_lote,
  COUNT(DISTINCT p.id) FILTER (WHERE p.estado = 'sin_lote')  AS pag_sin_lote,
  COUNT(DISTINCT p.id) FILTER (WHERE ABS(p.diferencia) > 50) AS pag_con_diff,
  COALESCE(SUM(ABS(p.diferencia)), 0)                        AS suma_diferencias,
  -- Pendientes
  COUNT(DISTINCT pe.id)                                       AS pendientes_total
FROM conciliaciones c
LEFT JOIN depositos  d  ON d.conciliacion_id  = c.id
LEFT JOIN pagos      p  ON p.conciliacion_id  = c.id
LEFT JOIN pendientes pe ON pe.conciliacion_id = c.id
GROUP BY c.id;

-- ── ROW LEVEL SECURITY ──────────────────────────────
-- Solo usuarios autenticados pueden ver y modificar sus datos
ALTER TABLE conciliaciones ENABLE ROW LEVEL SECURITY;
ALTER TABLE depositos       ENABLE ROW LEVEL SECURITY;
ALTER TABLE pagos           ENABLE ROW LEVEL SECURITY;
ALTER TABLE pendientes      ENABLE ROW LEVEL SECURITY;

-- Política: cualquier usuario autenticado puede ver y modificar
-- (podés restringir por empresa más adelante)
CREATE POLICY "auth_all" ON conciliaciones FOR ALL TO authenticated USING (true) WITH CHECK (true);
CREATE POLICY "auth_all" ON depositos       FOR ALL TO authenticated USING (true) WITH CHECK (true);
CREATE POLICY "auth_all" ON pagos           FOR ALL TO authenticated USING (true) WITH CHECK (true);
CREATE POLICY "auth_all" ON pendientes      FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- Dar acceso a la vista
GRANT SELECT ON v_resumen TO authenticated;

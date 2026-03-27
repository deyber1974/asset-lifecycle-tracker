# Asset Lifecycle Tracker

Sistema de trazabilidad de activos tecnológicos con detección automática de inconsistencias, clasificación por IA y dashboard visual en tiempo real.

**Demo en vivo:** [Ver Dashboard en Google Sheets](https://docs.google.com/spreadsheets/d/1G28H5F9kAnYEWMH1rMx1OHpDjVsYI8iIWl9DO4hXflY)

---

## Contexto del problema

La empresa gestiona laptops, handhelds, tablets e impresoras que pasan por estados: `In Stock → In Transit → In Use → Repair`. Actualmente existen:

- Activos "fantasma" en tránsito por meses sin actualización
- Estado en `assets.csv` desincronizado con `status_history`
- Ubicaciones registradas que no coinciden con el último movimiento
- Activos en uso sin responsable asignado
- Garantías vencidas con activos aún operando

**Resultado del análisis sobre el dataset**: 333 inconsistencias en 120 activos, con USD 219,865 en riesgo.

---

## Arquitectura de la solución

```
data/               ← CSVs de entrada (assets, status_history, movements)
scripts/
  diagnosis.py      ← Análisis + detección de inconsistencias (Python + pandas)
  ai_analysis.py    ← Clasificación y priorización con Claude AI
  export_to_sheets.py ← Exportación del dashboard a Google Sheets
  RefreshDashboard.gs ← Botón de refresh en Google Apps Script
reports/            ← JSONs generados (diagnosis_report, ai_analysis_report)
n8n/
  workflow.json     ← Workflow de monitoreo automático (importable en n8n)
```

### Flujo de datos

```
CSV files → diagnosis.py → diagnosis_report.json
                                    ↓
                          ai_analysis.py → ai_analysis_report.json
                                    ↓
                         export_to_sheets.py → Dashboard Google Sheets
                                    ↓
                    n8n workflow (cron diario 9am)
                                    ↓
                    ¿Hay issues CRÍTICOS? → Alerta Slack
```

---

## Por qué Google Sheets como capa de visualización

Google Sheets fue elegido intencionalmente sobre otras alternativas (Tableau, Power BI, dashboard web custom) por tres razones concretas:

| Criterio | Google Sheets | Alternativa custom |
|---|---|---|
| **Acceso** | Cualquier persona con el link, sin instalar nada | Requiere deploy, acceso a servidor |
| **Colaboración** | Comentarios, filtros, edición en tiempo real | Requiere desarrollo adicional |
| **Integración** | API nativa + Apps Script para automatización | Requiere construir desde cero |

Desde el punto de vista operativo, **el equipo que gestiona activos ya trabaja en Google Workspace**. Llevar el dashboard a una herramienta que ya usan elimina la fricción de adopción — que suele ser el principal motivo por el que las soluciones de monitoreo no se usan en la práctica.

Además, Google Sheets permite que el equipo agregue columnas, filtros o notas sin tocar código, lo cual es clave para una herramienta de gestión operativa.

---

## Requisitos

- Python 3.9+
- `pip install -r requirements.txt`
- API key de Anthropic: `export ANTHROPIC_API_KEY=sk-ant-...`
- Credenciales de Google Service Account (`credentials.json`) para export_to_sheets.py
- n8n o instancia compatible para el workflow de automatización

---

## Cómo ejecutar

### 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 2. Ejecutar diagnóstico (sin API key)

```bash
python3 scripts/diagnosis.py
# → genera reports/diagnosis_report.json
```

### 3. Ejecutar análisis con IA

```bash
export ANTHROPIC_API_KEY=sk-ant-...
python3 scripts/ai_analysis.py
# → genera reports/ai_analysis_report.json
```

### 4. Exportar dashboard a Google Sheets

```bash
python3 scripts/export_to_sheets.py
# → actualiza el dashboard en Google Sheets
```

### 5. Importar workflow en n8n

1. Abrir n8n o instancia compatible
2. Ir a **Workflows → Import from file**
3. Seleccionar `n8n/workflow.json`
4. Configurar variable `SLACK_WEBHOOK_URL` (opcional)
5. Activar el workflow

---

## Tipos de inconsistencias detectadas

| Tipo | Severidad | Descripción |
|------|-----------|-------------|
| `IN_TRANSIT_STUCK_30D` | CRITICAL | En tránsito > 30 días |
| `STATUS_HISTORY_MISMATCH` | HIGH | Estado actual ≠ último historial |
| `IN_USE_NO_ASSIGNEE` | HIGH | "In Use" sin responsable |
| `REPAIR_STUCK` | HIGH | En reparación > 30 días |
| `LOCATION_MISMATCH` | MEDIUM | Ubicación ≠ último movimiento |
| `IN_STOCK_WITH_ASSIGNEE` | MEDIUM | "In Stock" con usuario asignado |
| `WARRANTY_EXPIRED_ACTIVE` | MEDIUM | Garantía vencida, activo operando |

---

## Resultados sobre el dataset

| Métrica | Valor |
|---------|-------|
| Total activos analizados | 120 |
| Total inconsistencias | 333 |
| Issues CRITICAL | 32 |
| Issues HIGH | 123 |
| Activos en tránsito >30 días | 32 (100% de los en tránsito) |
| Valor en riesgo (CRITICAL+HIGH) | USD 219,865 |

---

## Uso de IA — Claude (Anthropic)

### Dónde se usa

**`ai_analysis.py`**: Una vez que `diagnosis.py` detecta los síntomas, se envía el reporte a **Claude claude-opus-4-6** para:

1. **Identificar causas raíz sistémicas** — no solo síntomas individuales (ej: "proceso de actualización de ubicación no automatizado")
2. **Priorizar por impacto** — considera tipo de activo, costo, días transcurridos y ubicación simultáneamente
3. **Generar recomendaciones accionables** — con esfuerzo estimado y plazo concreto
4. **Producir un executive summary** — en lenguaje natural para reportar a gerencia

### Por qué IA vs solo reglas

| Enfoque | Qué puede | Qué no puede |
|---------|-----------|--------------|
| Reglas (diagnosis.py) | Detectar síntomas exactos, calcular días, comparar campos | Interpretar contexto, priorizar combinaciones complejas, lenguaje natural |
| Claude AI | Razonar sobre causas raíz, priorizar según múltiples variables, detectar patrones entre issues | Calcular días exactos, comparar registros uno a uno |

El valor de la IA está en que **cruza la información**: sabe que un Laptop de $1,965 en HUB-CL en tránsito 843 días con garantía vencida es más urgente que un Handheld de $700 en Repair 35 días, sin necesidad de reglas de priorización hard-coded.

---

## Métricas de impacto de la solución

### KPI 1 — Reducción de inconsistencias

- **Línea base**: 333 inconsistencias / 120 activos = **2.8 issues por activo**
- **Target**: < 0.5 issues por activo (83% de reducción)
- **Cómo medirlo**: ejecutar `diagnosis.py` semanalmente y comparar `total_issues`

### KPI 2 — Tiempo de resolución de desvíos en tránsito

- **Línea base**: 100% de activos en tránsito llevan >30 días (promedio estimado: >900 días)
- **Target**: 0 activos en tránsito >30 días
- **Cómo medirlo**: campo `in_transit_over_30d` en el reporte semanal

---

## Propuestas de mejora al proceso

Más allá de la solución técnica, los datos revelan problemas de proceso:

1. **No existe cierre de tránsito automático** — cuando un activo llega a destino, nadie actualiza el estado. Solución: webhook desde el sistema de recepción → actualizar `status` vía API.
2. **Asignaciones sin validación** — se puede asignar un activo a un usuario sin cambiar su estado a "In Use". Solución: regla de negocio que fuerce el cambio de estado al asignar.
3. **Garantías no monitoreadas** — nadie revisa proactivamente las fechas de vencimiento. Solución: alerta automática 60 días antes del vencimiento (ya modelado en el workflow n8n).

---

## Variables de entorno

| Variable | Requerida | Descripción |
|----------|-----------|-------------|
| `ANTHROPIC_API_KEY` | Sí (para AI) | API key de Anthropic |
| `REFERENCE_DATE` | No | Fecha de referencia `YYYY-MM-DD` (default: hoy) |
| `SLACK_WEBHOOK_URL` | No | Webhook para alertas en n8n |

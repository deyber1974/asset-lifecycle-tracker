"""
export_to_sheets.py — Exporta los resultados del análisis a Google Sheets.
Dashboard con KPI cards, gráficos de barras + dona, y botón de refresh.
"""

import json
import time
from datetime import datetime
from pathlib import Path

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# ------------------------------------------------------------
# Configuración
# ------------------------------------------------------------

SPREADSHEET_ID   = "1G28H5F9kAnYEWMH1rMx1OHpDjVsYI8iIWl9DO4hXflY"
CREDENTIALS_FILE = Path(__file__).parent.parent / "credentials.json"
DATA_DIR         = Path(__file__).parent.parent / "data"
REPORTS_DIR      = Path(__file__).parent.parent / "reports"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Paleta de colores
C = {
    "navy":        {"red": 0.11, "green": 0.20, "blue": 0.37},
    "blue":        {"red": 0.13, "green": 0.44, "blue": 0.71},
    "light_blue":  {"red": 0.88, "green": 0.93, "blue": 0.98},
    "critical":    {"red": 0.83, "green": 0.18, "blue": 0.18},
    "critical_bg": {"red": 0.99, "green": 0.90, "blue": 0.90},
    "high":        {"red": 0.85, "green": 0.40, "blue": 0.10},
    "high_bg":     {"red": 0.99, "green": 0.94, "blue": 0.87},
    "medium":      {"red": 0.70, "green": 0.52, "blue": 0.00},
    "medium_bg":   {"red": 1.00, "green": 0.97, "blue": 0.82},
    "green":       {"red": 0.13, "green": 0.52, "blue": 0.13},
    "green_bg":    {"red": 0.88, "green": 0.97, "blue": 0.88},
    "gray_bg":     {"red": 0.95, "green": 0.95, "blue": 0.95},
    "gray_light":  {"red": 0.98, "green": 0.98, "blue": 0.98},
    "white":       {"red": 1.00, "green": 1.00, "blue": 1.00},
    "dark_text":   {"red": 0.15, "green": 0.15, "blue": 0.15},
    "muted":       {"red": 0.50, "green": 0.50, "blue": 0.55},
    "refresh_btn": {"red": 0.13, "green": 0.55, "blue": 0.13},
}

SEVERITY_BG = {
    "CRITICAL": C["critical_bg"],
    "HIGH":     C["high_bg"],
    "MEDIUM":   C["medium_bg"],
    "LOW":      C["green_bg"],
}
SEVERITY_FG = {
    "CRITICAL": C["critical"],
    "HIGH":     C["high"],
    "MEDIUM":   C["medium"],
    "LOW":      C["green"],
}


# ------------------------------------------------------------
# Conexión
# ------------------------------------------------------------

def connect() -> gspread.Spreadsheet:
    creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def get_or_create_sheet(spreadsheet, title: str) -> gspread.Worksheet:
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=1000, cols=30)


def col_letter(n: int) -> str:
    """Número de columna a letra (1→A, 26→Z, 27→AA)."""
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


# ------------------------------------------------------------
# Charts helpers
# ------------------------------------------------------------

def delete_all_charts(spreadsheet, sheet_id: int):
    """Elimina todos los gráficos existentes en una hoja antes de recrearlos."""
    meta = spreadsheet.fetch_sheet_metadata()
    charts = []
    for s in meta.get("sheets", []):
        if s["properties"]["sheetId"] == sheet_id:
            charts = s.get("charts", [])
            break
    if charts:
        reqs = [{"deleteEmbeddedObject": {"objectId": c["chartId"]}} for c in charts]
        spreadsheet.batch_update({"requests": reqs})


def write_chart_data(spreadsheet, issue_types: dict, severity: dict) -> int:
    """
    Escribe los datos fuente de los gráficos en una hoja dedicada '_ChartData'.
    Al usar una hoja separada se evita interferir con el layout del Dashboard
    y los gráficos siempre tienen datos disponibles.
    Devuelve el sheetId de '_ChartData'.
    """
    cws = get_or_create_sheet(spreadsheet, "_ChartData")
    cws.clear()

    # Bloque 1 (filas 1-8): Issues por tipo — bar chart
    sorted_types = sorted(issue_types.items(), key=lambda x: -x[1])
    bar_data = [["Tipo", "Cantidad"]]
    for t, cnt in sorted_types[:7]:
        bar_data.append([t.replace("_", " ").title(), cnt])

    # Bloque 2 (filas 10-14): Severidad — donut chart
    sev_data = [["Severidad", "Cantidad"]]
    for sev_name in ["CRITICAL", "HIGH", "MEDIUM", "LOW"]:
        cnt = severity.get(sev_name, 0)
        if cnt > 0:
            sev_data.append([sev_name, cnt])

    # Escribir ambos bloques en la hoja
    cws.update(bar_data, "A1", value_input_option="USER_ENTERED")
    cws.update(sev_data, "A10", value_input_option="USER_ENTERED")

    # Ocultar la hoja para que no se vea en las pestañas
    spreadsheet.batch_update({"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": cws.id, "hidden": True},
            "fields": "hidden",
        }
    }]})

    return cws.id


def add_charts(spreadsheet, ws: gspread.Worksheet, issue_types: dict, severity: dict):
    """
    Agrega dos gráficos al dashboard:
      1. Gráfico de barras horizontales — Issues por tipo
      2. Gráfico de dona — Distribución por severidad

    Los datos fuente se escriben en una hoja separada '_ChartData' (oculta),
    lo que garantiza que los gráficos siempre lean datos válidos sin afectar
    el layout visual del Dashboard.
    """
    sid = ws.id
    cid = write_chart_data(spreadsheet, issue_types, severity)

    sorted_types = sorted(issue_types.items(), key=lambda x: -x[1])
    n_types = min(len(sorted_types), 7)
    sev_count = sum(1 for v in severity.values() if v > 0)

    # --- Gráfico 1: Barras horizontales — Issues por tipo ---
    bar_chart = {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Issues por Tipo",
                    "titleTextFormat": {
                        "bold": True, "fontSize": 11,
                        "foregroundColor": C["navy"],
                    },
                    "basicChart": {
                        "chartType": "BAR",
                        "legendPosition": "NO_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Cantidad"},
                            {"position": "LEFT_AXIS",   "title": ""},
                        ],
                        "domains": [{
                            "domain": {
                                "sourceRange": {"sources": [{
                                    "sheetId": cid,
                                    "startRowIndex": 1, "endRowIndex": 1 + n_types,
                                    "startColumnIndex": 0, "endColumnIndex": 1,
                                }]}
                            }
                        }],
                        "series": [{
                            "series": {
                                "sourceRange": {"sources": [{
                                    "sheetId": cid,
                                    "startRowIndex": 1, "endRowIndex": 1 + n_types,
                                    "startColumnIndex": 1, "endColumnIndex": 2,
                                }]}
                            },
                            "targetAxis": "BOTTOM_AXIS",
                            "colorStyle": {"rgbColor": C["blue"]},
                        }],
                        "headerCount": 0,
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": sid, "rowIndex": 14, "columnIndex": 0},
                        "offsetXPixels": 5, "offsetYPixels": 5,
                        "widthPixels": 500, "heightPixels": 320,
                    }
                },
            }
        }
    }

    # --- Gráfico 2: Dona — Severidad ---
    donut_chart = {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Distribución por Severidad",
                    "titleTextFormat": {
                        "bold": True, "fontSize": 11,
                        "foregroundColor": C["navy"],
                    },
                    "pieChart": {
                        "legendPosition": "RIGHT_LEGEND",
                        "pieHole": 0.45,
                        "domain": {
                            "sourceRange": {"sources": [{
                                "sheetId": cid,
                                "startRowIndex": 10, "endRowIndex": 10 + sev_count,
                                "startColumnIndex": 0, "endColumnIndex": 1,
                            }]}
                        },
                        "series": {
                            "sourceRange": {"sources": [{
                                "sheetId": cid,
                                "startRowIndex": 10, "endRowIndex": 10 + sev_count,
                                "startColumnIndex": 1, "endColumnIndex": 2,
                            }]}
                        },
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": sid, "rowIndex": 14, "columnIndex": 6},
                        "offsetXPixels": 5, "offsetYPixels": 5,
                        "widthPixels": 420, "heightPixels": 320,
                    }
                },
            }
        }
    }

    spreadsheet.batch_update({"requests": [bar_chart, donut_chart]})


# ------------------------------------------------------------
# Dashboard principal
# ------------------------------------------------------------

def export_dashboard(ws: gspread.Worksheet, report: dict, spreadsheet):
    s   = report["summary"]
    ai  = report.get("ai_analysis", {})
    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    issue_types = s["issue_types"]
    severity    = s["severity_breakdown"]

    # Eliminar gráficos anteriores y unmerge todas las celdas
    delete_all_charts(spreadsheet, ws.id)
    spreadsheet.batch_update({"requests": [{
        "unmergeCells": {
            "range": {"sheetId": ws.id,
                      "startRowIndex": 0, "endRowIndex": 200,
                      "startColumnIndex": 0, "endColumnIndex": 20}
        }
    }]})

    # Top 10 CRITICAL por costo
    top10 = sorted(
        [i for i in report["issues"] if i.get("severity") == "CRITICAL"],
        key=lambda x: -x.get("cost_usd", 0)
    )[:10]

    # ----------------------------------------------------------
    # Layout: 12 columnas A-L (4 cards de 3 cols cada una)
    # Columnas N-O (índices 13-14) = datos fuente para gráficos
    # ----------------------------------------------------------
    COLS = 12  # A-L visible
    empty = [""] * COLS

    def row(*vals):
        r = list(vals)
        return r + [""] * (COLS - len(r))

    rows = []

    # Fila 1 — Título
    rows.append(row("ASSET LIFECYCLE TRACKER"))

    # Fila 2 — Subtítulo
    rows.append(row(f"Reporte de Inconsistencias y Trazabilidad de Activos  |  {now}"))

    # Fila 3 — Timestamp botón (columna K-L, índices 10-11)
    r3 = [""] * COLS
    r3[10] = f"🔄  Actualizado: {now}"
    rows.append(r3)

    # Fila 4 — espaciador
    rows.append(empty[:])

    # Helpers de contexto
    total_i = s["total_issues"] or 1
    def pct(n): return f"{round(n / total_i * 100, 1)}% del total"

    # ------- KPI ROW 1 (filas 5-7): 4 cards iguales A-C | D-F | G-I | J-L -------
    rows.append(["📦  TOTAL ACTIVOS", "", "", "⚠️  TOTAL ISSUES",  "", "",
                 "🔴  CRÍTICOS",      "", "", "💰  VALOR EN RIESGO", "", ""])
    rows.append([s["total_assets"], "", "", s["total_issues"], "", "",
                 severity["CRITICAL"], "", "",
                 f"USD {s['at_risk_value_usd']:,.0f}", "", ""])
    rows.append(["activos en el sistema", "", "", f"en {s['total_assets']} activos", "", "",
                 pct(severity["CRITICAL"]), "", "",
                 "riesgo CRITICAL + HIGH", "", ""])

    # Fila 8 — espaciador
    rows.append(empty[:])

    # ------- KPI ROW 2 (filas 9-11): 4 cards iguales -------
    rows.append(["🚚  EN TRÁNSITO >30d", "", "", "🔧  REPAIR ATASCADO", "", "",
                 "👤  SIN ASIGNADO",     "", "", "📍  UBICACIÓN INCORRECTA", "", ""])
    rows.append([s["in_transit_over_30d"], "", "",
                 issue_types.get("REPAIR_STUCK", 0), "", "",
                 issue_types.get("IN_USE_NO_ASSIGNEE", 0), "", "",
                 issue_types.get("LOCATION_MISMATCH", 0), "", ""])
    rows.append([pct(s["in_transit_over_30d"]), "", "",
                 pct(issue_types.get("REPAIR_STUCK", 0)), "", "",
                 "activos 'In Use'", "", "",
                 pct(issue_types.get("LOCATION_MISMATCH", 0)), "", ""])

    # Fila 12 — espaciador
    rows.append(empty[:])

    # ------- ENCABEZADOS GRÁFICOS (fila 13): mitad izq | mitad der -------
    r13 = [""] * COLS
    r13[0] = "📊  ISSUES POR TIPO"
    r13[6] = "🍩  DISTRIBUCIÓN POR SEVERIDAD"
    rows.append(r13)

    # Filas 14-30 — espacio reservado para los gráficos (overlay)
    for _ in range(17):
        rows.append(empty[:])

    # Fila 31 — espaciador
    rows.append(empty[:])

    # ------- RESUMEN POR SEVERIDAD (filas 32-36) -------
    rows.append(row("📋  RESUMEN POR SEVERIDAD — Distribución del total de issues detectados"))
    rows.append(["🔴  CRITICAL",  "", "", "🟠  HIGH",    "", "",
                 "🟡  MEDIUM",    "", "", "✅  LOW",     "", ""])
    rows.append([severity["CRITICAL"], "", "", severity["HIGH"], "", "",
                 severity["MEDIUM"],   "", "", severity.get("LOW", 0), "", ""])
    rows.append([pct(severity["CRITICAL"]), "", "", pct(severity["HIGH"]), "", "",
                 pct(severity["MEDIUM"]),   "", "", pct(severity.get("LOW", 0)), "", ""])

    # Fila 36 — espaciador
    rows.append(empty[:])

    # ------- TOP 10 CRÍTICOS (filas 37-49) -------
    rows.append(row("🔴  TOP 10 ACTIVOS CRÍTICOS — Mayor Costo en Riesgo"))
    rows.append(["Asset ID", "Tipo", "Problema", "Días en estado",
                 "Ubicación", "Asignado a", "Costo USD", "Fecha inicio", "", "", "", ""])

    for issue in top10:
        rows.append([
            issue.get("asset_id", ""),
            issue.get("asset_type", ""),
            issue.get("type", "").replace("_", " "),
            issue.get("days_in_state", ""),
            issue.get("location", issue.get("current_location", "")),
            issue.get("assigned_to", "") or "Sin asignar",
            f"USD {issue.get('cost_usd', 0):,.0f}",
            issue.get("since_date", issue.get("last_history_date", "")),
            "", "", "", ""
        ])

    total_cost = sum(i.get("cost_usd", 0) for i in top10)
    rows.append(["", "", "", "", "", "TOTAL TOP 10:",
                 f"USD {total_cost:,.0f}", "", "", "", "", ""])

    # Fila 45 — espaciador
    rows.append(empty[:])

    # ------- RESUMEN IA (si existe) -------
    if ai.get("executive_summary"):
        rows.append(row("🤖  ANÁLISIS IA — RESUMEN EJECUTIVO"))
        rows.append(row(ai["executive_summary"]))
        rows.append(empty[:])

    if ai.get("priority_actions"):
        rows.append(row("✅  ACCIONES PRIORITARIAS", "", "IMPACTO", "", "ESFUERZO"))
        for pa in ai["priority_actions"]:
            rows.append(row(
                f"{pa.get('rank', '')}. {pa.get('action', '')}",
                "", pa.get("expected_impact", ""), "", pa.get("effort", "").upper()
            ))
        rows.append(empty[:])

    # ----------------------------------------------------------
    # Escribir todos los datos
    # ----------------------------------------------------------
    ws.clear()
    ws.update(rows, value_input_option="USER_ENTERED")

    # ----------------------------------------------------------
    # Formateo
    # ----------------------------------------------------------
    batch_fmt = []

    def bf(range_: str, fmt_: dict):
        batch_fmt.append({"range": range_, "format": fmt_})

    # --- Título (fila 1) ---
    ws.merge_cells("A1:L1")
    bf("A1:L1", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "fontSize": 20,
                       "foregroundColor": C["white"]},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
    })

    # --- Subtítulo (fila 2) ---
    ws.merge_cells("A2:L2")
    bf("A2:L2", {
        "backgroundColor": C["blue"],
        "textFormat": {"italic": True, "fontSize": 10,
                       "foregroundColor": C["white"]},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
    })

    # --- Botón refresh (K3:L3) ---
    bf("K3:L3", {
        "backgroundColor": C["refresh_btn"],
        "textFormat": {"bold": True, "fontSize": 9,
                       "foregroundColor": C["white"]},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
    })

    # --- KPI cards helper ---
    def kpi_card(col_start: int, col_end: int, row_start: int,
                 bg: dict, num_color: dict, accent: dict = None):
        letter_s = col_letter(col_start)
        letter_e = col_letter(col_end)
        r_label = row_start
        r_value = row_start + 1
        r_sub   = row_start + 2
        ws.merge_cells(f"{letter_s}{r_label}:{letter_e}{r_label}")
        ws.merge_cells(f"{letter_s}{r_value}:{letter_e}{r_value}")
        ws.merge_cells(f"{letter_s}{r_sub}:{letter_e}{r_sub}")
        top_border = {}
        if accent:
            top_border = {"borders": {"top": {
                "style": "SOLID_MEDIUM",
                "colorStyle": {"rgbColor": accent},
            }}}
        # Label
        bf(f"{letter_s}{r_label}:{letter_e}{r_label}", {
            "backgroundColor": bg,
            "textFormat": {"bold": True, "fontSize": 9,
                           "foregroundColor": C["muted"]},
            "horizontalAlignment": "CENTER",
            **top_border,
        })
        # Número grande
        bf(f"{letter_s}{r_value}:{letter_e}{r_value}", {
            "backgroundColor": bg,
            "textFormat": {"bold": True, "fontSize": 24,
                           "foregroundColor": num_color},
            "horizontalAlignment": "CENTER",
        })
        # Sub-etiqueta
        bf(f"{letter_s}{r_sub}:{letter_e}{r_sub}", {
            "backgroundColor": bg,
            "textFormat": {"italic": True, "fontSize": 8,
                           "foregroundColor": C["muted"]},
            "horizontalAlignment": "CENTER",
        })

    # KPI Row 1
    kpi_card(1,  3,  5, C["light_blue"],  C["navy"],     accent=C["navy"])
    kpi_card(4,  6,  5, C["light_blue"],  C["navy"],     accent=C["blue"])
    kpi_card(7,  9,  5, C["critical_bg"], C["critical"], accent=C["critical"])
    kpi_card(10, 12, 5, C["high_bg"],     C["high"],     accent=C["high"])

    # KPI Row 2
    kpi_card(1,  3,  9, C["critical_bg"], C["critical"], accent=C["critical"])
    kpi_card(4,  6,  9, C["high_bg"],     C["high"],     accent=C["high"])
    kpi_card(7,  9,  9, C["medium_bg"],   C["medium"],   accent=C["medium"])
    kpi_card(10, 12, 9, C["high_bg"],     C["high"],     accent=C["high"])

    # --- Encabezados de sección gráficos (fila 13) ---
    ws.merge_cells("A13:F13")
    ws.merge_cells("G13:L13")
    bf("A13:F13", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "fontSize": 10,
                       "foregroundColor": C["white"]},
        "horizontalAlignment": "LEFT",
    })
    bf("G13:L13", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "fontSize": 10,
                       "foregroundColor": C["white"]},
        "horizontalAlignment": "LEFT",
    })

    # --- Severidad breakdown (filas 32-35) ---
    sev_header_row = 32
    ws.merge_cells(f"A{sev_header_row}:L{sev_header_row}")
    bf(f"A{sev_header_row}:L{sev_header_row}", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": C["white"]},
        "horizontalAlignment": "LEFT",
    })
    kpi_card(1,  3,  sev_header_row + 1, C["critical_bg"], C["critical"], accent=C["critical"])
    kpi_card(4,  6,  sev_header_row + 1, C["high_bg"],     C["high"],     accent=C["high"])
    kpi_card(7,  9,  sev_header_row + 1, C["medium_bg"],   C["medium"],   accent=C["medium"])
    kpi_card(10, 12, sev_header_row + 1, C["green_bg"],    C["green"],    accent=C["green"])

    # --- Top 10 header (fila 37) ---
    thin_border = {"style": "SOLID", "colorStyle": {"rgbColor": {"red": 0.82, "green": 0.82, "blue": 0.82}}}
    cell_borders = {"borders": {
        "top": thin_border, "bottom": thin_border,
        "left": thin_border, "right": thin_border,
    }}

    top10_start_row = 37
    ws.merge_cells(f"A{top10_start_row}:L{top10_start_row}")
    bf(f"A{top10_start_row}:L{top10_start_row}", {
        "backgroundColor": C["critical"],
        "textFormat": {"bold": True, "fontSize": 11,
                       "foregroundColor": C["white"]},
        "horizontalAlignment": "LEFT",
    })
    # Table header row
    bf(f"A{top10_start_row+1}:H{top10_start_row+1}", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "foregroundColor": C["white"]},
        "horizontalAlignment": "CENTER",
        **cell_borders,
    })
    # Data rows alternadas con bordes
    for i in range(len(top10)):
        r = top10_start_row + 2 + i
        row_bg = C["critical_bg"] if i % 2 == 0 else C["white"]
        bf(f"A{r}:H{r}", {"backgroundColor": row_bg, **cell_borders})
        bf(f"G{r}", {"textFormat": {"bold": True, "foregroundColor": C["critical"]}})
        bf(f"D{r}", {"horizontalAlignment": "CENTER",
                     "textFormat": {"bold": True, "foregroundColor": C["navy"]}})

    # Totals row
    totals_row = top10_start_row + 2 + len(top10)
    bf(f"A{totals_row}:H{totals_row}", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "foregroundColor": C["white"]},
        **cell_borders,
    })

    # --- AI section headers (si existen) ---
    ai_start = totals_row + 2
    if ai.get("executive_summary"):
        ws.merge_cells(f"A{ai_start}:L{ai_start}")
        bf(f"A{ai_start}:L{ai_start}", {
            "backgroundColor": C["navy"],
            "textFormat": {"bold": True, "fontSize": 10,
                           "foregroundColor": C["white"]},
        })
        ws.merge_cells(f"A{ai_start+1}:L{ai_start+1}")
        bf(f"A{ai_start+1}:L{ai_start+1}", {
            "backgroundColor": C["light_blue"],
            "textFormat": {"italic": True, "foregroundColor": C["navy"]},
        })

    # Aplicar todos los formatos de una sola vez
    ws.batch_format(batch_fmt)

    # ----------------------------------------------------------
    # Dimensiones de columnas y filas
    # ----------------------------------------------------------
    dim_reqs = [
        # Columnas A-C (1-3): 150px
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": 3},
            "properties": {"pixelSize": 150}, "fields": "pixelSize"
        }},
        # Columnas D-F: 150px
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": 3, "endIndex": 6},
            "properties": {"pixelSize": 150}, "fields": "pixelSize"
        }},
        # Columnas G-I: 150px
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": 6, "endIndex": 9},
            "properties": {"pixelSize": 150}, "fields": "pixelSize"
        }},
        # Columnas J-L: 150px
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": 9, "endIndex": 12},
            "properties": {"pixelSize": 150}, "fields": "pixelSize"
        }},
        # Fila 1 (título): 55px
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "ROWS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 55}, "fields": "pixelSize"
        }},
        # Filas 5-7 y 9-11 (KPI cards row 1 & 2): 38px
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "ROWS",
                      "startIndex": 4, "endIndex": 11},
            "properties": {"pixelSize": 38}, "fields": "pixelSize"
        }},
        # Filas de gráficos (14-30): 20px de alto base
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "ROWS",
                      "startIndex": 13, "endIndex": 30},
            "properties": {"pixelSize": 21}, "fields": "pixelSize"
        }},
        # Fila 32 header severidad: 32px
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "ROWS",
                      "startIndex": 31, "endIndex": 32},
            "properties": {"pixelSize": 32}, "fields": "pixelSize"
        }},
        # Filas 33-35 (severidad cards): 38px
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "ROWS",
                      "startIndex": 32, "endIndex": 35},
            "properties": {"pixelSize": 38}, "fields": "pixelSize"
        }},
    ]
    spreadsheet.batch_update({"requests": dim_reqs})

    # Freeze primeras 3 filas
    ws.freeze(rows=3)

    # Agregar gráficos
    add_charts(spreadsheet, ws, issue_types, severity)


# ------------------------------------------------------------
# Issues
# ------------------------------------------------------------

def export_issues(ws: gspread.Worksheet, report: dict):
    headers = ["Asset ID", "Tipo", "Problema", "Severidad", "Descripción",
               "Ubicación", "Asignado a", "Días", "Costo USD", "Fecha"]
    rows = [headers]
    for issue in report["issues"]:
        rows.append([
            issue.get("asset_id", ""),
            issue.get("asset_type", ""),
            issue.get("type", "").replace("_", " "),
            issue.get("severity", ""),
            issue.get("description", ""),
            issue.get("location", issue.get("current_location", "")),
            issue.get("assigned_to", "") or "",
            issue.get("days_in_state", ""),
            issue.get("cost_usd", ""),
            issue.get("since_date", issue.get("last_history_date",
                      issue.get("last_movement_date", ""))),
        ])

    ws.clear()
    ws.update(rows, value_input_option="USER_ENTERED")

    n = len(headers)
    ws.format(f"A1:{col_letter(n)}1", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "foregroundColor": C["white"]},
        "horizontalAlignment": "CENTER",
    })

    batch = []
    for i, issue in enumerate(report["issues"], start=2):
        sev = issue.get("severity", "LOW")
        batch.append({"range": f"A{i}:{col_letter(n)}{i}",
                      "format": {"backgroundColor": SEVERITY_BG.get(sev, C["white"])}})
        batch.append({"range": f"D{i}",
                      "format": {"backgroundColor": SEVERITY_BG.get(sev, C["white"]),
                                 "textFormat": {"bold": True,
                                                "foregroundColor": SEVERITY_FG.get(sev, C["dark_text"])},
                                 "horizontalAlignment": "CENTER"}})
    if batch:
        ws.batch_format(batch)
    ws.freeze(rows=1)

    ws.spreadsheet.batch_update({"requests": [
        {"updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": 4, "endIndex": 5},
            "properties": {"pixelSize": 360}, "fields": "pixelSize"
        }}
    ]})


# ------------------------------------------------------------
# En Tránsito
# ------------------------------------------------------------

def export_transit(ws: gspread.Worksheet, report: dict):
    transit = sorted(
        [i for i in report["issues"]
         if i["type"] in ("IN_TRANSIT_STUCK_30D", "IN_TRANSIT_STUCK_15D")],
        key=lambda x: -x.get("days_in_state", 0)
    )
    headers = ["Asset ID", "Tipo", "Severidad", "Días en Tránsito",
               "Umbral", "Ubicación", "Asignado a", "Desde", "Costo USD"]
    rows = [
        ["🚚  ACTIVOS EN TRÁNSITO — DETALLE"] + [""] * 8,
        [f"Total: {len(transit)} activos | Todos superan el límite de 30 días"] + [""] * 8,
        [""] * 9,
        headers,
    ]
    for issue in transit:
        rows.append([
            issue.get("asset_id", ""),
            issue.get("asset_type", ""),
            issue.get("severity", ""),
            issue.get("days_in_state", 0),
            "> 30 días" if issue["type"] == "IN_TRANSIT_STUCK_30D" else "> 15 días",
            issue.get("location", ""),
            issue.get("assigned_to", "") or "Sin asignar",
            issue.get("since_date", ""),
            issue.get("cost_usd", ""),
        ])
    total_cost = sum(i.get("cost_usd", 0) for i in transit)
    rows.append(["TOTAL", len(transit), "", "", "", "", "", "Valor total:",
                 f"USD {total_cost:,.0f}"])

    ws.clear()
    ws.update(rows, value_input_option="USER_ENTERED")

    ws.merge_cells("A1:I1")
    ws.format("A1:I1", {"backgroundColor": C["critical"],
                         "textFormat": {"bold": True, "fontSize": 13,
                                        "foregroundColor": C["white"]},
                         "horizontalAlignment": "CENTER"})
    ws.merge_cells("A2:I2")
    ws.format("A2:I2", {"backgroundColor": C["critical_bg"],
                         "textFormat": {"italic": True, "foregroundColor": C["critical"]},
                         "horizontalAlignment": "CENTER"})
    ws.format("A4:I4", {"backgroundColor": C["navy"],
                         "textFormat": {"bold": True, "foregroundColor": C["white"]},
                         "horizontalAlignment": "CENTER"})

    batch = []
    for i, issue in enumerate(transit, start=5):
        sev = issue.get("severity", "LOW")
        batch.append({"range": f"A{i}:I{i}",
                      "format": {"backgroundColor": SEVERITY_BG.get(sev, C["white"])}})
    if batch:
        ws.batch_format(batch)

    total_row = len(transit) + 5
    ws.format(f"A{total_row}:I{total_row}", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "foregroundColor": C["white"]}
    })
    ws.freeze(rows=4)


# ------------------------------------------------------------
# CSV crudo
# ------------------------------------------------------------

def export_raw_csv(ws: gspread.Worksheet, csv_path: Path):
    df = pd.read_csv(csv_path)
    data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    ws.clear()
    ws.update(data, value_input_option="USER_ENTERED")
    n = len(df.columns)
    ws.format(f"A1:{col_letter(n)}1", {
        "backgroundColor": C["navy"],
        "textFormat": {"bold": True, "foregroundColor": C["white"]},
        "horizontalAlignment": "CENTER",
    })
    ws.freeze(rows=1)


# ------------------------------------------------------------
# Main
# ------------------------------------------------------------

def print_apps_script_instructions():
    gs_path = Path(__file__).parent / "RefreshDashboard.gs"
    print(f"""
{'='*60}
SETUP BOTÓN REFRESH (una sola vez, ~90 segundos)
{'='*60}
1. Abrí el spreadsheet
2. Click: Extensiones → Apps Script
3. Borrá el código existente
4. Pegá el contenido de: {gs_path}
5. Guardá (Ctrl+S) y cerrá la pestaña
6. Recargá el spreadsheet → aparece menú "Asset Tracker"
7. OPCIONAL: Insertar → Dibujo → dibujá un botón
   → Click derecho → Asignar secuencia → refreshDashboard
{'='*60}""")


def main():
    print("=" * 60)
    print("ASSET LIFECYCLE TRACKER — EXPORT TO SHEETS")
    print("=" * 60)

    report_path = REPORTS_DIR / "diagnosis_report.json"
    if not report_path.exists():
        print("ERROR: Ejecutá diagnosis.py primero.")
        return

    with open(report_path, encoding="utf-8") as f:
        report = json.load(f)

    ai_path = REPORTS_DIR / "ai_analysis_report.json"
    if ai_path.exists():
        with open(ai_path, encoding="utf-8") as f:
            ai_report = json.load(f)
        # Mergear: mantener issues/summary del diagnóstico, agregar ai_analysis del reporte IA
        report["ai_analysis"] = ai_report.get("ai_analysis", {})
        print("  Usando reporte con análisis IA")
    else:
        print("  Usando reporte de diagnóstico (sin análisis IA)")

    print("\nConectando con Google Sheets...")
    spreadsheet = connect()
    print(f"  Conectado: {spreadsheet.title}")

    tabs = [
        ("Dashboard",      lambda ws: export_dashboard(ws, report, spreadsheet)),
        ("Issues",         lambda ws: export_issues(ws, report)),
        ("En Tránsito",    lambda ws: export_transit(ws, report)),
        ("assets",         lambda ws: export_raw_csv(ws, DATA_DIR / "assets.csv")),
        ("movements",      lambda ws: export_raw_csv(ws, DATA_DIR / "movements.csv")),
        ("status_history", lambda ws: export_raw_csv(ws, DATA_DIR / "status_history.csv")),
    ]

    for title, fn in tabs:
        print(f"  Exportando '{title}'...")
        ws = get_or_create_sheet(spreadsheet, title)
        fn(ws)
        time.sleep(8)  # evitar rate limit de Sheets API (300 req/min)

    print(f"\nSheet actualizado exitosamente.")
    print(f"URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")

    print_apps_script_instructions()


if __name__ == "__main__":
    main()

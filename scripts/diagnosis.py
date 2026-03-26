"""
diagnosis.py — Módulo de diagnóstico de activos tecnológicos.

QUÉ HACE:
  Carga los 3 CSVs (assets, status_history, movements), detecta todas las
  inconsistencias y desvíos, y guarda un reporte estructurado en JSON.

POR QUÉ ASÍ:
  - Separamos el diagnóstico de la clasificación IA para poder correrlos
    de forma independiente (el diagnóstico no requiere API key).
  - Usamos pandas porque manipular fechas y hacer joins sobre CSVs en pure
    Python sería mucho más verboso y propenso a errores.
  - Cada tipo de inconsistencia está en su propia función para facilitar
    el testeo y la extensión futura.

TIPOS DE INCONSISTENCIAS DETECTADAS:
  1. IN_TRANSIT_STUCK_30D  — En tránsito > 30 días  (CRITICAL/HIGH)
  2. IN_TRANSIT_STUCK_15D  — En tránsito > 15 días  (MEDIUM)
  3. REPAIR_STUCK          — En reparación > 30 días (HIGH/MEDIUM)
  4. STATUS_HISTORY_MISMATCH — Estado actual ≠ último historial (HIGH)
  5. IN_USE_NO_ASSIGNEE    — "In Use" sin usuario asignado (HIGH)
  6. IN_STOCK_WITH_ASSIGNEE — "In Stock" con usuario asignado (MEDIUM)
  7. LOCATION_MISMATCH     — Ubicación actual ≠ último movimiento (MEDIUM)
  8. WARRANTY_EXPIRED_ACTIVE — Garantía vencida y activo en uso (MEDIUM)
"""

import json
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# ------------------------------------------------------------
# Configuración
# ------------------------------------------------------------

# Fecha de referencia para calcular días transcurridos.
# Se puede sobreescribir con variable de entorno: REFERENCE_DATE=2024-09-01
import os
_ref_env = os.getenv("REFERENCE_DATE")
REFERENCE_DATE = date.fromisoformat(_ref_env) if _ref_env else date(2026, 3, 26)

DATA_DIR = Path(__file__).parent.parent / "data"
REPORTS_DIR = Path(__file__).parent.parent / "reports"


# ------------------------------------------------------------
# Carga de datos
# ------------------------------------------------------------

def load_data():
    """Carga y parsea los 3 archivos CSV."""
    assets = pd.read_csv(
        DATA_DIR / "assets.csv",
        parse_dates=["purchase_date", "last_update", "warranty_end"],
    )
    history = pd.read_csv(DATA_DIR / "status_history.csv", parse_dates=["status_date"])
    movements = pd.read_csv(DATA_DIR / "movements.csv", parse_dates=["movement_date"])
    return assets, history, movements


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------

def _days_since(d) -> int:
    """Calcula días entre una fecha y REFERENCE_DATE."""
    if isinstance(d, pd.Timestamp):
        d = d.date()
    return (REFERENCE_DATE - d).days


def _last_status_per_asset(history: pd.DataFrame) -> pd.DataFrame:
    """
    Para cada asset, devuelve el último status registrado en el historial.

    Por qué: queremos comparar el estado actual (assets.csv) contra la última
    entrada conocida en el historial para detectar divergencias.
    """
    return (
        history.sort_values("status_date")
        .groupby("asset_id", as_index=False)
        .last()[["asset_id", "status", "status_date"]]
        .rename(columns={"status": "hist_status", "status_date": "hist_date"})
    )


def _last_location_per_asset(movements: pd.DataFrame) -> pd.DataFrame:
    """
    Para cada asset, devuelve la última ubicación de destino registrada en movements.

    Por qué: el campo `location` en assets.csv debería coincidir con el último
    `to_location` de movements. Si no coincide, hay un desfase de datos.
    """
    return (
        movements.sort_values("movement_date")
        .groupby("asset_id", as_index=False)
        .last()[["asset_id", "to_location", "movement_date"]]
        .rename(columns={"to_location": "mov_location", "movement_date": "mov_date"})
    )


def _transit_entry_date(asset_id: str, history: pd.DataFrame, fallback) -> date:
    """
    Devuelve la fecha en que el asset entró en su estado 'In Transit' actual.

    Estrategia: recorre el historial de más reciente a más antiguo y busca
    el punto donde el estado dejó de ser 'In Transit'. La entrada siguiente
    es cuando empezó el tránsito actual.

    Si no hay historial, usa el campo last_update del asset (fallback).
    """
    ah = history[history["asset_id"] == asset_id].sort_values("status_date")
    rows = ah[["status", "status_date"]].values.tolist()

    for i in range(len(rows) - 1, -1, -1):
        if rows[i][0] != "In Transit":
            if i + 1 < len(rows):
                d = rows[i + 1][1]
                return d.date() if isinstance(d, pd.Timestamp) else d
            break

    # Si todo el historial es In Transit, tomamos la fecha más temprana
    in_t = ah[ah["status"] == "In Transit"]
    if not in_t.empty:
        d = in_t.iloc[0]["status_date"]
        return d.date() if isinstance(d, pd.Timestamp) else d

    return fallback.date() if isinstance(fallback, pd.Timestamp) else fallback


# ------------------------------------------------------------
# Detectores de inconsistencias
# ------------------------------------------------------------

def check_transit_stuck(assets: pd.DataFrame, history: pd.DataFrame) -> list:
    """
    Detecta activos que llevan demasiado tiempo en estado 'In Transit'.

    Regla de negocio:
      > 30 días → desvío confirmado (HIGH / CRITICAL si > 180 días)
      > 15 días → advertencia temprana (MEDIUM)

    Por qué importa: un activo en tránsito prolongado puede estar perdido,
    olvidado en un hub, o con su registro desactualizado. El costo acumulado
    por inmovilización puede ser significativo.
    """
    issues = []
    in_transit = assets[assets["status"] == "In Transit"]

    for _, row in in_transit.iterrows():
        entry = _transit_entry_date(row["asset_id"], history, row["last_update"])
        days = _days_since(entry)

        base = {
            "asset_id": row["asset_id"],
            "asset_type": row["asset_type"],
            "location": row["location"],
            "assigned_to": row["assigned_to"] if pd.notna(row["assigned_to"]) else None,
            "cost_usd": float(row["cost_usd"]),
            "days_in_state": days,
            "since_date": entry.isoformat(),
        }

        if days > 30:
            severity = "CRITICAL" if days > 180 else "HIGH"
            issues.append({
                **base,
                "type": "IN_TRANSIT_STUCK_30D",
                "severity": severity,
                "description": f"Asset en tránsito por {days} días (límite: 30 días)",
            })
        elif days > 15:
            issues.append({
                **base,
                "type": "IN_TRANSIT_STUCK_15D",
                "severity": "MEDIUM",
                "description": f"Asset en tránsito por {days} días (advertencia: >15 días)",
            })

    return issues


def check_repair_stuck(assets: pd.DataFrame, history: pd.DataFrame) -> list:
    """
    Detecta activos atascados en Repair > 30 días.

    Por qué importa: una reparación prolongada puede indicar que el activo
    fue dado de baja informalmente pero sigue en el inventario activo,
    generando ruido en los reportes y potencialmente duplicando pedidos.
    """
    issues = []
    in_repair = assets[assets["status"] == "Repair"]

    for _, row in in_repair.iterrows():
        ah = history[history["asset_id"] == row["asset_id"]].sort_values("status_date")
        repair_entries = ah[ah["status"] == "Repair"]

        if not repair_entries.empty:
            repair_date = repair_entries.iloc[-1]["status_date"]
        else:
            repair_date = row["last_update"]

        days = _days_since(repair_date)

        if days > 30:
            issues.append({
                "asset_id": row["asset_id"],
                "asset_type": row["asset_type"],
                "type": "REPAIR_STUCK",
                "severity": "HIGH" if days > 90 else "MEDIUM",
                "description": f"Asset en reparación por {days} días",
                "days_in_state": days,
                "since_date": (repair_date.date() if isinstance(repair_date, pd.Timestamp) else repair_date).isoformat(),
                "location": row["location"],
                "assigned_to": row["assigned_to"] if pd.notna(row["assigned_to"]) else None,
                "cost_usd": float(row["cost_usd"]),
            })

    return issues


def check_status_history_mismatch(assets: pd.DataFrame, history: pd.DataFrame) -> list:
    """
    Detecta activos cuyo estado en assets.csv difiere del último registro
    en status_history.

    Por qué importa: indica que el estado fue actualizado en uno de los
    sistemas pero no en el otro, lo que rompe la trazabilidad. Este es el
    síntoma más directo de procesos manuales mal coordinados.
    """
    last = _last_status_per_asset(history)
    merged = assets.merge(last, on="asset_id", how="left")

    issues = []
    mismatch = merged[
        merged["hist_status"].notna() & (merged["status"] != merged["hist_status"])
    ]

    for _, row in mismatch.iterrows():
        issues.append({
            "asset_id": row["asset_id"],
            "asset_type": row["asset_type"],
            "type": "STATUS_HISTORY_MISMATCH",
            "severity": "HIGH",
            "description": (
                f"Estado actual '{row['status']}' ≠ último historial "
                f"'{row['hist_status']}' ({row['hist_date'].date()})"
            ),
            "current_status": row["status"],
            "last_history_status": row["hist_status"],
            "last_history_date": row["hist_date"].date().isoformat(),
            "location": row["location"],
            "cost_usd": float(row["cost_usd"]),
        })

    return issues


def check_assignment_mismatch(assets: pd.DataFrame) -> list:
    """
    Detecta inconsistencias de asignación según el estado del activo:
      - 'In Use' sin usuario → no se sabe quién lo tiene
      - 'In Stock' con usuario → el activo está almacenado pero figura asignado

    Por qué importa: impacta directamente en auditorías y en la posibilidad
    de recuperar activos cuando se necesitan.
    """
    issues = []

    # In Use sin assigned_to
    for _, row in assets[
        (assets["status"] == "In Use") & assets["assigned_to"].isna()
    ].iterrows():
        issues.append({
            "asset_id": row["asset_id"],
            "asset_type": row["asset_type"],
            "type": "IN_USE_NO_ASSIGNEE",
            "severity": "HIGH",
            "description": "Asset 'In Use' sin usuario asignado — no hay trazabilidad del responsable",
            "location": row["location"],
            "cost_usd": float(row["cost_usd"]),
        })

    # In Stock con assigned_to
    for _, row in assets[
        (assets["status"] == "In Stock") & assets["assigned_to"].notna()
    ].iterrows():
        issues.append({
            "asset_id": row["asset_id"],
            "asset_type": row["asset_type"],
            "type": "IN_STOCK_WITH_ASSIGNEE",
            "severity": "MEDIUM",
            "description": f"Asset 'In Stock' tiene usuario asignado: {row['assigned_to']}",
            "location": row["location"],
            "assigned_to": row["assigned_to"],
            "cost_usd": float(row["cost_usd"]),
        })

    return issues


def check_location_mismatch(assets: pd.DataFrame, movements: pd.DataFrame) -> list:
    """
    Detecta activos cuya ubicación en assets.csv no coincide con el último
    destino registrado en movements.

    Por qué importa: significa que el activo fue movido físicamente pero no
    se actualizó su ubicación en el sistema, o viceversa. Esto genera
    'activos fantasma' durante inventarios.
    """
    last_loc = _last_location_per_asset(movements)
    merged = assets.merge(last_loc, on="asset_id", how="left")

    issues = []
    mismatch = merged[
        merged["mov_location"].notna() & (merged["location"] != merged["mov_location"])
    ]

    for _, row in mismatch.iterrows():
        issues.append({
            "asset_id": row["asset_id"],
            "asset_type": row["asset_type"],
            "type": "LOCATION_MISMATCH",
            "severity": "MEDIUM",
            "description": (
                f"Ubicación actual '{row['location']}' ≠ último movimiento "
                f"'{row['mov_location']}' ({row['mov_date'].date()})"
            ),
            "current_location": row["location"],
            "last_movement_location": row["mov_location"],
            "last_movement_date": row["mov_date"].date().isoformat(),
            "status": row["status"],
            "cost_usd": float(row["cost_usd"]),
        })

    return issues


def check_warranty_expired(assets: pd.DataFrame) -> list:
    """
    Detecta activos con garantía vencida que siguen en uso o en tránsito.

    Por qué importa: operar activos sin garantía expone a la empresa a costos
    de reparación no cubiertos. Priorizar su reemplazo reduce riesgo operacional.
    """
    ref_ts = pd.Timestamp(REFERENCE_DATE)
    expired = assets[
        (assets["warranty_end"] < ref_ts)
        & assets["status"].isin(["In Use", "In Transit"])
    ]

    issues = []
    for _, row in expired.iterrows():
        days_expired = _days_since(row["warranty_end"])
        issues.append({
            "asset_id": row["asset_id"],
            "asset_type": row["asset_type"],
            "type": "WARRANTY_EXPIRED_ACTIVE",
            "severity": "MEDIUM",
            "description": (
                f"Garantía expirada hace {days_expired} días — "
                f"activo sigue '{row['status']}'"
            ),
            "warranty_end": row["warranty_end"].date().isoformat(),
            "days_expired": days_expired,
            "status": row["status"],
            "location": row["location"],
            "cost_usd": float(row["cost_usd"]),
        })

    return issues


# ------------------------------------------------------------
# Construcción del reporte
# ------------------------------------------------------------

SEVERITY_ORDER = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}


def build_report(assets, history, movements) -> dict:
    """
    Ejecuta todos los detectores y consolida el reporte final.

    El reporte tiene dos secciones:
      - summary: métricas de alto nivel para el dashboard / n8n
      - issues: lista ordenada por severidad con el detalle de cada problema
    """
    all_issues: list = []
    all_issues += check_transit_stuck(assets, history)
    all_issues += check_repair_stuck(assets, history)
    all_issues += check_status_history_mismatch(assets, history)
    all_issues += check_assignment_mismatch(assets)
    all_issues += check_location_mismatch(assets, movements)
    all_issues += check_warranty_expired(assets)

    # Ordenar por severidad
    all_issues.sort(key=lambda x: SEVERITY_ORDER.get(x.get("severity", "LOW"), 3))

    # Contar por severidad y tipo
    sev_counts = {"CRITICAL": 0, "HIGH": 0, "MEDIUM": 0, "LOW": 0}
    type_counts: dict = {}
    for issue in all_issues:
        sev = issue.get("severity", "LOW")
        sev_counts[sev] = sev_counts.get(sev, 0) + 1
        t = issue["type"]
        type_counts[t] = type_counts.get(t, 0) + 1

    in_transit_assets = assets[assets["status"] == "In Transit"]
    stuck_30 = [i for i in all_issues if i["type"] == "IN_TRANSIT_STUCK_30D"]
    stuck_15 = [i for i in all_issues if i["type"] == "IN_TRANSIT_STUCK_15D"]

    return {
        "generated_at": datetime.now().isoformat(),
        "reference_date": REFERENCE_DATE.isoformat(),
        "summary": {
            "total_assets": len(assets),
            "total_issues": len(all_issues),
            "severity_breakdown": sev_counts,
            "in_transit_total": len(in_transit_assets),
            "in_transit_over_30d": len(stuck_30),
            "in_transit_over_15d": len(stuck_15) + len(stuck_30),
            "issue_types": type_counts,
            # Valor total de activos con problemas críticos/altos
            "at_risk_value_usd": sum(
                i.get("cost_usd", 0)
                for i in all_issues
                if i.get("severity") in ("CRITICAL", "HIGH")
            ),
        },
        "issues": all_issues,
    }


# ------------------------------------------------------------
# Entrypoint
# ------------------------------------------------------------

def main():
    print("=" * 60)
    print("ASSET LIFECYCLE TRACKER — DIAGNOSIS")
    print(f"Reference date: {REFERENCE_DATE}")
    print("=" * 60)

    print("\nLoading data...")
    assets, history, movements = load_data()
    print(f"  Assets loaded:          {len(assets)}")
    print(f"  Status history entries: {len(history)}")
    print(f"  Movements:              {len(movements)}")

    print("\nRunning checks...")
    report = build_report(assets, history, movements)

    REPORTS_DIR.mkdir(exist_ok=True)
    out = REPORTS_DIR / "diagnosis_report.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)

    s = report["summary"]
    print("\n" + "=" * 60)
    print("RESULTS")
    print("=" * 60)
    print(f"Total assets analyzed : {s['total_assets']}")
    print(f"Total issues found    : {s['total_issues']}")
    print(f"At-risk value (USD)   : ${s['at_risk_value_usd']:,.0f}")
    print(f"\nSeverity breakdown:")
    for sev, cnt in s["severity_breakdown"].items():
        bar = "█" * cnt
        print(f"  {sev:<10} {cnt:>3}  {bar}")
    print(f"\nIn Transit status:")
    print(f"  Total in transit    : {s['in_transit_total']}")
    print(f"  In transit > 15 days: {s['in_transit_over_15d']}")
    print(f"  In transit > 30 days: {s['in_transit_over_30d']}")
    print(f"\nIssue types:")
    for t, cnt in sorted(s["issue_types"].items(), key=lambda x: -x[1]):
        print(f"  {t:<35} {cnt}")
    print(f"\nReport saved to: {out}")
    return report


if __name__ == "__main__":
    main()

"""
ai_analysis.py — Clasificación y priorización de inconsistencias con Claude AI.

QUÉ HACE:
  Lee el reporte JSON generado por diagnosis.py y envía los datos a Claude
  (claude-opus-4-6) para obtener:
    1. Diagnóstico de causas raíz
    2. Priorización por impacto operacional y financiero
    3. Recomendaciones accionables
    4. Clasificación individual de cada issue CRITICAL/HIGH

POR QUÉ USAR IA AQUÍ (y no solo reglas):
  - Las reglas del script de diagnóstico detectan el "qué" (el síntoma).
    Claude interpreta el "por qué" y el "qué hacer" en lenguaje natural,
    sin necesidad de codificar heurísticas para cada combinación de contexto.
  - La priorización considera múltiples variables simultáneamente
    (tipo de activo, costo, días transcurridos, ubicación, asignación)
    de forma que una cascada de if/else no puede hacer eficientemente.
  - Claude puede detectar patrones emergentes entre issues
    (ej: "todos los activos en HUB-AR tienen location_mismatch") que
    scripts de reglas individuales no capturan.

CÓMO SE USA LA IA:
  - Se hace UNA llamada a la API con el resumen ejecutivo + los issues críticos.
  - El prompt está diseñado para obtener salida estructurada (JSON) que
    puede ser consumida directamente por n8n o un dashboard.
  - Usamos claude-opus-4-6 porque es el modelo más capaz para razonamiento
    complejo y análisis de datos no estructurados.

USO:
  export ANTHROPIC_API_KEY=sk-ant-...
  python3 scripts/ai_analysis.py
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

import anthropic

REPORTS_DIR = Path(__file__).parent.parent / "reports"
INPUT_REPORT = REPORTS_DIR / "diagnosis_report.json"
OUTPUT_REPORT = REPORTS_DIR / "ai_analysis_report.json"


def load_diagnosis() -> dict:
    if not INPUT_REPORT.exists():
        sys.exit(f"ERROR: {INPUT_REPORT} not found. Run diagnosis.py first.")
    with open(INPUT_REPORT, encoding="utf-8") as f:
        return json.load(f)


def build_prompt(report: dict) -> str:
    """
    Construye el prompt que se envía a Claude.

    Decisiones de diseño del prompt:
    - Pedimos JSON estructurado para poder parsearlo programáticamente.
    - Limitamos los issues enviados a los CRITICAL/HIGH para no exceder
      el context window ni desperdiciar tokens en issues de baja prioridad.
    - Incluimos el valor monetario en riesgo para que Claude pueda
      cuantificar el impacto en sus recomendaciones.
    """
    s = report["summary"]
    critical_high = [
        i for i in report["issues"] if i.get("severity") in ("CRITICAL", "HIGH")
    ][:25]  # top 25 para no saturar el contexto

    return f"""Eres un experto en IT Asset Management (ITAM) con foco en trazabilidad de ciclo de vida.

Analiza el siguiente reporte de diagnóstico de activos tecnológicos y devuelve un JSON con exactamente esta estructura:

{{
  "root_causes": [
    {{"id": 1, "cause": "...", "evidence": "...", "affected_assets": N}}
  ],
  "priority_actions": [
    {{
      "rank": 1,
      "action": "...",
      "rationale": "...",
      "expected_impact": "...",
      "effort": "low|medium|high"
    }}
  ],
  "issue_classifications": [
    {{
      "asset_id": "...",
      "issue_type": "...",
      "business_impact": "...",
      "recommended_action": "...",
      "urgency_days": N
    }}
  ],
  "kpis": {{
    "inconsistency_rate_pct": N,
    "at_risk_value_usd": N,
    "avg_transit_days_top10": N,
    "estimated_resolution_days": N
  }},
  "executive_summary": "..."
}}

DATOS DEL REPORTE:
- Fecha de referencia: {report['reference_date']}
- Total activos: {s['total_assets']}
- Total inconsistencias encontradas: {s['total_issues']}
- Tasa de inconsistencia: {s['total_issues'] / s['total_assets'] * 100:.1f}%
- Valor monetario en riesgo: USD {s['at_risk_value_usd']:,.0f}
- Activos en tránsito: {s['in_transit_total']} (todos >30 días)
- Activos en tránsito >30 días: {s['in_transit_over_30d']}

DESGLOSE POR SEVERIDAD:
{json.dumps(s['severity_breakdown'], indent=2)}

TIPOS DE PROBLEMAS:
{json.dumps(s['issue_types'], indent=2)}

TOP ISSUES CRÍTICOS Y ALTOS (muestra de {len(critical_high)}):
{json.dumps(critical_high, indent=2, ensure_ascii=False)}

INSTRUCCIONES:
- root_causes: identifica 3-5 causas raíz sistémicas (no síntomas individuales)
- priority_actions: 5 acciones ordenadas por impacto/esfuerzo, con plazos concretos
- issue_classifications: solo para los assets más críticos de la lista anterior
- kpis: calcula o estima los indicadores pedidos basándote en los datos
- executive_summary: máximo 3 oraciones para un CTO o gerente de IT
- Responde SOLO el JSON, sin texto adicional, sin markdown, sin explicaciones fuera del JSON"""


def call_claude(prompt: str) -> str:
    """
    Llama a la API de Claude y devuelve el texto de la respuesta.

    Por qué claude-opus-4-6: es el modelo más capaz de Anthropic para
    razonamiento complejo. Para un challenge de este tipo, el costo
    adicional vs Haiku/Sonnet está justificado por la calidad del análisis.
    """
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit("ERROR: ANTHROPIC_API_KEY environment variable not set.")

    client = anthropic.Anthropic(api_key=api_key)

    print("  Calling Claude claude-opus-4-6...")
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def parse_ai_response(raw: str) -> dict:
    """
    Parsea la respuesta JSON de Claude.
    Si Claude devuelve texto extra, intenta extraer el JSON.
    """
    raw = raw.strip()
    # Eliminar posibles bloques de código markdown
    if raw.startswith("```"):
        lines = raw.split("\n")
        raw = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Fallback: buscar primer { y último }
        start = raw.find("{")
        end = raw.rfind("}") + 1
        if start >= 0 and end > start:
            return json.loads(raw[start:end])
        raise ValueError(f"Could not parse Claude response as JSON:\n{raw[:500]}")


def main():
    print("=" * 60)
    print("ASSET LIFECYCLE TRACKER — AI ANALYSIS")
    print("=" * 60)

    print("\nLoading diagnosis report...")
    report = load_diagnosis()
    s = report["summary"]
    print(f"  Issues to analyze: {s['total_issues']}")
    print(f"  CRITICAL: {s['severity_breakdown']['CRITICAL']}")
    print(f"  HIGH:     {s['severity_breakdown']['HIGH']}")

    print("\nBuilding prompt for Claude...")
    prompt = build_prompt(report)

    print("\nSending to Claude AI...")
    raw_response = call_claude(prompt)

    print("  Parsing response...")
    ai_result = parse_ai_response(raw_response)

    # Combinar reporte original con análisis IA
    enriched = {
        **report,
        "ai_analysis": {
            "model": "claude-opus-4-6",
            "analyzed_at": datetime.now().isoformat(),
            **ai_result,
        },
    }

    REPORTS_DIR.mkdir(exist_ok=True)
    with open(OUTPUT_REPORT, "w", encoding="utf-8") as f:
        json.dump(enriched, f, indent=2, ensure_ascii=False)

    # Imprimir resumen
    ai = enriched["ai_analysis"]
    print("\n" + "=" * 60)
    print("AI ANALYSIS RESULTS")
    print("=" * 60)

    print("\nEXECUTIVE SUMMARY:")
    print(f"  {ai.get('executive_summary', 'N/A')}")

    print("\nROOT CAUSES IDENTIFIED:")
    for rc in ai.get("root_causes", []):
        print(f"  {rc['id']}. {rc['cause']}")
        print(f"     Evidence: {rc['evidence']}")

    print("\nPRIORITY ACTIONS:")
    for pa in ai.get("priority_actions", []):
        print(f"  [{pa['rank']}] {pa['action']}")
        print(f"      Impact: {pa['expected_impact']} | Effort: {pa['effort']}")

    print("\nKEY METRICS (AI-calculated):")
    for k, v in ai.get("kpis", {}).items():
        print(f"  {k}: {v}")

    print(f"\nFull enriched report saved to: {OUTPUT_REPORT}")


if __name__ == "__main__":
    main()

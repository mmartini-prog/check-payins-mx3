"""
app.py — Check Payins México (Refactored)
==========================================
Arquitectura en capas:
    rules_mx.py       → Configuración de reglas por entidad
    mapping.py        → Lógica de asignación de processors
    parsers.py        → Lectura y limpieza de archivos
    reconciliation.py → Motor de conciliación + export Excel
    app.py            → UI Streamlit (solo presentación)
"""

import streamlit as st
import pandas as pd
from datetime import datetime

from rules_mx import ENTITY_CONFIG, TREASURY_KEYWORDS, TREASURY_ACCOUNTS
from mapping import EntityMapper, parse_manual_mapping
from parsers import parse_kyriba, parse_payins
from reconciliation import reconcile, build_excel

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Check Payins MX",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .stApp { background-color: #05051a; color: #ffffff; }
    .block-container { padding-top: 1.5rem; padding-left: 2.5rem; padding-right: 2.5rem; }
    section[data-testid="stSidebar"] {
        background-color: #07071f;
        border-right: 1px solid #1a1aff;
    }
    h1,h2,h3 { color: #ffffff !important; font-weight: 800 !important; }
    p, label, span { color: #c8d4ff; }

    div[data-testid="metric-container"] {
        background: linear-gradient(135deg, #0a0a2e, #101050);
        border: 1px solid #1a1aff;
        border-radius: 16px;
        padding: 1rem 1.2rem;
        box-shadow: 0 0 18px rgba(26,26,255,.25);
    }
    div[data-testid="metric-container"] label { color: #a0b4ff !important; font-size:.85rem !important; }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        color: #ffffff !important; font-size:1.6rem !important; font-weight:800 !important;
    }

    .stButton button, .stDownloadButton button {
        background: linear-gradient(90deg, #1a1aff, #4b5cff);
        color: white; border: none; border-radius: 12px;
        font-weight: 700; padding: .7rem 1.4rem;
    }
    .stDataFrame { border: 1px solid #1a1aff; border-radius: 12px; overflow: hidden; }
    .stFileUploader {
        background-color: #0a0a2e; border: 1px dashed #4b5cff;
        border-radius: 14px; padding: .8rem;
    }

    .hero {
        background: linear-gradient(135deg, #07071f, #101050);
        border: 1px solid #1a1aff; border-radius: 22px;
        padding: 2rem; margin-bottom: 1.5rem;
        box-shadow: 0 0 30px rgba(26,26,255,.25);
    }
    .hero-title  { color:#fff; font-size:2.2rem; font-weight:900; margin-bottom:.3rem; }
    .hero-sub    { color:#a0b4ff; font-size:.95rem; }

    .debug-box {
        background: #0a0a2e; border: 1px solid #2a2a5a;
        border-radius: 10px; padding: .8rem 1rem; margin:.4rem 0;
        font-family: monospace; font-size: .82rem; color: #c8d4ff;
    }
    .warn-pill  { background:#3d2d00; color:#ffd060; border-radius:8px; padding:2px 8px; font-size:.8rem; margin:2px; display:inline-block; }
    .error-pill { background:#3d0a0a; color:#ff8080; border-radius:8px; padding:2px 8px; font-size:.8rem; margin:2px; display:inline-block; }
    .ok-pill    { background:#0a3d1a; color:#60ff90; border-radius:8px; padding:2px 8px; font-size:.8rem; margin:2px; display:inline-block; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <div class="hero-title">📊 Check Payins México</div>
    <div class="hero-sub">
        Conciliación automática de movimientos bancarios Kyriba vs estimaciones Payins.<br>
        Soporte multi-entidad: Dlocal Mexico · Demerge Mexico
    </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## ⚙️ Configuración")

    entity = st.selectbox(
        "Entidad",
        options=list(ENTITY_CONFIG.keys()),
        help="Seleccioná la entidad legal a analizar.",
    )

    tolerance = st.number_input(
        "Tolerancia sin alerta (%)",
        value=10, min_value=0, max_value=100, step=1,
        help="Diferencia % por debajo de la cual se marca como OK.",
    )

    st.markdown("---")
    st.markdown("### Columnas Payins (override)")
    st.caption("Dejá vacío para autodetección.")
    col_date      = st.text_input("Columna Fecha",      value="")
    col_amount    = st.text_input("Columna Monto",      value="")
    col_processor = st.text_input("Columna Procesador", value="")

    st.markdown("---")
    st.markdown("### Mapping manual")
    st.caption("Una línea por cuenta. Ejemplo:\nAA368=Kushki\nAA639=Banorte")
    manual_mapping_text = st.text_area("Cuenta=Processor", value="", height=120)

    st.markdown("---")
    st.markdown("### Treasury overrides")
    st.caption("Palabras clave adicionales para excluir movimientos internos.")
    extra_treasury_text = st.text_area(
        "Keywords extra (una por línea)", value="", height=80
    )


# ─────────────────────────────────────────────────────────────────────────────
# UPLOAD AREA
# ─────────────────────────────────────────────────────────────────────────────

up_col1, up_col2 = st.columns(2)

with up_col1:
    kyriba_files = st.file_uploader(
        "🏦 Archivos Kyriba / Banco",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="kyriba_upload",
    )

with up_col2:
    payins_files = st.file_uploader(
        "📋 Archivos Payins / Gross Profit",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="payins_upload",
    )


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS DE UI
# ─────────────────────────────────────────────────────────────────────────────

def _show_parse_issues(label: str, results: list):
    """Muestra warnings y errores de parsing en un expander."""
    all_warnings = [w for r in results for w in r.warnings]
    all_errors   = [e for r in results for e in r.errors]

    if not all_warnings and not all_errors:
        return

    with st.expander(f"⚠️ Mensajes de parsing — {label}", expanded=bool(all_errors)):
        for e in all_errors:
            st.markdown(f'<span class="error-pill">ERROR</span> {e}', unsafe_allow_html=True)
        for w in all_warnings:
            st.markdown(f'<span class="warn-pill">WARN</span> {w}', unsafe_allow_html=True)


def _fmt_mxn(value: float) -> str:
    return f"${value:,.0f}"


def _fmt_pct(value: float) -> str:
    return f"{value:+.1f}%"


# ─────────────────────────────────────────────────────────────────────────────
# MAIN FLOW
# ─────────────────────────────────────────────────────────────────────────────

if not kyriba_files or not payins_files:
    st.info(
        "👆 Subí los archivos Kyriba (extractos bancarios) y Payins (estimaciones) "
        "en los paneles de arriba para comenzar la conciliación."
    )
    st.stop()


# ── Preparar configuración de entidad ──────────────────────────────────────────
manual_map = parse_manual_mapping(manual_mapping_text)
mapper = EntityMapper(entity_name=entity, manual_map=manual_map)

extra_keywords = [k.strip().upper() for k in extra_treasury_text.splitlines() if k.strip()]
all_treasury_keywords = list(TREASURY_KEYWORDS) + extra_keywords


# ── Parse Kyriba ───────────────────────────────────────────────────────────────
kyriba_results = []
kyriba_dfs = []

with st.spinner("Leyendo archivos Kyriba..."):
    for file in kyriba_files:
        file_bytes = file.read()
        result = parse_kyriba(
            file_bytes          = file_bytes,
            file_name           = file.name,
            treasury_keywords   = all_treasury_keywords,
            treasury_accounts   = TREASURY_ACCOUNTS,
            manual_account_map  = manual_map,
            processor_mapper    = mapper.assign_kyriba_row,
        )
        kyriba_results.append(result)
        if result.ok:
            kyriba_dfs.append(result.df)

_show_parse_issues("Kyriba", kyriba_results)

if not kyriba_dfs:
    st.error("❌ No se pudo procesar ningún archivo Kyriba. Revisá los mensajes de error arriba.")
    st.stop()

kyriba_df = pd.concat(kyriba_dfs, ignore_index=True)
kyriba_df["Transaction date"] = pd.to_datetime(kyriba_df["Transaction date"], errors="coerce")
kyriba_df = kyriba_df[kyriba_df["Transaction date"].notna()].copy()

st.success(
    f"✅ Kyriba: **{len(kyriba_files)}** archivo(s) leídos — "
    f"**{len(kyriba_df):,}** movimientos totales · "
    f"**{kyriba_df['Processor'].notna().sum():,}** mapeados · "
    f"**{kyriba_df['Processor'].isna().sum():,}** sin mapear"
)


# ── Parse Payins ───────────────────────────────────────────────────────────────
payins_results = []
payins_dfs = []

with st.spinner("Leyendo archivos Payins..."):
    for file in payins_files:
        file_bytes = file.read()
        result = parse_payins(
            file_bytes           = file_bytes,
            file_name            = file.name,
            valid_agents         = mapper.valid_agents,
            processor_normalizer = mapper.normalize,
            col_date             = col_date,
            col_amount           = col_amount,
            col_processor        = col_processor,
        )
        payins_results.append(result)
        if result.ok:
            payins_dfs.append(result.df)

_show_parse_issues("Payins", payins_results)

if not payins_dfs:
    st.error("❌ No se pudo procesar ningún archivo Payins. Revisá los mensajes de error arriba.")
    st.stop()

payins_df = pd.concat(payins_dfs, ignore_index=True)
payins_df["Date"] = pd.to_datetime(payins_df["Date"], errors="coerce")
payins_df = payins_df[payins_df["Date"].notna()].copy()

st.success(
    f"✅ Payins: **{len(payins_files)}** archivo(s) leídos — "
    f"**{len(payins_df):,}** registros totales"
)


# ── DEBUG PANEL ────────────────────────────────────────────────────────────────

unmapped_kyriba = kyriba_df[kyriba_df["Processor"].isna()].copy()

with st.expander("🔎 Debug — Cuentas, processors y no mapeados", expanded=False):
    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("**Cuentas Kyriba detectadas**")
        accounts = kyriba_df[["Account code","Account ID","Company","Source file"]].drop_duplicates()
        st.dataframe(accounts, use_container_width=True, hide_index=True)

    with c2:
        st.markdown("**Processors Kyriba**")
        kyr_procs = kyriba_df["Processor"].value_counts().reset_index()
        kyr_procs.columns = ["Processor", "Movimientos"]
        st.dataframe(kyr_procs, use_container_width=True, hide_index=True)

    with c3:
        st.markdown("**Processors Payins**")
        pay_procs = payins_df["Processor"].value_counts().reset_index()
        pay_procs.columns = ["Processor", "Registros"]
        st.dataframe(pay_procs, use_container_width=True, hide_index=True)

    if not unmapped_kyriba.empty:
        st.markdown(f"**⚠️ Kyriba sin mapear ({len(unmapped_kyriba)} filas)**")
        st.markdown(
            "_Podés asignarlos usando el panel 'Mapping manual' en la barra lateral._"
        )
        st.dataframe(
            unmapped_kyriba[["Account code","Account ID","Company","Description","Complementary info","Credit","Day","Source file"]],
            use_container_width=True, hide_index=True
        )

    # Processors solo en un lado
    kyr_set = set(kyriba_df["Processor"].dropna().unique())
    pay_set = set(payins_df["Processor"].dropna().unique())
    only_kyr = sorted(kyr_set - pay_set)
    only_pay = sorted(pay_set - kyr_set)

    if only_kyr or only_pay:
        dd1, dd2 = st.columns(2)
        with dd1:
            st.markdown("**Solo en Kyriba** _(sin dato Payins)_")
            for p in only_kyr:
                st.markdown(f'<span class="warn-pill">{p}</span>', unsafe_allow_html=True)
        with dd2:
            st.markdown("**Solo en Payins** _(sin dato Banco)_")
            for p in only_pay:
                st.markdown(f'<span class="warn-pill">{p}</span>', unsafe_allow_html=True)


# ── FILTRO DE FECHAS ───────────────────────────────────────────────────────────

kyriba_df["_date"] = kyriba_df["Transaction date"].dt.date
payins_df["_date"] = payins_df["Date"].dt.date

kyriba_min, kyriba_max = kyriba_df["_date"].min(), kyriba_df["_date"].max()
payins_min, payins_max = payins_df["_date"].min(), payins_df["_date"].max()

st.markdown(
    f"📅 **Rango Kyriba:** {kyriba_min} → {kyriba_max} &nbsp;|&nbsp; "
    f"**Rango Payins:** {payins_min} → {payins_max}"
)

default_start = max(kyriba_min, payins_min)
default_end   = min(kyriba_max, payins_max)

if default_start > default_end:
    st.error(
        "⚠️ No hay fechas en común entre Kyriba y Payins. "
        "Verificá que los archivos correspondan al mismo período."
    )
    st.stop()

date_range = st.date_input(
    "Rango de fechas a conciliar",
    value=(default_start, default_end),
    min_value=min(kyriba_min, payins_min),
    max_value=max(kyriba_max, payins_max),
)

if len(date_range) != 2:
    st.info("Seleccioná una fecha de inicio y una de fin.")
    st.stop()

start_date, end_date = date_range

kyriba_filtered = kyriba_df[
    (kyriba_df["_date"] >= start_date) & (kyriba_df["_date"] <= end_date)
].drop(columns=["_date"])

payins_filtered = payins_df[
    (payins_df["_date"] >= start_date) & (payins_df["_date"] <= end_date)
].drop(columns=["_date"])

kyriba_df = kyriba_df.drop(columns=["_date"])
payins_df = payins_df.drop(columns=["_date"])

st.info(
    f"Analizando **{start_date}** al **{end_date}** — "
    f"Kyriba: {len(kyriba_filtered):,} movimientos · Payins: {len(payins_filtered):,} registros"
)


# ── SELECTOR DE PROCESSORS ─────────────────────────────────────────────────────

all_processors = sorted(
    set(kyriba_filtered["Processor"].dropna().unique())
    | set(payins_filtered["Processor"].dropna().unique())
)

selected_processors = st.multiselect(
    "Procesadores a analizar",
    options=all_processors,
    default=all_processors,
    help="Filtrá los processors que querés incluir en la conciliación.",
)


# ── BOTÓN ANALIZAR ─────────────────────────────────────────────────────────────

st.markdown("---")
run_btn = st.button("▶ Ejecutar conciliación", type="primary", use_container_width=True)

if not run_btn:
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# ANÁLISIS
# ─────────────────────────────────────────────────────────────────────────────

with st.spinner("Calculando conciliación..."):
    result = reconcile(
        kyriba_df           = kyriba_filtered,
        payins_df           = payins_filtered,
        tolerance_pct       = tolerance,
        selected_processors = selected_processors,
    )

detail  = result["detail"]
summary = result["summary"]
kpis    = result["kpis"]
no_match= result["no_match"]


# ── KPIs ───────────────────────────────────────────────────────────────────────

st.subheader("📈 KPIs globales")
k1, k2, k3, k4, k5, k6 = st.columns(6)

k1.metric("🏦 Banco total",      _fmt_mxn(kpis["total_banco"]))
k2.metric("📋 Payins total",     _fmt_mxn(kpis["total_payins"]))
k3.metric("📉 Diferencia",       _fmt_mxn(kpis["total_diff"]))
k4.metric("% Diferencia",        _fmt_pct(kpis["total_pct"]))
k5.metric("✅ Processors OK",    kpis["n_ok"])
k6.metric("⚠️ Processors alerta", kpis["n_warn"])


# ── SUMMARY ────────────────────────────────────────────────────────────────────

st.subheader("📊 Resumen por processor")

def _color_estado(val):
    colors = {
        "✅ OK":         "background-color: #0a3d1a; color: #60ff90",
        "🟡 Banco mayor":"background-color: #3d2d00; color: #ffd060",
        "🔴 Banco menor":"background-color: #3d0a0a; color: #ff8080",
        "⚠️ Sin dato Payins":"background-color: #1a1a2e; color: #a0a0ff",
        "⬜ Sin dato Banco":"background-color: #1a1a2e; color: #a0a0ff",
    }
    return colors.get(val, "")

styled_summary = summary.style\
    .format({"Banco": "{:,.0f}", "Payins estimados": "{:,.0f}", "Diferencia": "{:,.0f}", "Dif %": "{:+.1f}%"})\
    .applymap(_color_estado, subset=["Estado"])

st.dataframe(styled_summary, use_container_width=True, hide_index=True)


# ── DETALLE ─────────────────────────────────────────────────────────────────────

st.subheader("📅 Detalle por processor + día")

# Filtro rápido de processor en el detalle
filter_proc = st.selectbox(
    "Filtrar detalle por processor",
    options=["(Todos)"] + sorted(detail["Processor"].unique()),
    index=0,
)

detail_view = detail if filter_proc == "(Todos)" else detail[detail["Processor"] == filter_proc]

styled_detail = detail_view.style\
    .format({"Banco": "{:,.0f}", "Payins estimados": "{:,.0f}", "Diferencia": "{:,.0f}", "Dif %": "{:+.1f}%"})\
    .applymap(_color_estado, subset=["Estado"])

st.dataframe(styled_detail, use_container_width=True, hide_index=True)


# ── RAW DATA ───────────────────────────────────────────────────────────────────

with st.expander("🗄️ Ver Kyriba combinado (filtrado)", expanded=False):
    st.dataframe(
        kyriba_filtered[kyriba_filtered["Processor"].isin(selected_processors)],
        use_container_width=True, hide_index=True
    )

with st.expander("🗄️ Ver Payins combinados (filtrado)", expanded=False):
    st.dataframe(
        payins_filtered[payins_filtered["Processor"].isin(selected_processors)],
        use_container_width=True, hide_index=True
    )

with st.expander("⬛ Kyriba no mapeado", expanded=not unmapped_kyriba.empty):
    if unmapped_kyriba.empty:
        st.success("No hay movimientos sin mapear. ✅")
    else:
        st.warning(f"{len(unmapped_kyriba)} movimientos sin processor asignado.")
        st.dataframe(
            unmapped_kyriba[["Account code","Account ID","Company","Description","Complementary info","Credit","Day","Source file"]],
            use_container_width=True, hide_index=True
        )


# ── EXPORT ─────────────────────────────────────────────────────────────────────

st.markdown("---")
st.subheader("⬇️ Exportar")

with st.spinner("Generando Excel..."):
    excel_bytes = build_excel(
        summary_df  = summary,
        detail_df   = detail,
        kyriba_df   = kyriba_filtered,
        payins_df   = payins_filtered,
        unmapped_df = unmapped_kyriba,
        no_match_df = no_match,
        entity      = entity,
        date_range  = (start_date, end_date),
    )

file_name = f"conciliacion_mx_{entity.lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

st.download_button(
    label     = "⬇️ Descargar Excel de conciliación",
    data      = excel_bytes,
    file_name = file_name,
    mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width = True,
)

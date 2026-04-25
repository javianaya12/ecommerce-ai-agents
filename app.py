import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from reportlab.lib.colors import HexColor
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.holtwinters import ExponentialSmoothing

st.set_page_config(
    page_title="Imporey Internacional | Dashboard Avanzado",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# ESTILOS PREMIUM
# =========================================================
st.markdown("""
<style>
:root {
    --bg: #f4f7fb;
    --card: rgba(255,255,255,0.92);
    --card-strong: #ffffff;
    --line: #e8eef6;
    --text: #0f172a;
    --muted: #64748b;
    --brand: #0B1F3A;
    --brand-2: #163A63;
    --brand-3: #204d84;
    --success: #16a34a;
    --warning: #d97706;
    --danger: #dc2626;
    --shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
    --radius-xl: 24px;
    --radius-lg: 18px;
    --radius-md: 14px;
}

html, body, [data-testid="stAppViewContainer"] {
    background:
        radial-gradient(circle at top left, rgba(32,77,132,0.06), transparent 22%),
        radial-gradient(circle at top right, rgba(11,31,58,0.05), transparent 18%),
        linear-gradient(180deg, #f7f9fc 0%, #f2f5fa 100%);
    color: var(--text);
}

.main .block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1550px;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0b1f3a 0%, #102a4d 100%);
    border-right: 1px solid rgba(255,255,255,0.05);
}

section[data-testid="stSidebar"] * {
    color: #e8eef7 !important;
}

.sidebar-panel {
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 18px;
    padding: 14px 14px 8px 14px;
    margin-bottom: 14px;
    backdrop-filter: blur(6px);
}

.sidebar-title {
    font-size: 0.95rem;
    font-weight: 700;
    color: #ffffff;
    margin-bottom: 8px;
}

.sidebar-sub {
    font-size: 0.8rem;
    color: #c9d7ea;
    margin-bottom: 8px;
}

[data-testid="stFileUploaderDropzone"] {
    background: rgba(255,255,255,0.07) !important;
    border: 1px dashed rgba(255,255,255,0.28) !important;
    border-radius: 16px !important;
}

/* FIX ÚNICO: hacer visible el nombre del archivo cargado en el sidebar.
   La regla global del sidebar pintaba todo el texto en claro; por eso aquí usamos
   un selector más específico para forzar texto oscuro dentro de la tarjeta del archivo. */
section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] {
    background: #ffffff !important;
    border: 1px solid rgba(226,232,240,0.95) !important;
    border-radius: 14px !important;
    opacity: 1 !important;
}

section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] * {
    color: #0f172a !important;
    opacity: 1 !important;
}

section[data-testid="stSidebar"] [data-testid="stFileUploaderFileName"],
section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] span,
section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] p,
section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] div {
    color: #0f172a !important;
    font-weight: 700 !important;
    opacity: 1 !important;
}

section[data-testid="stSidebar"] [data-testid="stFileUploaderFileSize"] {
    color: #64748b !important;
    font-weight: 600 !important;
    opacity: 1 !important;
}

section[data-testid="stSidebar"] [data-testid="stFileUploaderDeleteBtn"] svg,
section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] svg {
    color: #334155 !important;
    fill: #334155 !important;
    opacity: 1 !important;
}

.hero-wrap {
    animation: fadeInUp 0.65s ease-out;
}

.hero-box {
    background:
        radial-gradient(circle at 80% 20%, rgba(255,255,255,0.14), transparent 18%),
        linear-gradient(135deg, #0B1F3A 0%, #173B63 56%, #1f568c 100%);
    color: white;
    padding: 30px 32px;
    border-radius: 26px;
    margin-bottom: 20px;
    box-shadow: 0 16px 40px rgba(11,31,58,0.20);
    position: relative;
    overflow: hidden;
}

.hero-box:after {
    content: "";
    position: absolute;
    right: -60px;
    top: -60px;
    width: 220px;
    height: 220px;
    background: radial-gradient(circle, rgba(255,255,255,0.10), transparent 60%);
    border-radius: 50%;
}

.hero-title {
    font-size: 2.2rem;
    font-weight: 800;
    margin-bottom: 4px;
    letter-spacing: -0.02em;
}

.hero-subtitle {
    font-size: 1rem;
    color: #dbe8f6;
}

.hero-badges {
    margin-top: 16px;
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
}

.hero-badge {
    background: rgba(255,255,255,0.10);
    border: 1px solid rgba(255,255,255,0.12);
    padding: 7px 12px;
    border-radius: 999px;
    font-size: 0.82rem;
    color: #e7f0fa;
}

.section-title {
    font-size: 1.08rem;
    font-weight: 800;
    color: var(--brand);
    margin-bottom: 12px;
}

.section-card {
    background: var(--card);
    border-radius: var(--radius-lg);
    padding: 18px 18px 14px 18px;
    box-shadow: var(--shadow);
    border: 1px solid var(--line);
    margin-bottom: 18px;
    backdrop-filter: blur(10px);
    animation: fadeInUp 0.5s ease-out;
}

.kpi-card {
    background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(248,250,252,0.96));
    border-radius: 18px;
    padding: 18px 18px 14px 18px;
    box-shadow: var(--shadow);
    border: 1px solid var(--line);
    min-height: 118px;
    position: relative;
    overflow: hidden;
    transition: transform 0.18s ease, box-shadow 0.18s ease;
    animation: fadeInUp 0.55s ease-out;
}

.kpi-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 16px 34px rgba(15, 23, 42, 0.12);
}

.kpi-topline {
    width: 100%;
    height: 5px;
    border-radius: 999px;
    margin-bottom: 12px;
}

.kpi-label {
    color: var(--muted);
    font-size: 0.92rem;
    margin-bottom: 8px;
    font-weight: 600;
}

.kpi-value {
    color: var(--text);
    font-size: 2rem;
    font-weight: 800;
    line-height: 1.1;
    letter-spacing: -0.02em;
}

.kpi-sub {
    color: #94a3b8;
    font-size: 0.82rem;
    margin-top: 8px;
}

.metric-mini {
    background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
    border: 1px solid var(--line);
    padding: 13px 14px;
    border-radius: 16px;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
}

.metric-mini-title {
    color: var(--muted);
    font-size: 0.82rem;
    margin-bottom: 4px;
    font-weight: 600;
}

.metric-mini-value {
    color: var(--text);
    font-size: 1.22rem;
    font-weight: 800;
}

.pill-row {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 6px;
    margin-bottom: 8px;
}

.small-chip {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    background: linear-gradient(180deg, #eef5ff 0%, #e8f1ff 100%);
    color: #184a8c;
    border: 1px solid #d9e7ff;
    padding: 7px 11px;
    border-radius: 999px;
    font-size: 0.82rem;
    font-weight: 600;
}

.insight-item, .recommend-item {
    background: #fbfdff;
    border: 1px solid #ebf0f6;
    border-radius: 14px;
    padding: 12px 14px;
    margin-bottom: 10px;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03);
}

.insight-index, .recommend-index {
    display: inline-block;
    min-width: 24px;
    height: 24px;
    text-align: center;
    line-height: 24px;
    border-radius: 999px;
    margin-right: 8px;
    font-size: 0.78rem;
    font-weight: 800;
}

.insight-index {
    background: #eaf3ff;
    color: #1953a6;
}

.recommend-index {
    background: #ecfdf3;
    color: #15803d;
}

.badge-good {
    color: #15803d;
    font-weight: 700;
}

.badge-warn {
    color: #b45309;
    font-weight: 700;
}

.badge-bad {
    color: #b91c1c;
    font-weight: 700;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
    margin-bottom: 6px;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 14px !important;
    padding: 10px 16px !important;
    background: rgba(255,255,255,0.82) !important;
    border: 1px solid var(--line) !important;
    color: var(--brand) !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03);
}

.stTabs [aria-selected="true"] {
    background: linear-gradient(180deg, #ffffff 0%, #eef4ff 100%) !important;
    border: 1px solid #d6e4ff !important;
}

div[data-testid="stDataFrame"] {
    border-radius: 14px !important;
    overflow: hidden;
    border: 1px solid var(--line);
}

button[kind="primary"] {
    border-radius: 12px !important;
}

.stDownloadButton > button {
    border-radius: 14px !important;
    min-height: 44px !important;
    font-weight: 700 !important;
    border: none !important;
    background: linear-gradient(135deg, #0B1F3A 0%, #173B63 100%) !important;
    color: white !important;
    box-shadow: 0 10px 24px rgba(11,31,58,0.20);
    transition: transform 0.18s ease, box-shadow 0.18s ease;
}

.stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 16px 34px rgba(11,31,58,0.26);
}

.upload-ok {
    background: linear-gradient(180deg, #ecfdf3 0%, #f5fff8 100%);
    border: 1px solid #bbf7d0;
    color: #166534;
    padding: 12px 14px;
    border-radius: 14px;
    font-weight: 600;
    margin-bottom: 14px;
    animation: fadeInUp 0.4s ease-out;
}

.upload-info {
    background: linear-gradient(180deg, #eff6ff 0%, #f8fbff 100%);
    border: 1px solid #bfdbfe;
    color: #1d4ed8;
    padding: 12px 14px;
    border-radius: 14px;
    font-weight: 600;
    margin-bottom: 12px;
}

.footer-note {
    color: #94a3b8;
    font-size: 0.85rem;
    text-align: center;
    margin-top: 10px;
}

@keyframes fadeInUp {
    from {
        opacity: 0;
        transform: translateY(8px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# HELPERS UI
# =========================================================
def get_kpi_tone(title, value_raw=None):
    title_lower = title.lower()

    if "% cancelado" in title_lower:
        if value_raw is None:
            return "#dc2626"
        if value_raw <= 2:
            return "#16a34a"
        if value_raw <= 5:
            return "#d97706"
        return "#dc2626"

    if "% entregado" in title_lower or "% pickeado" in title_lower or "sla <24" in title_lower:
        if value_raw is None:
            return "#16a34a"
        if value_raw >= 95:
            return "#16a34a"
        if value_raw >= 85:
            return "#d97706"
        return "#dc2626"

    if "tiempo ciclo" in title_lower:
        if value_raw is None or pd.isna(value_raw):
            return "#0B1F3A"
        if value_raw <= 24:
            return "#16a34a"
        if value_raw <= 48:
            return "#d97706"
        return "#dc2626"

    return "#0B1F3A"


def render_kpi_card(title, value, subtitle="", tone="#0B1F3A"):
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-topline" style="background:{tone};"></div>
        <div class="kpi-label">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)


def render_mini_metric(title, value):
    st.markdown(f"""
    <div class="metric-mini">
        <div class="metric-mini-title">{title}</div>
        <div class="metric-mini-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)


def render_section_open(title=None):
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    if title:
        st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


def render_section_close():
    st.markdown("</div>", unsafe_allow_html=True)


def render_upload_success(text):
    st.markdown(f'<div class="upload-ok">✅ {text}</div>', unsafe_allow_html=True)


def render_upload_info(text):
    st.markdown(f'<div class="upload-info">ℹ️ {text}</div>', unsafe_allow_html=True)


def render_insight_list(items):
    for i, item in enumerate(items, start=1):
        st.markdown(
            f'<div class="insight-item"><span class="insight-index">{i}</span>{item}</div>',
            unsafe_allow_html=True
        )


def render_recommend_list(items):
    for i, item in enumerate(items, start=1):
        st.markdown(
            f'<div class="recommend-item"><span class="recommend-index">{i}</span>{item}</div>',
            unsafe_allow_html=True
        )


def set_plotly_theme(fig):
    fig.update_layout(
        template="plotly_white",
        margin=dict(l=10, r=10, t=18, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#0f172a"),
        legend=dict(
            bgcolor="rgba(255,255,255,0.75)",
            bordercolor="#e8eef6",
            borderwidth=1
        )
    )
    fig.update_xaxes(
        showgrid=False,
        linecolor="#d8e1ec",
        tickfont=dict(color="#475569")
    )
    fig.update_yaxes(
        gridcolor="#edf2f7",
        zeroline=False,
        linecolor="#d8e1ec",
        tickfont=dict(color="#475569")
    )
    return fig

# =========================================================
# HEADER
# =========================================================
st.markdown("""
<div class="hero-wrap">
    <div class="hero-box">
        <div class="hero-title">Imporey Internacional</div>
        <div class="hero-subtitle">Dashboard Avanzado de Operación, Fulfillment y Logística</div>
        <div class="hero-badges">
            <span class="hero-badge">Forecast 30 días</span>
            <span class="hero-badge">SLA operativo</span>
            <span class="hero-badge">KPIs ejecutivos</span>
            <span class="hero-badge">Dashboard premium</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# MAPEO DE COLUMNAS
# =========================================================
COLUMN_MAP = {
    "Almacén": "warehouse",
    "UNE": "business_unit",
    "Canal": "channel",
    "Transportista": "carrier",
    "Status": "status",
    "Cantidad": "qty",
    "Created on": "created_on",
    "Pedido": "order_id",
    "ID de Envío": "shipment_id",
    "Producto/Código de barras": "barcode",
    "Producto": "product",
    "Finalización Pick": "pick_finished",
    "Finalización Pack": "pack_finished",
    "Fecha de finalización": "delivery_finished",
    "Pick": "pick_doc",
    "Pack": "pack_doc",
    "Entrega": "delivery_doc",
    "KPI Tiempo": "kpi_time",
    "Fecha de inicio": "start_time",
    "Entregado": "delivered_flag",
    "Pickeado": "picked_flag",
    "Tiempo dedicado": "time_spent",
}

# =========================================================
# HELPERS DE NEGOCIO - SIN CAMBIOS
# =========================================================
def parse_duration_to_hours(value):
    if pd.isna(value):
        return np.nan
    value = str(value).strip()
    if value == "" or value.lower() == "nan":
        return np.nan
    try:
        parts = value.split(":")
        if len(parts) == 3:
            h = float(parts[0])
            m = float(parts[1])
            s = float(parts[2])
            return h + (m / 60.0) + (s / 3600.0)
        return np.nan
    except Exception:
        return np.nan


def classify_sla_bucket(hours):
    if pd.isna(hours):
        return "Sin dato"
    if hours < 24:
        return "<24 h"
    if hours <= 48:
        return "24-48 h"
    return ">48 h"


def fmt_int(x):
    try:
        return f"{int(round(x)):,}"
    except Exception:
        return "0"


def fmt_pct(x):
    try:
        return f"{x:.1f}%"
    except Exception:
        return "0.0%"


def fmt_hours(x):
    try:
        if pd.isna(x):
            return "-"
        return f"{x:.1f} h"
    except Exception:
        return "-"


def safe_nunique(series):
    return series.dropna().nunique()


def normalize_score_weights(raw_weights):
    # Mejora 7: pesos del performance score ajustables y normalizados.
    total = sum(raw_weights)
    if total <= 0:
        return (0.40, 0.30, 0.20, 0.10)
    return tuple(w / total for w in raw_weights)


def validate_missing_critical_columns(df):
    # Mejora 4: validación explícita de columnas críticas reconocidas.
    critical_cols = ["order_id", "shipment_id", "created_on", "status", "warehouse"]
    missing = []
    for col in critical_cols:
        if col not in df.columns or df[col].isna().all():
            missing.append(col)
    return missing


def common_column_config():
    # Mejora 3: formato visual sin convertir números a texto; conserva ordenamiento real.
    return {
        "performance_score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%.1f"),
        "orders": st.column_config.NumberColumn("Pedidos", format="%d"),
        "shipments": st.column_config.NumberColumn("Envíos", format="%d"),
        "units": st.column_config.NumberColumn("Unidades", format="%d"),
        "records": st.column_config.NumberColumn("Registros", format="%d"),
        "delivered_rate": st.column_config.NumberColumn("% Entregado", format="%.1f%%"),
        "cancelled_rate": st.column_config.NumberColumn("% Cancelado", format="%.1f%%"),
        "share_orders": st.column_config.NumberColumn("Participación pedidos", format="%.1f%%"),
        "share_shipments": st.column_config.NumberColumn("Participación envíos", format="%.1f%%"),
        "share_units": st.column_config.NumberColumn("Participación unidades", format="%.1f%%"),
        "share": st.column_config.NumberColumn("Participación", format="%.1f%%"),
        "avg_cycle_hours": st.column_config.NumberColumn("Ciclo promedio", format="%.1f h"),
        "avg_time_spent": st.column_config.NumberColumn("Tiempo dedicado", format="%.1f h"),
        "forecast_orders": st.column_config.NumberColumn("Pedidos pronosticados", format="%d"),
        "lower_bound": st.column_config.NumberColumn("Escenario bajo", format="%d"),
        "upper_bound": st.column_config.NumberColumn("Escenario alto", format="%d"),
    }


# Mejora 1: cache para evitar reprocesar la normalización en cada interacción.
@st.cache_data(show_spinner=False)
def normalize_dataframe(df):
    df = df.copy()
    df = df.rename(columns=COLUMN_MAP)

    for col in COLUMN_MAP.values():
        if col not in df.columns:
            df[col] = np.nan

    date_cols = ["created_on", "pick_finished", "pack_finished", "delivery_finished", "start_time"]
    for c in date_cols:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    numeric_cols = ["qty", "delivered_flag", "picked_flag"]
    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    text_cols = ["warehouse", "business_unit", "channel", "carrier", "status", "product", "kpi_time"]
    for c in text_cols:
        df[c] = df[c].astype(str).replace("nan", "").fillna("").str.strip()

    df["qty"] = df["qty"].fillna(0)
    df["is_delivered"] = np.where((df["delivered_flag"] > 0) | (df["status"].str.lower() == "entregado"), 1, 0)
    df["is_picked"] = np.where(df["picked_flag"] > 0, 1, 0)
    df["is_cancelled"] = np.where(df["status"].str.lower().str.contains("cancel", na=False), 1, 0)

    df["analysis_date"] = pd.to_datetime(df["created_on"]).dt.date
    df["analysis_datetime"] = pd.to_datetime(df["created_on"])

    df["time_spent_hours"] = df["time_spent"].apply(parse_duration_to_hours)
    df["cycle_time_hours"] = (df["delivery_finished"] - df["created_on"]).dt.total_seconds() / 3600
    df["pick_time_hours"] = (df["pick_finished"] - df["created_on"]).dt.total_seconds() / 3600
    df["pack_time_hours"] = (df["pack_finished"] - df["created_on"]).dt.total_seconds() / 3600

    df["cycle_time_hours"] = np.where(df["cycle_time_hours"] < 0, np.nan, df["cycle_time_hours"])
    df["pick_time_hours"] = np.where(df["pick_time_hours"] < 0, np.nan, df["pick_time_hours"])
    df["pack_time_hours"] = np.where(df["pack_time_hours"] < 0, np.nan, df["pack_time_hours"])

    df["sla_bucket"] = df["cycle_time_hours"].apply(classify_sla_bucket)

    return df


def get_status_light(score):
    if pd.isna(score):
        return "⚪ Sin dato"
    if score >= 85:
        return "🟢 Alto"
    if score >= 70:
        return "🟡 Medio"
    return "🔴 Riesgo"


def calc_score(df_base, volume_col, score_weights):
    df = df_base.copy()

    w_delivery, w_cycle, w_cancel, w_volume = score_weights

    max_volume = df[volume_col].max() if len(df) else 0
    df["volume_score"] = (df[volume_col] / max_volume) * 100 if max_volume > 0 else 0
    df["delivery_score"] = df["delivered_rate"].clip(lower=0, upper=100)
    df["cancel_score"] = (100 - df["cancelled_rate"]).clip(lower=0, upper=100)

    valid_cycle = df["avg_cycle_hours"].replace([np.inf, -np.inf], np.nan)
    max_cycle = valid_cycle.max()

    if pd.notna(max_cycle) and max_cycle > 0:
        df["cycle_score"] = (100 - ((valid_cycle / max_cycle) * 100)).clip(lower=0, upper=100)
    else:
        df["cycle_score"] = 100

    # Mejora 7: performance score usa pesos ajustables desde el sidebar.
    df["performance_score"] = (
        df["delivery_score"] * w_delivery +
        df["cycle_score"] * w_cycle +
        df["cancel_score"] * w_cancel +
        df["volume_score"] * w_volume
    ).round(1)

    df["status_light"] = df["performance_score"].apply(get_status_light)
    return df


# Mejora 1: cache para agregaciones diarias pesadas.
@st.cache_data(show_spinner=False)
def build_daily_operations(df):
    daily = (
        df.groupby("analysis_date", dropna=True)
        .agg(
            orders=("order_id", pd.Series.nunique),
            shipments=("shipment_id", pd.Series.nunique),
            units=("qty", "sum"),
            delivered=("is_delivered", "sum"),
            cancelled=("is_cancelled", "sum"),
            picked=("is_picked", "sum")
        )
        .reset_index()
        .sort_values("analysis_date")
    )
    daily["analysis_date"] = pd.to_datetime(daily["analysis_date"])
    return daily


# Mejora 1: cache para forecast; Mejora 5: métricas MAE/MAPE de calidad del forecast.
@st.cache_data(show_spinner=False)
def build_forecast_next_period(daily_ops, forecast_days=30):
    if daily_ops.empty or len(daily_ops) < 14:
        return pd.DataFrame(), "insuficiente_historial", {"mae_7d": np.nan, "mape_7d": np.nan}

    ts = daily_ops[["analysis_date", "orders"]].copy()
    ts = ts.dropna().sort_values("analysis_date")
    ts = ts.set_index("analysis_date").asfreq("D")
    ts["orders"] = ts["orders"].fillna(0)

    train = ts["orders"].astype(float)
    seasonal_periods = 7 if len(train) >= 28 else None

    def calculate_quality(real_values, fitted_values):
        comp = pd.DataFrame({"real": real_values, "fitted": fitted_values}).dropna().tail(7)
        if comp.empty:
            return {"mae_7d": np.nan, "mape_7d": np.nan}
        mae = (comp["real"] - comp["fitted"]).abs().mean()
        non_zero = comp[comp["real"] != 0].copy()
        if non_zero.empty:
            mape = np.nan
        else:
            mape = ((non_zero["real"] - non_zero["fitted"]).abs() / non_zero["real"]).mean() * 100
        return {"mae_7d": float(mae), "mape_7d": float(mape) if pd.notna(mape) else np.nan}

    try:
        if seasonal_periods:
            model = ExponentialSmoothing(
                train,
                trend="add",
                seasonal="add",
                seasonal_periods=seasonal_periods,
                initialization_method="estimated"
            ).fit(optimized=True)
            model_name = "Holt-Winters (tendencia + estacionalidad semanal)"
        else:
            model = ExponentialSmoothing(
                train,
                trend="add",
                seasonal=None,
                initialization_method="estimated"
            ).fit(optimized=True)
            model_name = "Exponential Smoothing con tendencia"
    except Exception:
        x = np.arange(len(train))
        y = train.values
        try:
            slope, intercept = np.polyfit(x, y, 1)
            future_x = np.arange(len(train), len(train) + forecast_days)
            future_y = slope * future_x + intercept
            future_y = np.where(future_y < 0, 0, future_y)

            future_dates = pd.date_range(train.index.max() + pd.Timedelta(days=1), periods=forecast_days, freq="D")
            forecast_df = pd.DataFrame({
                "analysis_date": future_dates,
                "forecast_orders": future_y
            })

            fitted_linear = pd.Series(slope * x + intercept, index=train.index)
            forecast_quality = calculate_quality(train, fitted_linear)
            resid_std = np.std(y - fitted_linear.values) if len(y) > 2 else 0
            forecast_df["lower_bound"] = np.maximum(forecast_df["forecast_orders"] - 1.28 * resid_std, 0)
            forecast_df["upper_bound"] = forecast_df["forecast_orders"] + 1.28 * resid_std

            return forecast_df, "Regresión lineal de respaldo", forecast_quality
        except Exception:
            return pd.DataFrame(), "error", {"mae_7d": np.nan, "mape_7d": np.nan}

    forecast_values = model.forecast(forecast_days)
    fitted = model.fittedvalues.reindex(train.index)
    residuals = train - fitted
    resid_std = residuals.std() if len(residuals.dropna()) > 3 else 0
    forecast_quality = calculate_quality(train, fitted)

    forecast_df = pd.DataFrame({
        "analysis_date": forecast_values.index,
        "forecast_orders": forecast_values.values
    })

    forecast_df["forecast_orders"] = np.where(forecast_df["forecast_orders"] < 0, 0, forecast_df["forecast_orders"])
    forecast_df["lower_bound"] = np.maximum(forecast_df["forecast_orders"] - 1.28 * resid_std, 0)
    forecast_df["upper_bound"] = forecast_df["forecast_orders"] + 1.28 * resid_std

    return forecast_df, model_name, forecast_quality


# Mejora 1: cache para performance por almacén.
@st.cache_data(show_spinner=False)
def build_warehouse_performance(df, score_weights):
    base = (
        df.groupby("warehouse", dropna=True)
        .agg(
            orders=("order_id", pd.Series.nunique),
            shipments=("shipment_id", pd.Series.nunique),
            units=("qty", "sum"),
            delivered=("is_delivered", "mean"),
            cancelled=("is_cancelled", "mean"),
            avg_cycle_hours=("cycle_time_hours", "mean"),
            avg_time_spent=("time_spent_hours", "mean")
        )
        .reset_index()
    )
    total_orders = base["orders"].sum() if len(base) else 0
    base["delivered_rate"] = base["delivered"] * 100
    base["cancelled_rate"] = base["cancelled"] * 100
    base["share_orders"] = np.where(total_orders > 0, base["orders"] / total_orders * 100, 0)
    base = base.drop(columns=["delivered", "cancelled"])
    base = calc_score(base, "orders", score_weights)
    return base.sort_values("orders", ascending=False)


# Mejora 1: cache para performance por carrier.
@st.cache_data(show_spinner=False)
def build_carrier_performance(df, score_weights):
    base = (
        df.groupby("carrier", dropna=True)
        .agg(
            shipments=("shipment_id", pd.Series.nunique),
            orders=("order_id", pd.Series.nunique),
            units=("qty", "sum"),
            delivered=("is_delivered", "mean"),
            cancelled=("is_cancelled", "mean"),
            avg_cycle_hours=("cycle_time_hours", "mean"),
            avg_time_spent=("time_spent_hours", "mean")
        )
        .reset_index()
    )
    total_shipments = base["shipments"].sum() if len(base) else 0
    base["delivered_rate"] = base["delivered"] * 100
    base["cancelled_rate"] = base["cancelled"] * 100
    base["share_shipments"] = np.where(total_shipments > 0, base["shipments"] / total_shipments * 100, 0)
    base = base.drop(columns=["delivered", "cancelled"])
    base = calc_score(base, "shipments", score_weights)
    return base.sort_values("shipments", ascending=False)


# Mejora 1: cache para performance por canal.
@st.cache_data(show_spinner=False)
def build_channel_performance(df):
    base = (
        df.groupby("channel", dropna=True)
        .agg(
            orders=("order_id", pd.Series.nunique),
            shipments=("shipment_id", pd.Series.nunique),
            units=("qty", "sum"),
            delivered=("is_delivered", "mean"),
            cancelled=("is_cancelled", "mean"),
            avg_cycle_hours=("cycle_time_hours", "mean")
        )
        .reset_index()
    )
    total_orders = base["orders"].sum() if len(base) else 0
    base["delivered_rate"] = base["delivered"] * 100
    base["cancelled_rate"] = base["cancelled"] * 100
    base["share_orders"] = np.where(total_orders > 0, base["orders"] / total_orders * 100, 0)
    base = base.drop(columns=["delivered", "cancelled"])
    return base.sort_values("orders", ascending=False)


# Mejora 1: cache para performance por producto.
@st.cache_data(show_spinner=False)
def build_product_performance(df):
    base = (
        df.groupby("product", dropna=True)
        .agg(
            orders=("order_id", pd.Series.nunique),
            shipments=("shipment_id", pd.Series.nunique),
            units=("qty", "sum"),
            delivered=("is_delivered", "mean"),
            cancelled=("is_cancelled", "mean")
        )
        .reset_index()
    )
    total_units = base["units"].sum() if len(base) else 0
    base["delivered_rate"] = base["delivered"] * 100
    base["cancelled_rate"] = base["cancelled"] * 100
    base["share_units"] = np.where(total_units > 0, base["units"] / total_units * 100, 0)
    base = base.drop(columns=["delivered", "cancelled"])
    return base.sort_values("units", ascending=False)


# Mejora 1: cache para resumen SLA.
@st.cache_data(show_spinner=False)
def build_sla_summary(df):
    tmp = df[df["sla_bucket"] != "Sin dato"].copy()
    if tmp.empty:
        return pd.DataFrame(columns=["sla_bucket", "records", "share"])
    base = tmp.groupby("sla_bucket").size().reset_index(name="records")
    total = base["records"].sum()
    base["share"] = np.where(total > 0, base["records"] / total * 100, 0)
    order_map = {"<24 h": 1, "24-48 h": 2, ">48 h": 3}
    base["order"] = base["sla_bucket"].map(order_map)
    return base.sort_values("order").drop(columns=["order"])


def generate_insights(df, warehouse_perf, carrier_perf, channel_perf, product_perf, forecast_df):
    insights = []

    if not warehouse_perf.empty:
        top_wh = warehouse_perf.iloc[0]
        risk_wh = warehouse_perf.sort_values("performance_score").iloc[0]
        insights.append(f"El almacén con mayor volumen es {top_wh['warehouse']} con {fmt_int(top_wh['orders'])} pedidos.")
        insights.append(f"El almacén con mayor riesgo operativo es {risk_wh['warehouse']} con score {risk_wh['performance_score']:.1f}.")

    if not carrier_perf.empty:
        top_car = carrier_perf.iloc[0]
        risk_car = carrier_perf.sort_values("performance_score").iloc[0]
        insights.append(f"El carrier principal es {top_car['carrier']} con {fmt_int(top_car['shipments'])} envíos.")
        insights.append(f"El carrier con mayor foco de atención es {risk_car['carrier']} con score {risk_car['performance_score']:.1f}.")

    if not channel_perf.empty:
        top_ch = channel_perf.iloc[0]
        insights.append(f"El canal con más pedidos es {top_ch['channel']} con {fmt_int(top_ch['orders'])} pedidos.")

    if not product_perf.empty:
        top_prod = product_perf.iloc[0]
        insights.append(f"El producto con mayor movimiento es {top_prod['product']} con {fmt_int(top_prod['units'])} unidades.")

    cancel_rate = df["is_cancelled"].mean() * 100 if len(df) else 0
    if cancel_rate > 3:
        insights.append(f"La tasa de cancelación está en {cancel_rate:.1f}%, por encima de un nivel deseable.")

    sla_gt48 = (df["sla_bucket"] == ">48 h").mean() * 100 if len(df) else 0
    if sla_gt48 > 20:
        insights.append(f"El {sla_gt48:.1f}% de los registros cae en un SLA mayor a 48 horas.")

    if not forecast_df.empty:
        monthly_forecast = forecast_df["forecast_orders"].sum()
        insights.append(f"El forecast estima alrededor de {fmt_int(monthly_forecast)} pedidos para los próximos 30 días.")

    return insights


def generate_recommendations(df, warehouse_perf, carrier_perf, channel_perf, product_perf, forecast_df):
    recommendations = []

    if not warehouse_perf.empty:
        risk_wh = warehouse_perf.sort_values("performance_score").iloc[0]
        recommendations.append(f"Priorizar el seguimiento del almacén {risk_wh['warehouse']} por su score operativo más bajo.")

    if not carrier_perf.empty:
        risk_car = carrier_perf.sort_values("performance_score").iloc[0]
        recommendations.append(f"Revisar el desempeño del carrier {risk_car['carrier']} para mejorar cumplimiento y tiempos de ciclo.")

    if not channel_perf.empty:
        top_ch = channel_perf.iloc[0]
        recommendations.append(f"Monitorear estrechamente el canal {top_ch['channel']} por su peso en el volumen total.")

    if not product_perf.empty:
        top_prod = product_perf.iloc[0]
        recommendations.append(f"Dar seguimiento especial al producto {top_prod['product']} por su alta carga operativa.")

    avg_cycle = df["cycle_time_hours"].mean()
    if pd.notna(avg_cycle) and avg_cycle > 48:
        recommendations.append("Reducir tiempos de ciclo debe ser una prioridad operativa en el corto plazo.")

    if not forecast_df.empty:
        peak_day = forecast_df.sort_values("forecast_orders", ascending=False).iloc[0]
        recommendations.append(
            f"Preparar capacidad para picos próximos al {peak_day['analysis_date'].date()}, fecha con mayor carga proyectada."
        )

    return recommendations


def format_table_for_display(df, hour_cols=None, pct_cols=None, int_cols=None, score_cols=None):
    out = df.copy()
    for c in hour_cols or []:
        if c in out.columns:
            out[c] = out[c].apply(fmt_hours)
    for c in pct_cols or []:
        if c in out.columns:
            out[c] = out[c].apply(lambda x: f"{x:.1f}%")
    for c in int_cols or []:
        if c in out.columns:
            out[c] = out[c].apply(fmt_int)
    for c in score_cols or []:
        if c in out.columns:
            out[c] = out[c].apply(lambda x: f"{x:.1f}")
    return out

# =========================================================
# PDF - SIN CAMBIOS DE LÓGICA
# =========================================================
def draw_kpi_card(c, x, y, w, h, title, value, fill_color="#F5F7FA", value_color="#0B1F3A"):
    c.setFillColor(HexColor(fill_color))
    c.roundRect(x, y, w, h, 10, fill=1, stroke=0)
    c.setFillColor(HexColor("#5B6575"))
    c.setFont("Helvetica", 9)
    c.drawString(x + 10, y + h - 18, str(title))
    c.setFillColor(HexColor(value_color))
    c.setFont("Helvetica-Bold", 16)
    c.drawString(x + 10, y + 14, str(value))


def draw_section_title(c, x, y, title):
    c.setFillColor(HexColor("#0B1F3A"))
    c.setFont("Helvetica-Bold", 13)
    c.drawString(x, y, title)
    c.setStrokeColor(HexColor("#D9DEE7"))
    c.setLineWidth(1)
    c.line(x, y - 6, 560, y - 6)


def add_wrapped_text(c, text, x, y, max_width, font_name="Helvetica", font_size=10, line_height=14, bullet=False):
    prefix = "• " if bullet else ""
    lines = simpleSplit(prefix + text, font_name, font_size, max_width)
    c.setFont(font_name, font_size)
    for line in lines:
        c.drawString(x, y, line)
        y -= line_height
    return y


def make_pdf(summary_dict, insights, recommendations, sla_summary, forecast_df, forecast_model_name):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    c.setFillColor(HexColor("#0B1F3A"))
    c.rect(0, height - 110, width, 110, fill=1, stroke=0)
    c.setFillColor(HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 22)
    c.drawString(40, height - 50, "Imporey Internacional")
    c.setFont("Helvetica", 12)
    c.drawString(40, height - 72, "Dashboard Avanzado Operativo y Logístico")
    c.drawString(40, height - 90, "Resumen ejecutivo generado automáticamente")

    y = height - 145
    draw_section_title(c, 40, y, "KPIs principales")
    y -= 35

    cards = list(summary_dict.items())
    card_w = 160
    card_h = 52
    gap_x = 15
    start_x = 40
    row_y = y

    for i, (k, v) in enumerate(cards[:6]):
        x = start_x + (i % 3) * (card_w + gap_x)
        current_y = row_y - (i // 3) * 70
        draw_kpi_card(c, x, current_y, card_w, card_h, k, v)

    y = row_y - 160
    draw_section_title(c, 40, y, "Hallazgos clave")
    y -= 22
    c.setFillColor(HexColor("#111827"))
    for ins in insights[:5]:
        y = add_wrapped_text(c, ins, 50, y, 500, bullet=True)
        y -= 2

    c.showPage()
    c.setFillColor(HexColor("#0B1F3A"))
    c.rect(0, height - 90, width, 90, fill=1, stroke=0)
    c.setFillColor(HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 18)
    c.drawString(40, height - 50, "Imporey Internacional")
    c.setFont("Helvetica", 11)
    c.drawString(40, height - 70, "SLA, forecast y recomendaciones")

    y = height - 120

    if sla_summary is not None and not sla_summary.empty:
        draw_section_title(c, 40, y, "Distribución SLA")
        y -= 20
        for _, row in sla_summary.iterrows():
            c.drawString(50, y, f"{row['sla_bucket']}: {fmt_int(row['records'])} registros ({row['share']:.1f}%)")
            y -= 14
        y -= 10

    if forecast_df is not None and not forecast_df.empty:
        draw_section_title(c, 40, y, "Forecast siguiente mes")
        y -= 20
        total_f = forecast_df["forecast_orders"].sum()
        avg_f = forecast_df["forecast_orders"].mean()
        c.drawString(50, y, f"Modelo usado: {forecast_model_name}")
        y -= 14
        c.drawString(50, y, f"Pedidos pronosticados próximos 30 días: {fmt_int(total_f)}")
        y -= 14
        c.drawString(50, y, f"Promedio diario pronosticado: {fmt_int(avg_f)}")
        y -= 20

    draw_section_title(c, 40, y, "Recomendaciones")
    y -= 20
    for rec in recommendations[:6]:
        y = add_wrapped_text(c, rec, 50, y, 500, bullet=True)
        y -= 2

    c.setFillColor(HexColor("#9CA3AF"))
    c.setFont("Helvetica", 8)
    c.drawRightString(570, 20, "Imporey Internacional | Reporte generado automáticamente")
    c.save()
    buffer.seek(0)
    return buffer


# Mejora 1: cache para lectura del archivo Excel desde bytes.
@st.cache_data(show_spinner=False)
def load_excel_from_bytes(file_bytes):
    return pd.read_excel(BytesIO(file_bytes))

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown('<div class="sidebar-panel">', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-title">Imporey Internacional</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">Dashboard avanzado de operación, fulfillment y logística</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="sidebar-panel">', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-title">1. Carga de archivo</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"], label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# CARGA
# =========================================================
if uploaded_file is not None:
    try:
        with st.spinner("Procesando archivo y construyendo dashboard..."):
            raw_df = load_excel_from_bytes(uploaded_file.getvalue())
            original_columns = raw_df.columns.tolist()
            df = normalize_dataframe(raw_df)

        # Mejora 4: aviso visible si columnas críticas no fueron reconocidas.
        missing_critical_cols = validate_missing_critical_columns(df)
        if missing_critical_cols:
            st.warning(f"⚠️ Columnas críticas sin datos reconocidos: {missing_critical_cols}. Verifica el formato del Excel.")

        render_upload_success("Archivo cargado correctamente. El dashboard ya está listo para analizarse.")

        with st.expander("Ver columnas detectadas"):
            st.write(original_columns)

        # Filtros en sidebar, solo cuando ya existe data
        with st.sidebar:
            st.markdown('<div class="sidebar-panel">', unsafe_allow_html=True)
            st.markdown('<div class="sidebar-title">2. Filtros</div>', unsafe_allow_html=True)
            st.markdown('<div class="sidebar-sub">Refina el análisis por segmento operativo</div>', unsafe_allow_html=True)

            warehouses = sorted([x for x in df["warehouse"].dropna().unique() if str(x).strip() != ""])
            channels = sorted([x for x in df["channel"].dropna().unique() if str(x).strip() != ""])
            carriers = sorted([x for x in df["carrier"].dropna().unique() if str(x).strip() != ""])
            statuses = sorted([x for x in df["status"].dropna().unique() if str(x).strip() != ""])

            valid_dates = pd.to_datetime(df["analysis_date"], errors="coerce").dropna()
            if not valid_dates.empty:
                min_date = valid_dates.min().date()
                max_date = valid_dates.max().date()
                # Mejora 2: filtro de fecha en sidebar.
                selected_date_range = st.date_input("Rango de fechas", value=(min_date, max_date), min_value=min_date, max_value=max_date)
            else:
                selected_date_range = None
                st.info("No se detectaron fechas válidas para filtrar.")

            # Mejora 2: filtros principales en sidebar.
            selected_warehouses = st.multiselect("Almacén", warehouses, default=warehouses)
            selected_channels = st.multiselect("Canal", channels, default=channels)
            selected_carriers = st.multiselect("Carrier", carriers, default=carriers)
            selected_statuses = st.multiselect("Status", statuses, default=statuses)
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="sidebar-panel">', unsafe_allow_html=True)
            st.markdown('<div class="sidebar-title">3. Pesos del score</div>', unsafe_allow_html=True)
            st.markdown('<div class="sidebar-sub">Ajusta la importancia de cada componente del performance score.</div>', unsafe_allow_html=True)
            # Mejora 7: sliders para pesos del performance score.
            with st.expander("Configurar pesos", expanded=False):
                w_delivery_raw = st.slider("Entrega (%)", 0, 100, 40)
                w_cycle_raw = st.slider("Ciclo (%)", 0, 100, 30)
                w_cancel_raw = st.slider("Cancelación (%)", 0, 100, 20)
                w_volume_raw = st.slider("Volumen (%)", 0, 100, 10)
                total_weight = w_delivery_raw + w_cycle_raw + w_cancel_raw + w_volume_raw
                st.caption(f"Total configurado: {total_weight}%")
            score_weights = normalize_score_weights((w_delivery_raw, w_cycle_raw, w_cancel_raw, w_volume_raw))
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="sidebar-panel">', unsafe_allow_html=True)
            st.markdown('<div class="sidebar-title">4. Estado del modelo</div>', unsafe_allow_html=True)
            st.markdown('<div class="sidebar-sub">El forecast se genera automáticamente con el historial disponible.</div>', unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        filtered = df.copy()
        # Mejora 2: aplicar filtro de fechas antes de calcular métricas.
        if selected_date_range and len(selected_date_range) == 2:
            start_date, end_date = selected_date_range
            filtered_dates = pd.to_datetime(filtered["analysis_date"], errors="coerce").dt.date
            filtered = filtered[(filtered_dates >= start_date) & (filtered_dates <= end_date)]

        if selected_warehouses:
            filtered = filtered[filtered["warehouse"].isin(selected_warehouses)]
        if selected_channels:
            filtered = filtered[filtered["channel"].isin(selected_channels)]
        if selected_carriers:
            filtered = filtered[filtered["carrier"].isin(selected_carriers)]
        if selected_statuses:
            filtered = filtered[filtered["status"].isin(selected_statuses)]

        if filtered.empty:
            st.warning("No hay datos con esos filtros.")
            st.stop()

        total_rows = len(filtered)
        total_orders = safe_nunique(filtered["order_id"])
        total_shipments = safe_nunique(filtered["shipment_id"])
        total_units = filtered["qty"].sum()
        delivered_rate = filtered["is_delivered"].mean() * 100 if len(filtered) else 0
        picked_rate = filtered["is_picked"].mean() * 100 if len(filtered) else 0
        cancel_rate = filtered["is_cancelled"].mean() * 100 if len(filtered) else 0
        avg_cycle_time = filtered["cycle_time_hours"].mean()

        with st.spinner("Calculando métricas, SLA y forecast..."):
            daily_ops = build_daily_operations(filtered)
            forecast_df, forecast_model_name, forecast_quality = build_forecast_next_period(daily_ops, forecast_days=30)
            warehouse_perf = build_warehouse_performance(filtered, score_weights)
            carrier_perf = build_carrier_performance(filtered, score_weights)
            channel_perf = build_channel_performance(filtered)
            product_perf = build_product_performance(filtered)
            sla_summary = build_sla_summary(filtered)

            pct_lt24 = 0
            if not sla_summary.empty:
                row_lt24 = sla_summary[sla_summary["sla_bucket"] == "<24 h"]
                if not row_lt24.empty:
                    pct_lt24 = float(row_lt24.iloc[0]["share"])

            insights = generate_insights(filtered, warehouse_perf, carrier_perf, channel_perf, product_perf, forecast_df)
            recommendations = generate_recommendations(filtered, warehouse_perf, carrier_perf, channel_perf, product_perf, forecast_df)

        # KPIS
        st.subheader("Resumen ejecutivo")
        r1 = st.columns(4)
        r2 = st.columns(4)

        with r1[0]:
            render_kpi_card("Registros", fmt_int(total_rows), "Total de filas analizadas", tone=get_kpi_tone("Registros"))
        with r1[1]:
            render_kpi_card("Pedidos únicos", fmt_int(total_orders), "Órdenes identificadas", tone=get_kpi_tone("Pedidos únicos"))
        with r1[2]:
            render_kpi_card("Envíos únicos", fmt_int(total_shipments), "Embarques / envíos", tone=get_kpi_tone("Envíos únicos"))
        with r1[3]:
            render_kpi_card("Unidades", fmt_int(total_units), "Volumen total", tone=get_kpi_tone("Unidades"))

        with r2[0]:
            render_kpi_card("% Entregado", fmt_pct(delivered_rate), "Cumplimiento general", tone=get_kpi_tone("% Entregado", delivered_rate))
        with r2[1]:
            render_kpi_card("% Pickeado", fmt_pct(picked_rate), "Avance operativo", tone=get_kpi_tone("% Pickeado", picked_rate))
        with r2[2]:
            render_kpi_card("% Cancelado", fmt_pct(cancel_rate), "Riesgo operativo", tone=get_kpi_tone("% Cancelado", cancel_rate))
        with r2[3]:
            render_kpi_card("SLA <24 h", fmt_pct(pct_lt24), "Velocidad de cierre", tone=get_kpi_tone("SLA <24 h", pct_lt24))

        st.caption(f"Tiempo ciclo promedio: {fmt_hours(avg_cycle_time)}")

        if not forecast_df.empty:
            st.markdown(
                f"""
                <div class="pill-row">
                    <span class="small-chip">📈 Modelo: {forecast_model_name}</span>
                    <span class="small-chip">📦 Forecast 30 días: {fmt_int(forecast_df['forecast_orders'].sum())} pedidos</span>
                    <span class="small-chip">📅 Promedio diario: {fmt_int(forecast_df['forecast_orders'].mean())}</span>
                </div>
                """,
                unsafe_allow_html=True
            )
            # Mejora 5: métricas de calidad del forecast visibles al usuario.
            mae_txt = "N/A" if pd.isna(forecast_quality.get("mae_7d")) else f"{forecast_quality.get('mae_7d'):.1f} pedidos"
            mape_txt = "N/A" if pd.isna(forecast_quality.get("mape_7d")) else f"{forecast_quality.get('mape_7d'):.1f}%"
            st.caption(f"Calidad del forecast | MAE últimos 7 días: {mae_txt} | MAPE últimos 7 días: {mape_txt}")

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "Resumen",
            "Forecast",
            "SLA",
            "Operación",
            "Detalle"
        ])

        with tab1:
            c_left, c_right = st.columns([1.3, 1])

            with c_left:
                render_section_open("Tendencia operativa diaria")

                if not daily_ops.empty:
                    fig_orders = go.Figure()
                    fig_orders.add_trace(go.Scatter(
                        x=daily_ops["analysis_date"],
                        y=daily_ops["orders"],
                        mode="lines+markers",
                        name="Pedidos",
                        line=dict(color="#173B63", width=3),
                        marker=dict(size=6)
                    ))
                    fig_orders.add_trace(go.Scatter(
                        x=daily_ops["analysis_date"],
                        y=daily_ops["delivered"],
                        mode="lines",
                        name="Entregados",
                        line=dict(color="#16a34a", width=2)
                    ))
                    fig_orders.add_trace(go.Scatter(
                        x=daily_ops["analysis_date"],
                        y=daily_ops["cancelled"],
                        mode="lines",
                        name="Cancelados",
                        line=dict(color="#dc2626", width=2)
                    ))
                    fig_orders.update_layout(
                        height=420,
                        legend_title="Serie",
                        xaxis_title="Fecha",
                        yaxis_title="Volumen"
                    )
                    set_plotly_theme(fig_orders)
                    st.plotly_chart(fig_orders, use_container_width=True)

                render_section_close()

            with c_right:
                render_section_open("Hallazgos automáticos")
                render_insight_list(insights)

                st.markdown('<div class="section-title" style="margin-top:14px;">Recomendaciones ejecutivas</div>', unsafe_allow_html=True)
                render_recommend_list(recommendations)
                render_section_close()

        with tab2:
            render_section_open("Pronóstico del siguiente mes")

            if not forecast_df.empty and not daily_ops.empty:
                forecast_total = forecast_df["forecast_orders"].sum()
                forecast_avg = forecast_df["forecast_orders"].mean()
                forecast_high = forecast_df["upper_bound"].sum()
                forecast_low = forecast_df["lower_bound"].sum()

                m1, m2, m3 = st.columns(3)
                with m1:
                    render_mini_metric("Pedidos esperados próximos 30 días", fmt_int(forecast_total))
                with m2:
                    render_mini_metric("Promedio diario esperado", fmt_int(forecast_avg))
                with m3:
                    render_mini_metric("Rango estimado", f"{fmt_int(forecast_low)} - {fmt_int(forecast_high)}")

                q1, q2 = st.columns(2)
                mae_txt = "N/A" if pd.isna(forecast_quality.get("mae_7d")) else f"{forecast_quality.get('mae_7d'):.1f} pedidos"
                mape_txt = "N/A" if pd.isna(forecast_quality.get("mape_7d")) else f"{forecast_quality.get('mape_7d'):.1f}%"
                with q1:
                    render_mini_metric("MAE últimos 7 días", mae_txt)
                with q2:
                    render_mini_metric("MAPE últimos 7 días", mape_txt)

                hist = daily_ops[["analysis_date", "orders"]].copy()
                fc = forecast_df[["analysis_date", "forecast_orders", "lower_bound", "upper_bound"]].copy()

                fig_fc = go.Figure()
                fig_fc.add_trace(go.Scatter(
                    x=hist["analysis_date"],
                    y=hist["orders"],
                    mode="lines",
                    name="Histórico",
                    line=dict(width=3, color="#173B63")
                ))
                fig_fc.add_trace(go.Scatter(
                    x=fc["analysis_date"],
                    y=fc["upper_bound"],
                    mode="lines",
                    line=dict(width=0),
                    showlegend=False,
                    hoverinfo="skip"
                ))
                fig_fc.add_trace(go.Scatter(
                    x=fc["analysis_date"],
                    y=fc["lower_bound"],
                    mode="lines",
                    line=dict(width=0),
                    fill="tonexty",
                    name="Banda estimada",
                    fillcolor="rgba(59,130,246,0.16)"
                ))
                fig_fc.add_trace(go.Scatter(
                    x=fc["analysis_date"],
                    y=fc["forecast_orders"],
                    mode="lines+markers",
                    name="Pronóstico",
                    line=dict(dash="dash", width=3, color="#2563eb"),
                    marker=dict(size=6)
                ))

                fig_fc.update_layout(
                    height=500,
                    xaxis_title="Fecha",
                    yaxis_title="Pedidos",
                    legend_title="Serie"
                )
                set_plotly_theme(fig_fc)
                st.plotly_chart(fig_fc, use_container_width=True)

                forecast_show = fc.copy()
                forecast_show["analysis_date"] = forecast_show["analysis_date"].dt.date
                forecast_show["forecast_orders"] = forecast_show["forecast_orders"].round(0).astype(int)
                forecast_show["lower_bound"] = forecast_show["lower_bound"].round(0).astype(int)
                forecast_show["upper_bound"] = forecast_show["upper_bound"].round(0).astype(int)
                forecast_show = forecast_show.rename(columns={
                    "analysis_date": "Fecha",
                    "forecast_orders": "Pedidos pronosticados",
                    "lower_bound": "Escenario bajo",
                    "upper_bound": "Escenario alto"
                })
                st.dataframe(forecast_show, use_container_width=True, hide_index=True, column_config=common_column_config())
            else:
                st.info("No hay suficiente historial diario para generar un forecast robusto.")

            render_section_close()

        with tab3:
            c1, c2 = st.columns([1, 2])

            with c1:
                render_section_open("Distribución SLA")
                if not sla_summary.empty:
                    show = sla_summary.copy()
                    # Mejora 3: column_config conserva tipos numéricos y permite ordenar correctamente.
                    st.dataframe(show, use_container_width=True, hide_index=True, column_config=common_column_config())
                else:
                    st.info("No hay datos suficientes para SLA.")
                render_section_close()

            with c2:
                render_section_open("Registros por rango SLA")
                if not sla_summary.empty:
                    fig_sla = px.bar(
                        sla_summary,
                        x="sla_bucket",
                        y="records",
                        text="records"
                    )
                    fig_sla.update_traces(
                        marker_color=["#16a34a", "#f59e0b", "#dc2626"]
                    )
                    fig_sla.update_layout(
                        height=430,
                        xaxis_title="Rango SLA",
                        yaxis_title="Registros"
                    )
                    set_plotly_theme(fig_sla)
                    st.plotly_chart(fig_sla, use_container_width=True)
                else:
                    st.info("No hay datos suficientes para SLA.")
                render_section_close()

        with tab4:
            c1, c2 = st.columns(2)

            with c1:
                render_section_open("Almacenes por score")
                if not warehouse_perf.empty:
                    top_wh = warehouse_perf.head(10).sort_values("performance_score", ascending=True)
                    fig_wh = px.bar(
                        top_wh,
                        x="performance_score",
                        y="warehouse",
                        orientation="h",
                        text="performance_score"
                    )
                    fig_wh.update_traces(marker_color="#173B63")
                    fig_wh.update_layout(
                        height=450,
                        xaxis_title="Score",
                        yaxis_title=""
                    )
                    set_plotly_theme(fig_wh)
                    st.plotly_chart(fig_wh, use_container_width=True)
                render_section_close()

            with c2:
                render_section_open("Carriers por score")
                if not carrier_perf.empty:
                    top_car = carrier_perf.head(10).sort_values("performance_score", ascending=True)
                    fig_car = px.bar(
                        top_car,
                        x="performance_score",
                        y="carrier",
                        orientation="h",
                        text="performance_score"
                    )
                    fig_car.update_traces(marker_color="#2563eb")
                    fig_car.update_layout(
                        height=450,
                        xaxis_title="Score",
                        yaxis_title=""
                    )
                    set_plotly_theme(fig_car)
                    st.plotly_chart(fig_car, use_container_width=True)
                render_section_close()

            c3, c4 = st.columns(2)

            with c3:
                render_section_open("Pedidos por canal")
                if not channel_perf.empty:
                    top_ch = channel_perf.head(10).sort_values("orders", ascending=True)
                    fig_ch = px.bar(
                        top_ch,
                        x="orders",
                        y="channel",
                        orientation="h",
                        text="orders"
                    )
                    fig_ch.update_traces(marker_color="#0f766e")
                    fig_ch.update_layout(
                        height=430,
                        xaxis_title="Pedidos",
                        yaxis_title=""
                    )
                    set_plotly_theme(fig_ch)
                    st.plotly_chart(fig_ch, use_container_width=True)
                render_section_close()

            with c4:
                render_section_open("Top productos")
                if not product_perf.empty:
                    top_prod = product_perf.head(10).sort_values("units", ascending=True)
                    fig_prod = px.bar(
                        top_prod,
                        x="units",
                        y="product",
                        orientation="h",
                        text="units"
                    )
                    fig_prod.update_traces(marker_color="#7c3aed")
                    fig_prod.update_layout(
                        height=430,
                        xaxis_title="Unidades",
                        yaxis_title=""
                    )
                    set_plotly_theme(fig_prod)
                    st.plotly_chart(fig_prod, use_container_width=True)
                render_section_close()

        with tab5:
            d1, d2 = st.columns(2)

            with d1:
                render_section_open("Detalle de almacenes")
                if not warehouse_perf.empty:
                    # Mejora 3: uso de column_config para mantener tipos y añadir barra visual al score.
                    st.dataframe(
                        warehouse_perf[[
                            "warehouse", "status_light", "performance_score", "orders", "shipments",
                            "units", "delivered_rate", "cancelled_rate", "avg_cycle_hours",
                            "avg_time_spent", "share_orders"
                        ]],
                        use_container_width=True,
                        hide_index=True,
                        column_config=common_column_config()
                    )
                render_section_close()

            with d2:
                render_section_open("Detalle de carriers")
                if not carrier_perf.empty:
                    # Mejora 3: uso de column_config para mantener tipos y añadir barra visual al score.
                    st.dataframe(
                        carrier_perf[[
                            "carrier", "status_light", "performance_score", "shipments", "orders",
                            "units", "delivered_rate", "cancelled_rate", "avg_cycle_hours",
                            "avg_time_spent", "share_shipments"
                        ]],
                        use_container_width=True,
                        hide_index=True,
                        column_config=common_column_config()
                    )
                render_section_close()

            d3, d4 = st.columns(2)

            with d3:
                render_section_open("Detalle de canales")
                if not channel_perf.empty:
                    # Mejora 3: column_config conserva ordenamiento numérico.
                    st.dataframe(
                        channel_perf,
                        use_container_width=True,
                        hide_index=True,
                        column_config=common_column_config()
                    )
                render_section_close()

            with d4:
                render_section_open("Detalle de productos")
                if not product_perf.empty:
                    # Mejora 3: column_config conserva ordenamiento numérico.
                    st.dataframe(
                        product_perf.head(25),
                        use_container_width=True,
                        hide_index=True,
                        column_config=common_column_config()
                    )
                render_section_close()

            with st.expander("Vista previa de datos procesados"):
                st.dataframe(filtered.head(100), use_container_width=True)

        summary_dict = {
            "Registros": fmt_int(total_rows),
            "Pedidos únicos": fmt_int(total_orders),
            "Envíos únicos": fmt_int(total_shipments),
            "Unidades": fmt_int(total_units),
            "% Entregado": fmt_pct(delivered_rate),
            "SLA <24h": fmt_pct(pct_lt24),
            "% Cancelado": fmt_pct(cancel_rate),
            "Tiempo ciclo prom.": fmt_hours(avg_cycle_time),
        }

        pdf_buffer = make_pdf(
            summary_dict,
            insights,
            recommendations,
            sla_summary,
            forecast_df,
            forecast_model_name
        )

        # Mejora 6: botón de PDF en sidebar para que esté visible sin bajar al final del dashboard.
        with st.sidebar:
            st.markdown('<div class="sidebar-panel">', unsafe_allow_html=True)
            st.markdown('<div class="sidebar-title">5. Exportar reporte</div>', unsafe_allow_html=True)
            st.download_button(
                label="⬇️ Descargar reporte PDF",
                data=pdf_buffer,
                file_name="dashboard_avanzado_imporey.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="footer-note">Dashboard premium listo para uso ejecutivo y presentación a cliente interno o externo.</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
else:
    render_upload_info("Sube un archivo Excel desde la barra lateral para generar el dashboard premium.")

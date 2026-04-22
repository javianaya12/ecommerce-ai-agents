import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from reportlab.lib.colors import HexColor

st.set_page_config(
    page_title="Imporey Internacional | Dashboard Avanzado",
    layout="wide"
)

# =========================================================
# ESTILOS
# =========================================================
st.markdown("""
<style>
.main {
    background-color: #f7f9fc;
}
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
}
.kpi-card {
    background: white;
    border-radius: 14px;
    padding: 18px 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    border-left: 6px solid #0B1F3A;
    margin-bottom: 12px;
}
.kpi-title {
    color: #64748B;
    font-size: 0.9rem;
    margin-bottom: 8px;
}
.kpi-value {
    color: #0B1F3A;
    font-size: 1.8rem;
    font-weight: 700;
}
.section-box {
    background: white;
    border-radius: 14px;
    padding: 16px 18px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    margin-bottom: 18px;
}
.hero-box {
    background: linear-gradient(135deg, #0B1F3A 0%, #163A63 100%);
    color: white;
    padding: 24px 28px;
    border-radius: 18px;
    margin-bottom: 20px;
}
.hero-title {
    font-size: 2rem;
    font-weight: 700;
    margin-bottom: 6px;
}
.hero-subtitle {
    font-size: 1rem;
    color: #D9E3F0;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# HEADER
# =========================================================
st.markdown("""
<div class="hero-box">
    <div class="hero-title">Imporey Internacional</div>
    <div class="hero-subtitle">Dashboard Avanzado de Operación, Fulfillment y Logística</div>
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
# HELPERS
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

    df["analysis_date"] = df["created_on"].dt.date
    df["analysis_datetime"] = df["created_on"]

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


def calc_score(df_base, volume_col):
    df = df_base.copy()

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

    df["performance_score"] = (
        df["delivery_score"] * 0.40 +
        df["cycle_score"] * 0.30 +
        df["cancel_score"] * 0.20 +
        df["volume_score"] * 0.10
    ).round(1)

    df["status_light"] = df["performance_score"].apply(get_status_light)
    return df


def build_daily_operations(df):
    return (
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


def build_forecast_next_period(daily_ops, forecast_days=30):
    if daily_ops.empty or len(daily_ops) < 2:
        return pd.DataFrame()

    tmp = daily_ops.copy()
    tmp["analysis_date"] = pd.to_datetime(tmp["analysis_date"])
    tmp = tmp.sort_values("analysis_date").reset_index(drop=True)

    x = np.arange(len(tmp))
    y = tmp["orders"].astype(float).values

    try:
        slope, intercept = np.polyfit(x, y, 1)
    except Exception:
        return pd.DataFrame()

    future_x = np.arange(len(tmp), len(tmp) + forecast_days)
    future_y = slope * future_x + intercept
    future_y = np.where(future_y < 0, 0, future_y)

    last_date = tmp["analysis_date"].max()
    future_dates = pd.date_range(last_date + pd.Timedelta(days=1), periods=forecast_days, freq="D")

    forecast_df = pd.DataFrame({
        "analysis_date": future_dates,
        "forecast_orders": future_y
    })

    return forecast_df


def build_warehouse_performance(df):
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
    base = calc_score(base, "orders")
    return base.sort_values("orders", ascending=False)


def build_carrier_performance(df):
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
    base = calc_score(base, "shipments")
    return base.sort_values("shipments", ascending=False)


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


def generate_insights(df, warehouse_perf, carrier_perf, channel_perf, product_perf):
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

    return insights


def generate_recommendations(df, warehouse_perf, carrier_perf, channel_perf, product_perf):
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


def render_kpi_card(title, value):
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)


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


def make_pdf(summary_dict, insights, recommendations, sla_summary):
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
    c.drawString(40, height - 70, "SLA y recomendaciones")

    y = height - 120

    if sla_summary is not None and not sla_summary.empty:
        draw_section_title(c, 40, y, "Distribución SLA")
        y -= 20
        for _, row in sla_summary.iterrows():
            c.drawString(50, y, f"{row['sla_bucket']}: {fmt_int(row['records'])} registros ({row['share']:.1f}%)")
            y -= 14
        y -= 10

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


# =========================================================
# CARGA
# =========================================================
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    try:
        with st.spinner("Leyendo archivo..."):
            df = pd.read_excel(uploaded_file)
            original_columns = df.columns.tolist()
            df = normalize_dataframe(df)

        st.success("Archivo cargado correctamente")

        with st.expander("Columnas detectadas"):
            st.write(original_columns)

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.subheader("Filtros")

        c1, c2, c3, c4 = st.columns(4)
        warehouses = sorted([x for x in df["warehouse"].dropna().unique() if str(x).strip() != ""])
        channels = sorted([x for x in df["channel"].dropna().unique() if str(x).strip() != ""])
        carriers = sorted([x for x in df["carrier"].dropna().unique() if str(x).strip() != ""])
        statuses = sorted([x for x in df["status"].dropna().unique() if str(x).strip() != ""])

        selected_warehouses = c1.multiselect("Almacén", warehouses, default=warehouses)
        selected_channels = c2.multiselect("Canal", channels, default=channels)
        selected_carriers = c3.multiselect("Carrier", carriers, default=carriers)
        selected_statuses = c4.multiselect("Status", statuses, default=statuses)
        st.markdown("</div>", unsafe_allow_html=True)

        filtered = df.copy()
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

        sla_summary = build_sla_summary(filtered)
        pct_lt24 = 0
        if not sla_summary.empty:
            row_lt24 = sla_summary[sla_summary["sla_bucket"] == "<24 h"]
            if not row_lt24.empty:
                pct_lt24 = float(row_lt24.iloc[0]["share"])

        st.subheader("Resumen ejecutivo")
        r1 = st.columns(4)
        r2 = st.columns(4)

        with r1[0]:
            render_kpi_card("Registros", fmt_int(total_rows))
        with r1[1]:
            render_kpi_card("Pedidos únicos", fmt_int(total_orders))
        with r1[2]:
            render_kpi_card("Envíos únicos", fmt_int(total_shipments))
        with r1[3]:
            render_kpi_card("Unidades", fmt_int(total_units))

        with r2[0]:
            render_kpi_card("% Entregado", fmt_pct(delivered_rate))
        with r2[1]:
            render_kpi_card("% Pickeado", fmt_pct(picked_rate))
        with r2[2]:
            render_kpi_card("% Cancelado", fmt_pct(cancel_rate))
        with r2[3]:
            render_kpi_card("SLA <24 h", fmt_pct(pct_lt24))

        st.caption(f"Tiempo ciclo promedio: {fmt_hours(avg_cycle_time)}")

        daily_ops = build_daily_operations(filtered)
        forecast_df = build_forecast_next_period(daily_ops, forecast_days=30)
        warehouse_perf = build_warehouse_performance(filtered)
        carrier_perf = build_carrier_performance(filtered)
        channel_perf = build_channel_performance(filtered)
        product_perf = build_product_performance(filtered)

        insights = generate_insights(filtered, warehouse_perf, carrier_perf, channel_perf, product_perf)
        recommendations = generate_recommendations(filtered, warehouse_perf, carrier_perf, channel_perf, product_perf)

        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown('<div class="section-box">', unsafe_allow_html=True)
            st.subheader("Tendencia operativa diaria")
            if not daily_ops.empty:
                fig1, ax1 = plt.subplots(figsize=(8, 4))
                ax1.plot(pd.to_datetime(daily_ops["analysis_date"]), daily_ops["orders"], marker="o")
                ax1.set_title("Pedidos por día")
                ax1.set_xlabel("Fecha")
                ax1.set_ylabel("Pedidos")
                plt.xticks(rotation=45)
                st.pyplot(fig1)

                fig2, ax2 = plt.subplots(figsize=(8, 4))
                ax2.plot(pd.to_datetime(daily_ops["analysis_date"]), daily_ops["delivered"], label="Entregados")
                ax2.plot(pd.to_datetime(daily_ops["analysis_date"]), daily_ops["cancelled"], label="Cancelados")
                ax2.set_title("Entregados vs cancelados")
                ax2.legend()
                plt.xticks(rotation=45)
                st.pyplot(fig2)
            st.markdown("</div>", unsafe_allow_html=True)

        with col_b:
            st.markdown('<div class="section-box">', unsafe_allow_html=True)
            st.subheader("Distribución SLA")
            if not sla_summary.empty:
                fig_sla, ax_sla = plt.subplots(figsize=(8, 4))
                ax_sla.bar(sla_summary["sla_bucket"], sla_summary["records"])
                ax_sla.set_title("Registros por rango SLA")
                ax_sla.set_xlabel("SLA")
                ax_sla.set_ylabel("Registros")
                st.pyplot(fig_sla)

                st.dataframe(
                    format_table_for_display(
                        sla_summary,
                        pct_cols=["share"],
                        int_cols=["records"]
                    ),
                    use_container_width=True
                )
            else:
                st.info("No hay datos suficientes para SLA.")
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.subheader("Pronóstico del siguiente mes")
        if not forecast_df.empty and not daily_ops.empty:
            fig_forecast, ax_forecast = plt.subplots(figsize=(10, 4))

            hist_dates = pd.to_datetime(daily_ops["analysis_date"])
            ax_forecast.plot(hist_dates, daily_ops["orders"], marker="o", label="Histórico pedidos")

            ax_forecast.plot(
                forecast_df["analysis_date"],
                forecast_df["forecast_orders"],
                marker="o",
                linestyle="--",
                label="Pronóstico siguiente mes"
            )

            ax_forecast.set_title("Pronóstico de pedidos para los próximos 30 días")
            ax_forecast.set_xlabel("Fecha")
            ax_forecast.set_ylabel("Pedidos")
            ax_forecast.legend()
            plt.xticks(rotation=45)
            st.pyplot(fig_forecast)

            forecast_show = forecast_df.copy()
            forecast_show["analysis_date"] = forecast_show["analysis_date"].dt.date
            forecast_show["forecast_orders"] = forecast_show["forecast_orders"].round(0).astype(int)
            forecast_show = forecast_show.rename(columns={
                "analysis_date": "Fecha pronosticada",
                "forecast_orders": "Pedidos pronosticados"
            })
            st.dataframe(forecast_show, use_container_width=True)
        else:
            st.info("No hay suficientes datos diarios para generar el pronóstico.")
        st.markdown("</div>", unsafe_allow_html=True)

        col_c, col_d = st.columns(2)

        with col_c:
            st.markdown('<div class="section-box">', unsafe_allow_html=True)
            st.subheader("Top 10 almacenes por score")
            if not warehouse_perf.empty:
                top_wh = warehouse_perf.head(10).sort_values("performance_score", ascending=True)
                fig3, ax3 = plt.subplots(figsize=(8, 5))
                ax3.barh(top_wh["warehouse"], top_wh["performance_score"])
                ax3.set_title("Score de desempeño por almacén")
                ax3.set_xlabel("Score")
                st.pyplot(fig3)

                st.dataframe(
                    format_table_for_display(
                        warehouse_perf[[
                            "warehouse", "status_light", "performance_score", "orders", "shipments",
                            "units", "delivered_rate", "cancelled_rate", "avg_cycle_hours",
                            "avg_time_spent", "share_orders"
                        ]],
                        hour_cols=["avg_cycle_hours", "avg_time_spent"],
                        pct_cols=["delivered_rate", "cancelled_rate", "share_orders"],
                        int_cols=["orders", "shipments", "units"],
                        score_cols=["performance_score"]
                    ),
                    use_container_width=True
                )
            st.markdown("</div>", unsafe_allow_html=True)

        with col_d:
            st.markdown('<div class="section-box">', unsafe_allow_html=True)
            st.subheader("Top 10 carriers por score")
            if not carrier_perf.empty:
                top_car = carrier_perf.head(10).sort_values("performance_score", ascending=True)
                fig4, ax4 = plt.subplots(figsize=(8, 5))
                ax4.barh(top_car["carrier"], top_car["performance_score"])
                ax4.set_title("Score de desempeño por carrier")
                ax4.set_xlabel("Score")
                st.pyplot(fig4)

                st.dataframe(
                    format_table_for_display(
                        carrier_perf[[
                            "carrier", "status_light", "performance_score", "shipments", "orders",
                            "units", "delivered_rate", "cancelled_rate", "avg_cycle_hours",
                            "avg_time_spent", "share_shipments"
                        ]],
                        hour_cols=["avg_cycle_hours", "avg_time_spent"],
                        pct_cols=["delivered_rate", "cancelled_rate", "share_shipments"],
                        int_cols=["shipments", "orders", "units"],
                        score_cols=["performance_score"]
                    ),
                    use_container_width=True
                )
            st.markdown("</div>", unsafe_allow_html=True)

        col_e, col_f = st.columns(2)

        with col_e:
            st.markdown('<div class="section-box">', unsafe_allow_html=True)
            st.subheader("Canales")
            if not channel_perf.empty:
                fig5, ax5 = plt.subplots(figsize=(8, 4))
                top_ch = channel_perf.head(10).sort_values("orders", ascending=True)
                ax5.barh(top_ch["channel"], top_ch["orders"])
                ax5.set_title("Pedidos por canal")
                ax5.set_xlabel("Pedidos")
                st.pyplot(fig5)
                st.dataframe(
                    format_table_for_display(
                        channel_perf,
                        hour_cols=["avg_cycle_hours"],
                        pct_cols=["delivered_rate", "cancelled_rate", "share_orders"],
                        int_cols=["orders", "shipments", "units"]
                    ),
                    use_container_width=True
                )
            st.markdown("</div>", unsafe_allow_html=True)

        with col_f:
            st.markdown('<div class="section-box">', unsafe_allow_html=True)
            st.subheader("Top productos")
            if not product_perf.empty:
                fig6, ax6 = plt.subplots(figsize=(8, 4))
                top_prod = product_perf.head(10).sort_values("units", ascending=True)
                ax6.barh(top_prod["product"], top_prod["units"])
                ax6.set_title("Top 10 productos por unidades")
                ax6.set_xlabel("Unidades")
                st.pyplot(fig6)
                st.dataframe(
                    format_table_for_display(
                        product_perf.head(20),
                        pct_cols=["delivered_rate", "cancelled_rate", "share_units"],
                        int_cols=["orders", "shipments", "units"]
                    ),
                    use_container_width=True
                )
            st.markdown("</div>", unsafe_allow_html=True)

        col_g, col_h = st.columns(2)

        with col_g:
            st.markdown('<div class="section-box">', unsafe_allow_html=True)
            st.subheader("Hallazgos automáticos")
            for i, insight in enumerate(insights, start=1):
                st.write(f"{i}. {insight}")
            st.markdown("</div>", unsafe_allow_html=True)

        with col_h:
            st.markdown('<div class="section-box">', unsafe_allow_html=True)
            st.subheader("Recomendaciones ejecutivas")
            for i, rec in enumerate(recommendations, start=1):
                st.write(f"{i}. {rec}")
            st.markdown("</div>", unsafe_allow_html=True)

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

        pdf_buffer = make_pdf(summary_dict, insights, recommendations, sla_summary)

        st.download_button(
            label="Descargar PDF avanzado",
            data=pdf_buffer,
            file_name="dashboard_avanzado_imporey.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
else:
    st.info("Sube tu archivo Excel para comenzar.")

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from reportlab.lib.colors import HexColor

st.set_page_config(page_title="E-Commerce Operations Dashboard", layout="wide")

# =========================================================
# CONFIG
# =========================================================
st.title("E-Commerce Operations Dashboard")
st.caption("Análisis automático de operación, fulfillment y logística a partir de Excel")

# =========================================================
# HELPERS
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

REQUIRED_COLUMNS = list(COLUMN_MAP.keys())


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
        df[c] = df[c].astype(str).fillna("").replace("nan", "").str.strip()

    df["qty"] = df["qty"].fillna(0)

    df["is_delivered"] = np.where(
        (df["delivered_flag"] > 0) | (df["status"].str.lower() == "entregado"),
        1, 0
    )

    df["is_picked"] = np.where(
        (df["picked_flag"] > 0),
        1, 0
    )

    df["is_cancelled"] = np.where(
        df["status"].str.lower().str.contains("cancel", na=False),
        1, 0
    )

    df["is_collected"] = np.where(
        df["status"].str.lower().str.contains("recolect", na=False),
        1, 0
    )

    df["analysis_date"] = df["created_on"].dt.date
    df["time_spent_hours"] = df["time_spent"].apply(parse_duration_to_hours)

    df["cycle_time_hours"] = (
        (df["delivery_finished"] - df["created_on"]).dt.total_seconds() / 3600
    )

    df["pick_time_hours"] = (
        (df["pick_finished"] - df["created_on"]).dt.total_seconds() / 3600
    )

    df["pack_time_hours"] = (
        (df["pack_finished"] - df["created_on"]).dt.total_seconds() / 3600
    )

    return df


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
    except:
        return np.nan


def fmt_int(x):
    try:
        return f"{int(x):,}"
    except:
        return "0"


def fmt_pct(x):
    try:
        return f"{x:.1f}%"
    except:
        return "0.0%"


def fmt_hours(x):
    try:
        if pd.isna(x):
            return "-"
        return f"{x:.1f} h"
    except:
        return "-"


def safe_nunique(series):
    return series.dropna().nunique()


def generate_insights(df, daily_ops, warehouse_perf, carrier_perf, channel_perf, product_perf):
    insights = []

    total_orders = safe_nunique(df["order_id"])
    total_shipments = safe_nunique(df["shipment_id"])
    delivered_rate = df["is_delivered"].mean() * 100 if len(df) else 0
    cancel_rate = df["is_cancelled"].mean() * 100 if len(df) else 0

    if not warehouse_perf.empty:
        top_wh = warehouse_perf.iloc[0]
        insights.append(
            f"El almacén con mayor volumen es {top_wh['warehouse']} con {fmt_int(top_wh['orders'])} pedidos "
            f"({top_wh['share_orders']:.1f}% del total)."
        )

    if len(warehouse_perf) > 1:
        valid_wh = warehouse_perf[warehouse_perf["orders"] > 0].copy()
        if not valid_wh.empty:
            worst_wh = valid_wh.sort_values("delivered_rate").iloc[0]
            insights.append(
                f"El almacén con menor porcentaje de entrega es {worst_wh['warehouse']} "
                f"con {worst_wh['delivered_rate']:.1f}%."
            )

    if not carrier_perf.empty:
        top_carrier = carrier_perf.iloc[0]
        insights.append(
            f"El carrier principal es {top_carrier['carrier']} con {fmt_int(top_carrier['shipments'])} envíos "
            f"({top_carrier['share_shipments']:.1f}% del total)."
        )

        valid_carrier = carrier_perf[carrier_perf["shipments"] > 0].copy()
        if not valid_carrier.empty:
            worst_carrier = valid_carrier.sort_values("delivered_rate").iloc[0]
            insights.append(
                f"El carrier con menor cumplimiento de entrega es {worst_carrier['carrier']} "
                f"con {worst_carrier['delivered_rate']:.1f}%."
            )

    if not channel_perf.empty:
        top_channel = channel_perf.iloc[0]
        insights.append(
            f"El canal con más pedidos es {top_channel['channel']} con {fmt_int(top_channel['orders'])} pedidos."
        )

    if not product_perf.empty:
        top_product = product_perf.iloc[0]
        insights.append(
            f"El producto con mayor movimiento es {top_product['product']} con {fmt_int(top_product['units'])} unidades."
        )

    if len(daily_ops) >= 7:
        recent_avg = daily_ops["orders"].tail(7).mean()
        global_avg = daily_ops["orders"].mean()
        if recent_avg > global_avg * 1.15:
            insights.append(
                "En los últimos días el volumen operativo está por encima del promedio general; conviene revisar capacidad operativa y distribución de carga entre carriers y almacenes."
            )
        elif recent_avg < global_avg * 0.85:
            insights.append(
                "En los últimos días el volumen operativo está por debajo del promedio; puede ser una ventana útil para ajustes operativos, mantenimiento o capacitación."
            )

    if cancel_rate > 3:
        insights.append(
            f"La tasa de cancelación está en {cancel_rate:.1f}%, lo que amerita revisar causas de cancelación, disponibilidad operativa y promesa logística."
        )

    if delivered_rate < 90:
        insights.append(
            f"El porcentaje entregado es {delivered_rate:.1f}%, por debajo de un nivel operativo fuerte; conviene revisar SLA por carrier y almacén."
        )

    if total_orders > 0 and total_shipments > 0 and total_shipments < total_orders * 0.8:
        insights.append(
            "Hay una diferencia relevante entre pedidos y envíos únicos; puede valer la pena revisar consolidación, órdenes incompletas o estatus pendientes."
        )

    return insights


def generate_recommendations(df, warehouse_perf, carrier_perf, channel_perf, product_perf):
    recommendations = []

    if not warehouse_perf.empty:
        top_wh = warehouse_perf.iloc[0]
        recommendations.append(
            f"Priorizar monitoreo del almacén {top_wh['warehouse']}, ya que concentra la mayor carga operativa y tiene un impacto directo en el desempeño general."
        )

    if not carrier_perf.empty:
        valid = carrier_perf[carrier_perf["shipments"] > 0].copy()
        if not valid.empty:
            worst = valid.sort_values("delivered_rate").iloc[0]
            recommendations.append(
                f"Revisar el desempeño del carrier {worst['carrier']}, ya que presenta el menor nivel de cumplimiento en entregas dentro de los datos analizados."
            )

    if not channel_perf.empty:
        top_channel = channel_perf.iloc[0]
        recommendations.append(
            f"Dar seguimiento especial al canal {top_channel['channel']}, ya que concentra el mayor volumen de pedidos y puede definir la presión operativa diaria."
        )

    if not product_perf.empty:
        top_product = product_perf.iloc[0]
        recommendations.append(
            f"Considerar una gestión prioritaria del producto {top_product['product']}, al ser el de mayor movimiento y carga operativa en el periodo."
        )

    cancel_rate = df["is_cancelled"].mean() * 100 if len(df) else 0
    if cancel_rate > 2:
        recommendations.append(
            "Implementar una revisión puntual de cancelaciones por origen, carrier y canal para reducir fricción operativa y evitar impacto en la promesa al cliente."
        )

    avg_cycle_time = df["cycle_time_hours"].mean()
    if pd.notna(avg_cycle_time) and avg_cycle_time > 48:
        recommendations.append(
            "Evaluar oportunidades de mejora en tiempos de ciclo, especialmente en pedidos con mayor permanencia entre creación y finalización."
        )

    return recommendations


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
    return daily


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
    base = base.sort_values("orders", ascending=False)
    return base


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
    base = base.sort_values("shipments", ascending=False)
    return base


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
    base = base.sort_values("orders", ascending=False)
    return base


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
    base = base.sort_values("units", ascending=False)
    return base


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
    c.setFont(font_name, font_size)
    prefix = "• " if bullet else ""
    lines = simpleSplit(prefix + text, font_name, font_size, max_width)
    for line in lines:
        c.drawString(x, y, line)
        y -= line_height
    return y


def make_pdf(summary_dict, insights, recommendations):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # =========================
    # PAGE 1 - PORTADA / RESUMEN
    # =========================
    c.setFillColor(HexColor("#0B1F3A"))
    c.rect(0, height - 110, width, 110, fill=1, stroke=0)

    c.setFillColor(HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 22)
    c.drawString(40, height - 50, "Imporey Internacional")

    c.setFont("Helvetica", 12)
    c.drawString(40, height - 72, "Dashboard Operativo y Logístico")
    c.drawString(40, height - 90, "Resumen ejecutivo generado automáticamente")

    c.setFillColor(HexColor("#6B7280"))
    c.setFont("Helvetica", 9)
    c.drawRightString(width - 40, height - 90, "Reporte corporativo")

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

    draw_section_title(c, 40, y, "Resumen ejecutivo")
    y -= 25

    executive_text = (
        "Este reporte presenta una vista ejecutiva del desempeño operativo y logístico "
        "a partir de la información cargada en el archivo Excel. El objetivo es facilitar "
        "la identificación de volumen operativo, cumplimiento, concentración de carga y "
        "áreas de oportunidad para la toma de decisiones."
    )
    c.setFillColor(HexColor("#374151"))
    y = add_wrapped_text(c, executive_text, 40, y, 520, font_size=10, line_height=14)

    y -= 10
    draw_section_title(c, 40, y, "Hallazgos clave")
    y -= 22

    c.setFillColor(HexColor("#111827"))
    for ins in insights[:5]:
        y = add_wrapped_text(c, ins, 50, y, 500, bullet=True)
        y -= 3
        if y < 70:
            c.showPage()
            y = height - 50

    # =========================
    # PAGE 2 - RECOMENDACIONES
    # =========================
    c.showPage()

    c.setFillColor(HexColor("#0B1F3A"))
    c.rect(0, height - 90, width, 90, fill=1, stroke=0)

    c.setFillColor(HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 18)
    c.drawString(40, height - 50, "Imporey Internacional")

    c.setFont("Helvetica", 11)
    c.drawString(40, height - 70, "Recomendaciones ejecutivas")

    y = height - 120
    draw_section_title(c, 40, y, "Recomendaciones prioritarias")
    y -= 25

    c.setFillColor(HexColor("#111827"))
    for rec in recommendations:
        y = add_wrapped_text(c, rec, 50, y, 500, bullet=True)
        y -= 4
        if y < 80:
            c.showPage()
            y = height - 50

    y -= 10
    if y > 120:
        draw_section_title(c, 40, y, "Conclusión")
        y -= 25

        conclusion = (
            "En conjunto, el análisis permite identificar los principales focos de volumen, "
            "cumplimiento y presión operativa. La utilidad del dashboard radica en convertir "
            "información transaccional en señales ejecutivas que faciliten priorización, seguimiento "
            "y mejora continua en la operación."
        )
        c.setFillColor(HexColor("#374151"))
        y = add_wrapped_text(c, conclusion, 40, y, 520, font_size=10, line_height=14)

    c.setFillColor(HexColor("#9CA3AF"))
    c.setFont("Helvetica", 8)
    c.drawRightString(570, 20, "Imporey Internacional | Reporte generado automáticamente")

    c.save()
    buffer.seek(0)
    return buffer


def format_table_for_display(df, hour_cols=None, pct_cols=None, int_cols=None):
    out = df.copy()
    hour_cols = hour_cols or []
    pct_cols = pct_cols or []
    int_cols = int_cols or []

    for c in hour_cols:
        if c in out.columns:
            out[c] = out[c].apply(fmt_hours)

    for c in pct_cols:
        if c in out.columns:
            out[c] = out[c].apply(lambda x: f"{x:.1f}%")

    for c in int_cols:
        if c in out.columns:
            out[c] = out[c].apply(fmt_int)

    return out


# =========================================================
# UPLOAD
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

        # =========================================================
        # FILTERS
        # =========================================================
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

        # =========================================================
        # KPIS
        # =========================================================
        total_rows = len(filtered)
        total_orders = safe_nunique(filtered["order_id"])
        total_shipments = safe_nunique(filtered["shipment_id"])
        total_units = filtered["qty"].sum()
        delivered_rate = filtered["is_delivered"].mean() * 100 if len(filtered) else 0
        picked_rate = filtered["is_picked"].mean() * 100 if len(filtered) else 0
        cancel_rate = filtered["is_cancelled"].mean() * 100 if len(filtered) else 0
        avg_cycle_time = filtered["cycle_time_hours"].mean()
        avg_time_spent = filtered["time_spent_hours"].mean()

        st.subheader("Resumen ejecutivo")
        k1, k2, k3, k4 = st.columns(4)
        k5, k6, k7, k8 = st.columns(4)

        k1.metric("Registros", fmt_int(total_rows))
        k2.metric("Pedidos únicos", fmt_int(total_orders))
        k3.metric("Envíos únicos", fmt_int(total_shipments))
        k4.metric("Unidades", fmt_int(total_units))
        k5.metric("% Entregado", fmt_pct(delivered_rate))
        k6.metric("% Pickeado", fmt_pct(picked_rate))
        k7.metric("% Cancelado", fmt_pct(cancel_rate))
        k8.metric("Tiempo ciclo promedio", fmt_hours(avg_cycle_time))

        st.caption(f"Tiempo dedicado promedio: {fmt_hours(avg_time_spent)}")

        # =========================================================
        # BUILD ANALYSIS
        # =========================================================
        daily_ops = build_daily_operations(filtered)
        warehouse_perf = build_warehouse_performance(filtered)
        carrier_perf = build_carrier_performance(filtered)
        channel_perf = build_channel_performance(filtered)
        product_perf = build_product_performance(filtered)

        insights = generate_insights(
            filtered, daily_ops, warehouse_perf, carrier_perf, channel_perf, product_perf
        )

        recommendations = generate_recommendations(
            filtered, warehouse_perf, carrier_perf, channel_perf, product_perf
        )

        # =========================================================
        # CHARTS
        # =========================================================
        st.subheader("Tendencia operativa diaria")

        if not daily_ops.empty:
            fig1, ax1 = plt.subplots(figsize=(10, 4))
            ax1.plot(daily_ops["analysis_date"], daily_ops["orders"], marker="o")
            ax1.set_title("Pedidos por día")
            ax1.set_xlabel("Fecha")
            ax1.set_ylabel("Pedidos")
            plt.xticks(rotation=45)
            st.pyplot(fig1)

            fig2, ax2 = plt.subplots(figsize=(10, 4))
            ax2.plot(daily_ops["analysis_date"], daily_ops["delivered"], label="Entregados")
            ax2.plot(daily_ops["analysis_date"], daily_ops["cancelled"], label="Cancelados")
            ax2.set_title("Entregados vs cancelados por día")
            ax2.set_xlabel("Fecha")
            ax2.set_ylabel("Registros")
            ax2.legend()
            plt.xticks(rotation=45)
            st.pyplot(fig2)

        st.subheader("Desempeño por almacén")
        if not warehouse_perf.empty:
            fig3, ax3 = plt.subplots(figsize=(8, 4))
            top_wh = warehouse_perf.head(10).sort_values("orders", ascending=True)
            ax3.barh(top_wh["warehouse"], top_wh["orders"])
            ax3.set_title("Pedidos por almacén")
            ax3.set_xlabel("Pedidos")
            st.pyplot(fig3)

            st.dataframe(
                format_table_for_display(
                    warehouse_perf,
                    hour_cols=["avg_cycle_hours", "avg_time_spent"],
                    pct_cols=["delivered_rate", "cancelled_rate", "share_orders"],
                    int_cols=["orders", "shipments", "units"]
                ),
                use_container_width=True
            )

        st.subheader("Desempeño por carrier")
        if not carrier_perf.empty:
            fig4, ax4 = plt.subplots(figsize=(8, 5))
            top_car = carrier_perf.head(10).sort_values("shipments", ascending=True)
            ax4.barh(top_car["carrier"], top_car["shipments"])
            ax4.set_title("Envíos por carrier")
            ax4.set_xlabel("Envíos")
            st.pyplot(fig4)

            st.dataframe(
                format_table_for_display(
                    carrier_perf,
                    hour_cols=["avg_cycle_hours", "avg_time_spent"],
                    pct_cols=["delivered_rate", "cancelled_rate", "share_shipments"],
                    int_cols=["shipments", "orders", "units"]
                ),
                use_container_width=True
            )

        st.subheader("Desempeño por canal")
        if not channel_perf.empty:
            fig5, ax5 = plt.subplots(figsize=(8, 5))
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

        st.subheader("Top productos")
        if not product_perf.empty:
            fig6, ax6 = plt.subplots(figsize=(10, 6))
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

        # =========================================================
        # INSIGHTS
        # =========================================================
        st.subheader("Hallazgos automáticos")
        for i, insight in enumerate(insights, start=1):
            st.write(f"{i}. {insight}")

        st.subheader("Recomendaciones ejecutivas")
        for i, rec in enumerate(recommendations, start=1):
            st.write(f"{i}. {rec}")

        # =========================================================
        # RAW DATA
        # =========================================================
        with st.expander("Vista previa de datos procesados"):
            st.dataframe(filtered.head(100), use_container_width=True)

        # =========================================================
        # PDF
        # =========================================================
        summary_dict = {
            "Registros": fmt_int(total_rows),
            "Pedidos únicos": fmt_int(total_orders),
            "Envíos únicos": fmt_int(total_shipments),
            "Unidades": fmt_int(total_units),
            "% Entregado": fmt_pct(delivered_rate),
            "% Pickeado": fmt_pct(picked_rate),
            "% Cancelado": fmt_pct(cancel_rate),
            "Tiempo ciclo prom.": fmt_hours(avg_cycle_time),
            "Tiempo dedicado prom.": fmt_hours(avg_time_spent),
        }

        pdf_buffer = make_pdf(summary_dict, insights, recommendations)

        st.download_button(
            label="Descargar PDF ejecutivo",
            data=pdf_buffer,
            file_name="dashboard_operativo_imporey.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")

else:
    st.info("Sube un archivo Excel para comenzar.")

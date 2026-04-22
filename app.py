import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from reportlab.lib.colors import HexColor

st.set_page_config(page_title="Imporey Internacional | Dashboard Operativo", layout="wide")

# =========================================================
# CONFIG
# =========================================================
st.title("Imporey Internacional | Dashboard Operativo")
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

    df["is_picked"] = np.where(df["picked_flag"] > 0, 1, 0)

    df["is_cancelled"] = np.where(
        df["status"].str.lower().str.contains("cancel", na=False),
        1, 0
    )

    df["is_collected"] = np.where(
        df["status"].str.lower().str.contains("recolect", na=False),
        1, 0
    )

    df["analysis_date"] = df["created_on"].dt.date
    df["analysis_datetime"] = df["created_on"]
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

    df["cycle_time_hours"] = np.where(df["cycle_time_hours"] < 0, np.nan, df["cycle_time_hours"])
    df["pick_time_hours"] = np.where(df["pick_time_hours"] < 0, np.nan, df["pick_time_hours"])
    df["pack_time_hours"] = np.where(df["pack_time_hours"] < 0, np.nan, df["pack_time_hours"])

    df["sla_bucket"] = df["cycle_time_hours"].apply(classify_sla_bucket)

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


def get_status_color(score):
    if pd.isna(score):
        return "⚪ Sin dato"
    if score >= 85:
        return "🟢 Alto"
    if score >= 70:
        return "🟡 Medio"
    return "🔴 Riesgo"


def calc_score(df_base, volume_col):
    df = df_base.copy()

    # Normalizaciones
    max_volume = df[volume_col].max() if len(df) else 0
    if max_volume > 0:
        df["volume_score"] = (df[volume_col] / max_volume) * 100
    else:
        df["volume_score"] = 0

    df["delivery_score"] = df["delivered_rate"].clip(lower=0, upper=100)
    df["cancel_score"] = (100 - df["cancelled_rate"]).clip(lower=0, upper=100)

    valid_cycle = df["avg_cycle_hours"].replace([np.inf, -np.inf], np.nan)
    max_cycle = valid_cycle.max()

    if pd.notna(max_cycle) and max_cycle > 0:
        df["cycle_score"] = (100 - ((valid_cycle / max_cycle) * 100)).clip(lower=0, upper=100)
    else:
        df["cycle_score"] = 100

    df["performance_score"] = (
        df["delivery_score"] * 0.40
        + df["cycle_score"] * 0.30
        + df["cancel_score"] * 0.20
        + df["volume_score"] * 0.10
    ).round(1)

    df["status_light"] = df["performance_score"].apply(get_status_color)
    return df


def compare_periods(df):
    tmp = df.dropna(subset=["analysis_datetime"]).copy()
    if tmp.empty:
        return None

    min_dt = tmp["analysis_datetime"].min()
    max_dt = tmp["analysis_datetime"].max()

    if pd.isna(min_dt) or pd.isna(max_dt):
        return None

    total_days = max((max_dt.date() - min_dt.date()).days + 1, 1)
    current_start = max_dt.normalize() - pd.Timedelta(days=total_days - 1)
    current_end = max_dt.normalize() + pd.Timedelta(days=1)

    prev_end = current_start
    prev_start = prev_end - pd.Timedelta(days=total_days)

    current_df = tmp[(tmp["analysis_datetime"] >= current_start) & (tmp["analysis_datetime"] < current_end)].copy()
    previous_df = tmp[(tmp["analysis_datetime"] >= prev_start) & (tmp["analysis_datetime"] < prev_end)].copy()

    def summarize(block):
        if block.empty:
            return {
                "orders": 0,
                "shipments": 0,
                "units": 0,
                "delivered_rate": 0,
                "cancel_rate": 0,
            }
        return {
            "orders": safe_nunique(block["order_id"]),
            "shipments": safe_nunique(block["shipment_id"]),
            "units": float(block["qty"].sum()),
            "delivered_rate": float(block["is_delivered"].mean() * 100),
            "cancel_rate": float(block["is_cancelled"].mean() * 100),
        }

    current_summary = summarize(current_df)
    previous_summary = summarize(previous_df)

    comp = []
    for metric in ["orders", "shipments", "units", "delivered_rate", "cancel_rate"]:
        current_val = current_summary[metric]
        prev_val = previous_summary[metric]
        delta_abs = current_val - prev_val
        delta_pct = ((delta_abs / prev_val) * 100) if prev_val not in [0, 0.0] else np.nan

        comp.append({
            "metric": metric,
            "current": current_val,
            "previous": prev_val,
            "delta_abs": delta_abs,
            "delta_pct": delta_pct
        })

    comparison_df = pd.DataFrame(comp)

    return {
        "days": total_days,
        "current_range": f"{current_start.date()} a {max_dt.date()}",
        "previous_range": f"{prev_start.date()} a {(prev_end - pd.Timedelta(days=1)).date()}",
        "table": comparison_df
    }


def generate_insights(df, daily_ops, warehouse_perf, carrier_perf, channel_perf, product_perf, period_comp):
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
            worst_wh = valid_wh.sort_values("performance_score").iloc[0]
            insights.append(
                f"El almacén con mayor riesgo operativo es {worst_wh['warehouse']} con score {worst_wh['performance_score']:.1f}."
            )

    if not carrier_perf.empty:
        top_carrier = carrier_perf.iloc[0]
        insights.append(
            f"El carrier principal es {top_carrier['carrier']} con {fmt_int(top_carrier['shipments'])} envíos "
            f"({top_carrier['share_shipments']:.1f}% del total)."
        )

        valid_carrier = carrier_perf[carrier_perf["shipments"] > 0].copy()
        if not valid_carrier.empty:
            worst_carrier = valid_carrier.sort_values("performance_score").iloc[0]
            insights.append(
                f"El carrier con mayor foco de atención es {worst_carrier['carrier']} con score {worst_carrier['performance_score']:.1f}."
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
                "En los últimos días el volumen operativo está por encima del promedio general; conviene revisar capacidad operativa y distribución de carga."
            )
        elif recent_avg < global_avg * 0.85:
            insights.append(
                "En los últimos días el volumen operativo está por debajo del promedio general, lo que puede abrir una ventana para ajustes operativos."
            )

    if cancel_rate > 3:
        insights.append(
            f"La tasa de cancelación está en {cancel_rate:.1f}%, lo que amerita revisar causas raíz por canal, carrier y disponibilidad."
        )

    if delivered_rate < 90:
        insights.append(
            f"El porcentaje entregado es {delivered_rate:.1f}%, por debajo de un nivel operativo fuerte; conviene revisar SLA por carrier y almacén."
        )

    sla_counts = df["sla_bucket"].value_counts(dropna=False)
    total_sla = sla_counts.sum() if len(sla_counts) else 0
    if total_sla > 0:
        gt48 = (sla_counts.get(">48 h", 0) / total_sla) * 100
        if gt48 > 25:
            insights.append(
                f"El {gt48:.1f}% de los registros con dato de ciclo cae en SLA mayor a 48 horas, lo que sugiere una oportunidad clara de reducción de tiempo."
            )

    if period_comp is not None and not period_comp["table"].empty:
        orders_row = period_comp["table"][period_comp["table"]["metric"] == "orders"]
        if not orders_row.empty:
            delta = orders_row.iloc[0]["delta_pct"]
            if pd.notna(delta):
                if delta > 10:
                    insights.append(
                        f"En el periodo actual los pedidos crecieron {delta:.1f}% vs el periodo anterior comparable."
                    )
                elif delta < -10:
                    insights.append(
                        f"En el periodo actual los pedidos cayeron {abs(delta):.1f}% vs el periodo anterior comparable."
                    )

    if total_orders > 0 and total_shipments > 0 and total_shipments < total_orders * 0.8:
        insights.append(
            "Hay una diferencia relevante entre pedidos y envíos únicos; puede valer la pena revisar consolidación, órdenes incompletas o estatus pendientes."
        )

    return insights


def generate_recommendations(df, warehouse_perf, carrier_perf, channel_perf, product_perf):
    recommendations = []

    if not warehouse_perf.empty:
        risk_wh = warehouse_perf.sort_values("performance_score").iloc[0]
        recommendations.append(
            f"Priorizar seguimiento del almacén {risk_wh['warehouse']}, ya que combina nivel de volumen con el menor score de desempeño dentro del análisis."
        )

    if not carrier_perf.empty:
        risk_carrier = carrier_perf.sort_values("performance_score").iloc[0]
        recommendations.append(
            f"Revisar el desempeño del carrier {risk_carrier['carrier']}, ya que presenta el menor score operativo y puede impactar el cumplimiento final."
        )

    if not channel_perf.empty:
        top_channel = channel_perf.iloc[0]
        recommendations.append(
            f"Dar seguimiento especial al canal {top_channel['channel']}, ya que concentra el mayor volumen de pedidos y define gran parte de la presión diaria."
        )

    if not product_perf.empty:
        top_product = product_perf.iloc[0]
        recommendations.append(
            f"Considerar una gestión prioritaria del producto {top_product['product']}, al ser el de mayor movimiento y carga operativa en el periodo."
        )

    cancel_rate = df["is_cancelled"].mean() * 100 if len(df) else 0
    if cancel_rate > 2:
        recommendations.append(
            "Implementar una revisión semanal de cancelaciones por origen, carrier y canal para reducir fricción operativa y evitar impacto al cliente."
        )

    avg_cycle_time = df["cycle_time_hours"].mean()
    if pd.notna(avg_cycle_time) and avg_cycle_time > 48:
        recommendations.append(
            "Evaluar oportunidades de mejora en tiempos de ciclo, especialmente en pedidos con mayor permanencia entre creación y finalización."
        )

    sla_gt48 = (df["sla_bucket"] == ">48 h").mean() * 100 if len(df) else 0
    if sla_gt48 > 20:
        recommendations.append(
            "Conviene establecer seguimiento puntual al segmento de pedidos con ciclo mayor a 48 horas para reducir incumplimiento operativo."
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
    base = calc_score(base, "orders")
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
    base = calc_score(base, "shipments")
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


def build_sla_summary(df):
    tmp = df.copy()
    tmp = tmp[tmp["sla_bucket"] != "Sin dato"]

    if tmp.empty:
        return pd.DataFrame(columns=["sla_bucket", "records", "share"])

    base = (
        tmp.groupby("sla_bucket")
        .size()
        .reset_index(name="records")
    )
    total = base["records"].sum()
    base["share"] = np.where(total > 0, base["records"] / total * 100, 0)

    order = {"<24 h": 1, "24-48 h": 2, ">48 h": 3}
    base["order"] = base["sla_bucket"].map(order)
    base = base.sort_values("order").drop(columns=["order"])

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


def make_pdf(summary_dict, insights, recommendations, period_comp, sla_summary):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # PAGE 1
    c.setFillColor(HexColor("#0B1F3A"))
    c.rect(0, height - 110, width, 110, fill=1, stroke=0)

    c.setFillColor(HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 22)
    c.drawString(40, height - 50, "Imporey Internacional")
    c.setFont("Helvetica", 12)
    c.drawString(40, height - 72, "Dashboard Operativo y Logístico")
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
        y -= 3

    # PAGE 2
    c.showPage()
    c.setFillColor(HexColor("#0B1F3A"))
    c.rect(0, height - 90, width, 90, fill=1, stroke=0)

    c.setFillColor(HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 18)
    c.drawString(40, height - 50, "Imporey Internacional")
    c.setFont("Helvetica", 11)
    c.drawString(40, height - 70, "Comparativos, SLA y recomendaciones")

    y = height - 120

    if period_comp is not None:
        draw_section_title(c, 40, y, "Comparativo de periodos")
        y -= 22
        c.setFillColor(HexColor("#374151"))
        c.setFont("Helvetica", 10)
        c.drawString(40, y, f"Periodo actual: {period_comp['current_range']}")
        y -= 14
        c.drawString(40, y, f"Periodo anterior: {period_comp['previous_range']}")
        y -= 22

        for _, row in period_comp["table"].iterrows():
            metric_name = {
                "orders": "Pedidos",
                "shipments": "Envíos",
                "units": "Unidades",
                "delivered_rate": "% Entregado",
                "cancel_rate": "% Cancelado"
            }.get(row["metric"], row["metric"])

            current_txt = fmt_pct(row["current"]) if "rate" in row["metric"] else fmt_int(row["current"])
            prev_txt = fmt_pct(row["previous"]) if "rate" in row["metric"] else fmt_int(row["previous"])

            if pd.notna(row["delta_pct"]):
                delta_txt = f"{row['delta_pct']:+.1f}%"
            else:
                delta_txt = "N/A"

            c.drawString(50, y, f"{metric_name}: actual {current_txt} | anterior {prev_txt} | variación {delta_txt}")
            y -= 14

        y -= 10

    if sla_summary is not None and not sla_summary.empty:
        draw_section_title(c, 40, y, "Distribución SLA")
        y -= 22
        for _, row in sla_summary.iterrows():
            c.drawString(50, y, f"{row['sla_bucket']}: {fmt_int(row['records'])} registros ({row['share']:.1f}%)")
            y -= 14
        y -= 10

    draw_section_title(c, 40, y, "Recomendaciones prioritarias")
    y -= 22
    for rec in recommendations[:6]:
        y = add_wrapped_text(c, rec, 50, y, 500, bullet=True)
        y -= 3

    c.setFillColor(HexColor("#9CA3AF"))
    c.setFont("Helvetica", 8)
    c.drawRightString(570, 20, "Imporey Internacional | Reporte generado automáticamente")

    c.save()
    buffer.seek(0)
    return buffer


def format_table_for_display(df, hour_cols=None, pct_cols=None, int_cols=None, score_cols=None):
    out = df.copy()
    hour_cols = hour_cols or []
    pct_cols = pct_cols or []
    int_cols = int_cols or []
    score_cols = score_cols or []

    for c in hour_cols:
        if c in out.columns:
            out[c] = out[c].apply(fmt_hours)

    for c in pct_cols:
        if c in out.columns:
            out[c] = out[c].apply(lambda x: f"{x:.1f}%")

    for c in int_cols:
        if c in out.columns:
            out[c] = out[c].apply(fmt_int)

    for c in score_cols:
        if c in out.columns:
            out[c] = out[c].apply(lambda x: f"{x:.1f}")

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

        # FILTERS
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

        # KPIS
        total_rows = len(filtered)
        total_orders = safe_nunique(filtered["order_id"])
        total_shipments = safe_nunique(filtered["shipment_id"])
        total_units = filtered["qty"].sum()
        delivered_rate = filtered["is_delivered"].mean() * 100 if len(filtered) else 0
        picked_rate = filtered["is_picked"].mean() * 100 if len(filtered) else 0
        cancel_rate = filtered["is_cancelled"].mean() * 100 if len(filtered) else 0
        avg_cycle_time = filtered["cycle_time_hours"].mean()
        avg_time_spent = filtered["time_spent_hours"].mean()

        sla_summary = build_sla_summary(filtered)
        pct_lt24 = 0
        if not sla_summary.empty:
            row_lt24 = sla_summary[sla_summary["sla_bucket"] == "<24 h"]
            if not row_lt24.empty:
                pct_lt24 = float(row_lt24.iloc[0]["share"])

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
        k8.metric("Tiempo ciclo prom.", fmt_hours(avg_cycle_time))

        st.caption(f"Tiempo dedicado promedio: {fmt_hours(avg_time_spent)} | SLA <24h: {fmt_pct(pct_lt24)}")

        # BUILD ANALYSIS
        daily_ops = build_daily_operations(filtered)
        warehouse_perf = build_warehouse_performance(filtered)
        carrier_perf = build_carrier_performance(filtered)
        channel_perf = build_channel_performance(filtered)
        product_perf = build_product_performance(filtered)
        period_comp = compare_periods(filtered)

        insights = generate_insights(
            filtered, daily_ops, warehouse_perf, carrier_perf, channel_perf, product_perf, period_comp
        )

        recommendations = generate_recommendations(
            filtered, warehouse_perf, carrier_perf, channel_perf, product_perf
        )

        # COMPARATIVO
        st.subheader("Comparativo de periodos")
        if period_comp is not None:
            st.write(f"**Periodo actual:** {period_comp['current_range']}")
            st.write(f"**Periodo anterior comparable:** {period_comp['previous_range']}")

            comp_show = period_comp["table"].copy()
            comp_show["Métrica"] = comp_show["metric"].map({
                "orders": "Pedidos",
                "shipments": "Envíos",
                "units": "Unidades",
                "delivered_rate": "% Entregado",
                "cancel_rate": "% Cancelado"
            })
            comp_show["Actual"] = comp_show.apply(
                lambda r: fmt_pct(r["current"]) if "rate" in r["metric"] else fmt_int(r["current"]), axis=1
            )
            comp_show["Anterior"] = comp_show.apply(
                lambda r: fmt_pct(r["previous"]) if "rate" in r["metric"] else fmt_int(r["previous"]), axis=1
            )
            comp_show["Variación"] = comp_show["delta_pct"].apply(
                lambda x: "N/A" if pd.isna(x) else f"{x:+.1f}%"
            )
            st.dataframe(comp_show[["Métrica", "Actual", "Anterior", "Variación"]], use_container_width=True)
        else:
            st.info("No hay suficiente información de fechas para construir comparativo de periodos.")

        # CHARTS
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

        # SLA
        st.subheader("Distribución SLA")
        if not sla_summary.empty:
            col1, col2 = st.columns([1, 1])

            with col1:
                st.dataframe(
                    format_table_for_display(
                        sla_summary,
                        pct_cols=["share"],
                        int_cols=["records"]
                    ),
                    use_container_width=True
                )

            with col2:
                fig_sla, ax_sla = plt.subplots(figsize=(7, 4))
                ax_sla.bar(sla_summary["sla_bucket"], sla_summary["records"])
                ax_sla.set_title("Registros por rango SLA")
                ax_sla.set_xlabel("SLA")
                ax_sla.set_ylabel("Registros")
                st.pyplot(fig_sla)
        else:
            st.info("No hay datos suficientes para construir SLA.")

        # SCORE ALMACEN
        st.subheader("Desempeño por almacén")
        if not warehouse_perf.empty:
            top_wh = warehouse_perf.head(10).sort_values("orders", ascending=True)

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

        # SCORE CARRIER
        st.subheader("Desempeño por carrier")
        if not carrier_perf.empty:
            top_car = carrier_perf.head(10).sort_values("shipments", ascending=True)

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

        # CANAL
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

        # PRODUCTOS
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

        # HALLAZGOS
        st.subheader("Hallazgos automáticos")
        for i, insight in enumerate(insights, start=1):
            st.write(f"{i}. {insight}")

        st.subheader("Recomendaciones ejecutivas")
        for i, rec in enumerate(recommendations, start=1):
            st.write(f"{i}. {rec}")

        # RAW DATA
        with st.expander("Vista previa de datos procesados"):
            st.dataframe(filtered.head(100), use_container_width=True)

        # PDF
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

        pdf_buffer = make_pdf(summary_dict, insights, recommendations, period_comp, sla_summary)

        st.download_button(
            label="Descargar PDF ejecutivo",
            data=pdf_buffer,
            file_name="dashboard_operativo_imporey_v2.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")

else:
    st.info("Sube un archivo Excel para comenzar.")

import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io

st.title("E-commerce Analyzer")

# Subir archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Vista previa")
    st.dataframe(df)

    # === ANALISIS SIMPLE ===
    total_sales = df["revenue"].sum()
    total_units = df["units_sold"].sum()
    top_product = df.groupby("product")["units_sold"].sum().idxmax()

    prediction = total_sales * 1.1

    st.subheader("Resultados")

    st.write("Ventas del mes:", total_sales)
    st.write("Unidades vendidas:", total_units)
    st.write("Producto top:", top_product)
    st.write("Predicción siguiente mes:", prediction)

    # === GENERAR PDF ===
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)

    pdf.drawString(50, 750, "Reporte E-commerce")
    pdf.drawString(50, 720, f"Ventas: {total_sales}")
    pdf.drawString(50, 700, f"Unidades: {total_units}")
    pdf.drawString(50, 680, f"Producto top: {top_product}")
    pdf.drawString(50, 660, f"Predicción: {prediction}")

    pdf.save()
    buffer.seek(0)

    st.download_button(
        label="Descargar PDF",
        data=buffer,
        file_name="reporte.pdf",
        mime="application/pdf"
    )

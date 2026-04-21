import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

st.set_page_config(page_title="E-Commerce Analyzer", layout="wide")

st.title("E-Commerce Analyzer")
st.write("Sube tu archivo Excel")

uploaded_file = st.file_uploader("Upload", type=["xlsx"])

def create_pdf_report(total_sales, total_units, top_product, prediction):
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    pdf.setTitle("Reporte E-Commerce")

    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(50, 750, "Reporte E-Commerce")

    pdf.setFont("Helvetica", 12)
    pdf.drawString(50, 710, f"Ventas totales: ${total_sales:,.2f}")
    pdf.drawString(50, 690, f"Unidades vendidas: {int(total_units)}")
    pdf.drawString(50, 670, f"Producto top: {top_product}")
    pdf.drawString(50, 650, f"Predicción siguiente periodo: ${prediction:,.2f}")

    pdf.save()
    buffer.seek(0)
    return buffer

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    st.subheader("Vista previa del archivo")
    st.dataframe(df)

    # Normalizar nombres de columnas
    df.columns = df.columns.str.strip().str.lower()

    # Renombrar columnas comunes automáticamente
    rename_map = {}

    if "fecha" in df.columns:
        rename_map["fecha"] = "date"
    if "producto" in df.columns:
        rename_map["producto"] = "product"
    if "ventas" in df.columns:
        rename_map["ventas"] = "sales"
    if "ingresos" in df.columns:
        rename_map["ingresos"] = "sales"
    if "unidades" in df.columns:
        rename_map["unidades"] = "units_sold"
    if "cantidad" in df.columns:
        rename_map["cantidad"] = "units_sold"

    df = df.rename(columns=rename_map)

    required_columns = ["date", "product", "sales", "units_sold"]

    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        st.error(f"Faltan estas columnas en el archivo: {missing}")
        st.info("Tu Excel debe tener estas columnas: date, product, sales, units_sold")
    else:
        total_sales = df["sales"].sum()
        total_units = df["units_sold"].sum()
        top_product = df.groupby("product")["sales"].sum().idxmax()
        prediction = total_sales * 1.10  # predicción simple +10%

        st.subheader("Resultados")
        st.metric("Ventas totales", f"${total_sales:,.2f}")
        st.metric("Unidades vendidas", int(total_units))
        st.metric("Producto top", top_product)
        st.metric("Predicción siguiente periodo", f"${prediction:,.2f}")

        pdf_file = create_pdf_report(total_sales, total_units, top_product, prediction)

        st.download_button(
            label="Descargar reporte PDF",
            data=pdf_file,
            file_name="reporte_ecommerce.pdf",
            mime="application/pdf"
        )

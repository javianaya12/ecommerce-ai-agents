import pandas as pd

def load_sales_data():
    df = pd.read_csv("data/sales.csv")

    ventas_mes = (df["cantidad"] * df["precio"]).sum()
    productos_vendidos = df["cantidad"].sum()
    producto_top = df.groupby("producto")["cantidad"].sum().idxmax()

    return {
        "ventas_mes": ventas_mes,
        "productos_vendidos": productos_vendidos,
        "producto_top": producto_top
    }

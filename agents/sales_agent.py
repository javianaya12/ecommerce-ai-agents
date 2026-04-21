def analyze_sales(data):
    ventas_mes = data["ventas_mes"]
    productos_vendidos = data["productos_vendidos"]
    producto_top = data["producto_top"]

    if ventas_mes > 100000:
        tendencia = "Ventas fuertes"
    else:
        tendencia = "Ventas moderadas"

    prediccion = ventas_mes * 1.10

    return f"""
=== SALES ANALYSIS ===
Ventas del mes: {ventas_mes}
Productos vendidos: {productos_vendidos}
Producto top: {producto_top}
Tendencia: {tendencia}
Predicción siguiente mes: {prediccion:.2f}
Recomendación: Mantener impulso comercial y revisar inventario del producto top.
"""

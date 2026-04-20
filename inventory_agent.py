def inventory_agent(data):
    stock = data["stock_actual"]
    producto_top = data["producto_top"]
    productos_vendidos = data["productos_vendidos"]

    if productos_vendidos <= 0:
        return "No hay ventas suficientes para estimar inventario."

    venta_promedio_simple = productos_vendidos / 30
    dias_cobertura = stock / venta_promedio_simple if venta_promedio_simple > 0 else 0

    if stock < 50:
        riesgo = "ALTO riesgo de desabasto"
        recomendacion = "Reordenar cuanto antes."
    elif stock < 100:
        riesgo = "Riesgo medio"
        recomendacion = "Monitorear rotación y preparar reabasto."
    else:
        riesgo = "Stock saludable"
        recomendacion = "Mantener nivel actual y revisar exceso de inventario."

    return f"""Producto top: {producto_top}
Stock actual: {stock}
Venta promedio diaria estimada: {venta_promedio_simple:.2f}
Días estimados de cobertura: {dias_cobertura:.2f}
Estado: {riesgo}
Recomendación: {recomendacion}"""

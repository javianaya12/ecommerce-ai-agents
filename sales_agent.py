def sales_agent(data):
    ventas = data["ventas_mes"]
    productos = data["productos_vendidos"]
    producto_top = data["producto_top"]

    if productos <= 0:
        return "No hay productos vendidos para analizar."

    ticket_promedio = ventas / productos

    if ventas >= 100000:
        tendencia = "Ventas fuertes"
        recomendacion = "Mantener impulso comercial y revisar inventario del producto top."
    elif ventas >= 50000:
        tendencia = "Ventas medias"
        recomendacion = "Empujar promociones en categorías con mejor margen."
    else:
        tendencia = "Ventas bajas"
        recomendacion = "Revisar campañas, conversión y productos con baja rotación."

    prediccion_siguiente_mes = ventas * 1.10

    return f"""Ventas del mes: {ventas}
Productos vendidos: {productos}
Producto top: {producto_top}
Ticket promedio: {ticket_promedio:.2f}
Tendencia: {tendencia}
Predicción siguiente mes: {prediccion_siguiente_mes:.2f}
Recomendación: {recomendacion}"""

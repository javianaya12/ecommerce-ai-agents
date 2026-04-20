def logistics_agent(data):
    pedidos_totales = data["pedidos_totales"]
    pedidos_retrasados = data["pedidos_retrasados"]
    warehouse_top_delay = data["warehouse_top_delay"]
    carrier_top_delay = data["carrier_top_delay"]

    if pedidos_totales <= 0:
        return "No hay pedidos para analizar logística."

    tasa_retraso = (pedidos_retrasados / pedidos_totales) * 100

    if tasa_retraso >= 20:
        estado = "Nivel de retraso alto"
        recomendacion = "Revisar operación del almacén y carrier con más incidencias."
    elif tasa_retraso >= 10:
        estado = "Nivel de retraso medio"
        recomendacion = "Monitorear saturación operativa y tiempos de corte."
    else:
        estado = "Nivel de retraso controlado"
        recomendacion = "Mantener seguimiento y prevención."

    return f"""Pedidos totales: {pedidos_totales}
Pedidos retrasados: {pedidos_retrasados}
Tasa de retraso: {tasa_retraso:.2f}%
Almacén con más retrasos: {warehouse_top_delay}
Carrier con más retrasos: {carrier_top_delay}
Estado: {estado}
Recomendación: {recomendacion}"""

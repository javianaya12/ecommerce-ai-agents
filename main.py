from agents.sales_agent import sales_agent
from agents.inventory_agent import inventory_agent
from agents.logistics_agent import logistics_agent


def run_system(data):
    print("\n=== SALES ANALYSIS ===")
    print(sales_agent(data))

    print("\n=== INVENTORY ANALYSIS ===")
    print(inventory_agent(data))

    print("\n=== LOGISTICS ANALYSIS ===")
    print(logistics_agent(data))


if __name__ == "__main__":
    sample_data = {
        "ventas_mes": 120000,
        "productos_vendidos": 350,
        "stock_actual": 80,
        "producto_top": "Tenis Nike",
        "pedidos_totales": 950,
        "pedidos_retrasados": 140,
        "warehouse_top_delay": "SANTA_ECOM",
        "carrier_top_delay": "FedEx"
    }

    run_system(sample_data)

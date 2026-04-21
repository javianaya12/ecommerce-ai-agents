from agents.sales_agent import analyze_sales
from data.load_data import load_sales_data

def main():
    sales_data = load_sales_data()
    print(analyze_sales(sales_data))

if __name__ == "__main__":
    main()

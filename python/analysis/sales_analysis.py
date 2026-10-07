import pandas as pd

from python.db import engine


def get_sales_data() -> pd.DataFrame:
    query = """
        SELECT
            id,
            train_number,
            departure_date,
            departure_station,
            destination_station,
            sold_tickets,
            ticket_price,
            carriage_type,
            total_revenue
        FROM trains
        ORDER BY departure_date, train_number;
    """

    return pd.read_sql(query, engine)


def get_sales_summary() -> dict:
    df = get_sales_data()

    return {
        "trains": int(df["id"].nunique()),
        "tickets": int(df["sold_tickets"].sum()),
        "revenue": float(df["total_revenue"].sum()),
        "average_ticket_price": float(df["ticket_price"].mean()),
    }


if __name__ == "__main__":
    data = get_sales_data()

    print("\nRailway Analytics")
    print("=" * 50)

    print(f"Количество поездов: {data['id'].nunique()}")
    print(f"Продано билетов: {data['sold_tickets'].sum()}")
    print(f"Общая выручка: {data['total_revenue'].sum():.2f}")
    print(f"Средняя цена билета: {data['ticket_price'].mean():.2f}")

    print("\nДанные:")
    print(data.to_string(index=False))
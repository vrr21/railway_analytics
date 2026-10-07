import pandas as pd

from python.db import engine


def get_route_statistics() -> pd.DataFrame:
    query = """
        SELECT
            departure_station,
            destination_station,
            COUNT(*) AS train_count,
            SUM(sold_tickets) AS total_tickets,
            SUM(total_revenue) AS total_revenue,
            AVG(ticket_price) AS average_ticket_price
        FROM trains
        GROUP BY
            departure_station,
            destination_station
        ORDER BY total_revenue DESC;
    """

    return pd.read_sql(query, engine)


if __name__ == "__main__":
    df = get_route_statistics()

    print("\nRoute statistics")
    print("=" * 70)
    print(df.to_string(index=False))
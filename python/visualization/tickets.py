import matplotlib.pyplot as plt
import pandas as pd

from python.db import engine


def plot_tickets_by_date() -> None:
    query = """
        SELECT
            departure_date,
            SUM(sold_tickets) AS total_tickets
        FROM trains
        GROUP BY departure_date
        ORDER BY departure_date;
    """

    df = pd.read_sql(query, engine)

    plt.figure(figsize=(12, 6))

    plt.plot(
        df["departure_date"],
        df["total_tickets"],
        marker="o"
    )

    plt.title("Продажи билетов по датам")
    plt.xlabel("Дата")
    plt.ylabel("Количество билетов")

    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.show()


if __name__ == "__main__":
    plot_tickets_by_date()
import matplotlib.pyplot as plt
import pandas as pd

from python.db import engine


def plot_revenue_by_train() -> None:
    query = """
        SELECT
            train_number,
            total_revenue
        FROM trains
        ORDER BY total_revenue DESC;
    """

    df = pd.read_sql(query, engine)

    plt.figure(figsize=(12, 6))

    plt.bar(
        df["train_number"].astype(str),
        df["total_revenue"]
    )

    plt.title("Выручка по поездам")
    plt.xlabel("Номер поезда")
    plt.ylabel("Выручка")

    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.show()


if __name__ == "__main__":
    plot_revenue_by_train()
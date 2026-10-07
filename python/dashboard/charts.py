import matplotlib.pyplot as plt
import pandas as pd


def revenue_by_train(df: pd.DataFrame):
    data = (
        df.groupby("train_number", as_index=False)["total_revenue"]
        .sum()
        .sort_values("total_revenue", ascending=False)
        .head(10)
    )

    fig, ax = plt.subplots(figsize=(10, 5))

    ax.bar(
        data["train_number"].astype(str),
        data["total_revenue"]
    )

    ax.set_title("Топ-10 поездов по выручке")
    ax.set_xlabel("Поезд")
    ax.set_ylabel("Выручка")

    plt.xticks(rotation=45)
    plt.tight_layout()

    return fig


def tickets_by_date(df: pd.DataFrame):
    data = (
        df.groupby("departure_date", as_index=False)["sold_tickets"]
        .sum()
        .sort_values("departure_date")
    )

    fig, ax = plt.subplots(figsize=(10, 5))

    ax.plot(
        data["departure_date"],
        data["sold_tickets"],
        marker="o"
    )

    ax.set_title("Продажи билетов по датам")
    ax.set_xlabel("Дата")
    ax.set_ylabel("Билеты")

    plt.xticks(rotation=45)
    plt.tight_layout()

    return fig


def revenue_by_route(df: pd.DataFrame):
    data = (
        df.assign(
            route=df["departure_station"]
            + " → "
            + df["destination_station"]
        )
        .groupby("route", as_index=False)["total_revenue"]
        .sum()
        .sort_values("total_revenue", ascending=False)
        .head(10)
    )

    fig, ax = plt.subplots(figsize=(10, 5))

    ax.barh(
        data["route"],
        data["total_revenue"]
    )

    ax.set_title("Выручка по маршрутам")
    ax.set_xlabel("Выручка")
    ax.set_ylabel("Маршрут")

    plt.tight_layout()

    return fig
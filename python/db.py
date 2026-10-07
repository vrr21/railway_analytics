from sqlalchemy import create_engine, text
import pandas as pd

from python.config import get_database_url


engine = create_engine(
    get_database_url(),
    pool_pre_ping=True,
)


def test_connection() -> None:

    with engine.connect() as connection:

        result = connection.execute(
            text("SELECT 1")
        )

        print(
            f"Database connection successful: "
            f"{result.scalar()}"
        )


def show_tables() -> None:

    query = text(
        """
        SELECT table_name
        FROM information_schema.tables
        WHERE table_schema = 'public'
        ORDER BY table_name;
        """
    )

    with engine.connect() as connection:

        result = connection.execute(query)

        print("\nTables in database:")

        for row in result:
            print(
                f"- {row[0]}"
            )


def show_trains() -> None:

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
        ORDER BY
            departure_date,
            train_number
        LIMIT 10;
    """

    df = pd.read_sql(
        query,
        engine,
    )

    print(
        "\nFirst 10 records from trains:"
    )

    print(
        df.to_string(index=False)
    )


if __name__ == "__main__":

    test_connection()
    show_tables()
    show_trains()
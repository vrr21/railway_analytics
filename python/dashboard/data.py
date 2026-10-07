import pandas as pd
from sqlalchemy import text
from python.db import engine


# ============================================================
# ВСПОМОГАТЕЛЬНОЕ
# ============================================================

def _normalize_dates(df: pd.DataFrame, column: str = "departure_date") -> pd.DataFrame:
    df = df.copy()
    if column in df.columns:
        df[column] = pd.to_datetime(df[column], errors="coerce")
    return df


# ============================================================
# ПОЕЗДА
# ============================================================

def load_trains() -> pd.DataFrame:
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
        ORDER BY id ASC;
    """
    return _normalize_dates(pd.read_sql(query, engine))


def get_train_options() -> pd.DataFrame:
    query = """
        SELECT
            id,
            train_number,
            departure_date,
            departure_station,
            destination_station
        FROM trains
        ORDER BY id ASC;
    """
    return _normalize_dates(pd.read_sql(query, engine))


def create_train(
    train_number: str,
    departure_date,
    departure_station: str,
    destination_station: str,
    sold_tickets: int,
    ticket_price: float,
    carriage_type: str,
) -> None:
    query = text("""
        INSERT INTO trains (
            train_number,
            departure_date,
            departure_station,
            destination_station,
            sold_tickets,
            ticket_price,
            carriage_type
        )
        VALUES (
            :train_number,
            :departure_date,
            :departure_station,
            :destination_station,
            :sold_tickets,
            :ticket_price,
            :carriage_type
        );
    """)

    with engine.begin() as connection:
        connection.execute(query, {
            "train_number": train_number.strip(),
            "departure_date": departure_date,
            "departure_station": departure_station.strip(),
            "destination_station": destination_station.strip(),
            "sold_tickets": int(sold_tickets),
            "ticket_price": float(ticket_price),
            "carriage_type": carriage_type,
        })


def update_train(
    train_id: int,
    train_number: str,
    departure_date,
    departure_station: str,
    destination_station: str,
    sold_tickets: int,
    ticket_price: float,
    carriage_type: str,
) -> None:
    query = text("""
        UPDATE trains
        SET
            train_number = :train_number,
            departure_date = :departure_date,
            departure_station = :departure_station,
            destination_station = :destination_station,
            sold_tickets = :sold_tickets,
            ticket_price = :ticket_price,
            carriage_type = :carriage_type
        WHERE id = :train_id;
    """)

    with engine.begin() as connection:
        connection.execute(query, {
            "train_id": int(train_id),
            "train_number": train_number.strip(),
            "departure_date": departure_date,
            "departure_station": departure_station.strip(),
            "destination_station": destination_station.strip(),
            "sold_tickets": int(sold_tickets),
            "ticket_price": float(ticket_price),
            "carriage_type": carriage_type,
        })


def delete_train(train_id: int) -> None:
    query = text("""
        DELETE FROM trains
        WHERE id = :train_id;
    """)

    with engine.begin() as connection:
        connection.execute(query, {"train_id": int(train_id)})


# ============================================================
# ПАССАЖИРЫ
# ============================================================

def load_passengers() -> pd.DataFrame:
    query = """
        SELECT
            p.id,
            p.full_name,
            p.train_id,
            t.train_number,
            t.departure_date,
            t.departure_station,
            t.destination_station
        FROM passengers p
        JOIN trains t
            ON t.id = p.train_id
        ORDER BY t.departure_date, p.full_name, p.id;
    """
    return _normalize_dates(pd.read_sql(query, engine))


def create_passenger(
    full_name: str,
    train_id: int,
) -> None:
    # Станции намеренно НЕ передаются из Python.
    # PostgreSQL сам получает маршрут из выбранного train_id.
    query = text("""
        INSERT INTO passengers (
            full_name,
            train_id
        )
        VALUES (
            :full_name,
            :train_id
        );
    """)

    with engine.begin() as connection:
        connection.execute(query, {
            "full_name": full_name.strip(),
            "train_id": int(train_id),
        })


def update_passenger(
    passenger_id: int,
    full_name: str,
    train_id: int,
) -> None:
    # При изменении поезда PostgreSQL автоматически
    # синхронизирует станции с новым поездом.
    query = text("""
        UPDATE passengers
        SET
            full_name = :full_name,
            train_id = :train_id
        WHERE id = :passenger_id;
    """)

    with engine.begin() as connection:
        connection.execute(query, {
            "passenger_id": int(passenger_id),
            "full_name": full_name.strip(),
            "train_id": int(train_id),
        })


def delete_passenger(passenger_id: int) -> None:
    query = text("""
        DELETE FROM passengers
        WHERE id = :passenger_id;
    """)

    with engine.begin() as connection:
        connection.execute(query, {"passenger_id": int(passenger_id)})


# ============================================================
# БАГАЖ
# ============================================================

def load_luggage() -> pd.DataFrame:
    query = """
        SELECT
            l.id,
            l.passenger_id,
            p.full_name,
            t.train_number,
            t.departure_date,
            t.departure_station,
            t.destination_station,
            l.luggage_count,
            l.extra_places
        FROM luggage l
        JOIN passengers p
            ON p.id = l.passenger_id
        JOIN trains t
            ON t.id = p.train_id
        ORDER BY l.id;
    """
    return _normalize_dates(pd.read_sql(query, engine))


def get_passenger_options() -> pd.DataFrame:
    query = """
        SELECT
            p.id,
            p.full_name,
            p.train_id,
            t.train_number,
            t.departure_date,
            t.departure_station,
            t.destination_station
        FROM passengers p
        JOIN trains t
            ON t.id = p.train_id
        ORDER BY p.full_name, t.departure_date, p.id;
    """
    return _normalize_dates(pd.read_sql(query, engine))


def create_luggage(
    passenger_id: int,
    luggage_count: int,
    extra_places: int,
) -> None:
    query = text("""
        INSERT INTO luggage (
            passenger_id,
            luggage_count,
            extra_places
        )
        VALUES (
            :passenger_id,
            :luggage_count,
            :extra_places
        );
    """)

    with engine.begin() as connection:
        connection.execute(query, {
            "passenger_id": int(passenger_id),
            "luggage_count": int(luggage_count),
            "extra_places": int(extra_places),
        })


def update_luggage(
    luggage_id: int,
    passenger_id: int,
    luggage_count: int,
    extra_places: int,
) -> None:
    query = text("""
        UPDATE luggage
        SET
            passenger_id = :passenger_id,
            luggage_count = :luggage_count,
            extra_places = :extra_places
        WHERE id = :luggage_id;
    """)

    with engine.begin() as connection:
        connection.execute(query, {
            "luggage_id": int(luggage_id),
            "passenger_id": int(passenger_id),
            "luggage_count": int(luggage_count),
            "extra_places": int(extra_places),
        })


def delete_luggage(luggage_id: int) -> None:
    query = text("""
        DELETE FROM luggage
        WHERE id = :luggage_id;
    """)

    with engine.begin() as connection:
        connection.execute(query, {"luggage_id": int(luggage_id)})


# ============================================================
# DASHBOARD
# ============================================================

def get_dashboard_summary() -> dict:
    trains = load_trains()
    passengers = load_passengers()
    luggage = load_luggage()

    return {
        "trains": int(len(trains)),
        "passengers": int(len(passengers)),
        "tickets": int(trains["sold_tickets"].sum()) if not trains.empty else 0,
        "revenue": float(trains["total_revenue"].sum()) if not trains.empty else 0.0,
        "average_ticket_price": (
            float(trains["ticket_price"].mean())
            if not trains.empty else 0.0
        ),
        "luggage": int(luggage["luggage_count"].sum()) if not luggage.empty else 0,
    }


# ============================================================
# АНАЛИТИКА
# ============================================================

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


def get_daily_statistics() -> pd.DataFrame:
    query = """
        SELECT
            departure_date,
            COUNT(*) AS train_count,
            SUM(sold_tickets) AS total_tickets,
            SUM(total_revenue) AS total_revenue
        FROM trains
        GROUP BY departure_date
        ORDER BY departure_date;
    """
    return _normalize_dates(pd.read_sql(query, engine))


def get_carriage_statistics() -> pd.DataFrame:
    query = """
        SELECT
            carriage_type,
            COUNT(*) AS train_count,
            SUM(sold_tickets) AS total_tickets,
            SUM(total_revenue) AS total_revenue,
            AVG(ticket_price) AS average_ticket_price
        FROM trains
        GROUP BY carriage_type
        ORDER BY total_revenue DESC;
    """
    return pd.read_sql(query, engine)

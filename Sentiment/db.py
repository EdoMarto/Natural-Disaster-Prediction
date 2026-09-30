import os

import pyodbc

# Set CLIMATEGECKO_DB_CONNECTION to a full ODBC connection string to use a remote database.
DEFAULT_CONNECTION = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost;DATABASE=ClimateGeckoDB;Trusted_Connection=yes;"
)


def get_connection():
    return pyodbc.connect(os.environ.get("CLIMATEGECKO_DB_CONNECTION", DEFAULT_CONNECTION))

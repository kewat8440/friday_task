import configparser
from pathlib import Path

import psycopg2
from psycopg2.extras import execute_values
import pandas as pd


# ---------- CONFIG LOADING ----------

config = configparser.ConfigParser()
configPath = Path(__file__).with_name("config.ini")

if not configPath.exists():
    raise FileNotFoundError(f"config.ini not found at: {configPath}")

config.read(configPath)

if "DB_Credentials" not in config:
    raise KeyError("Section [DB_Credentials] not found in config.ini")

if "Table_Names" not in config:
    raise KeyError("Section [Table_Names] not found in config.ini")

dbUsername = config["DB_Credentials"]["username"]
dbPassword = config["DB_Credentials"]["password"]
dbHost = config["DB_Credentials"]["host"]
dbPort = config["DB_Credentials"]["port"]
dbName = config["DB_Credentials"]["db_name"]

DATA_TABLE_NAME = config["Table_Names"]["data_table"]
SUMMARY_TABLE_NAME = config["Table_Names"]["summary_table"]
LOG_TABLE_NAME = config["Table_Names"]["log_table"]


# ---------- DB CONNECTION ----------

def getConnection():
    return psycopg2.connect(
        host=dbHost,
        port=dbPort,
        user=dbUsername,
        password=dbPassword,
        dbname=dbName
    )


# ---------- LOG TABLE HELPERS ----------

def ensureLogTable(conn):
    createSql = f"""
    CREATE TABLE IF NOT EXISTS {LOG_TABLE_NAME} (
        id SERIAL PRIMARY KEY,
        run_time      TIMESTAMPTZ DEFAULT NOW(),
        file_name     TEXT,
        status        TEXT,
        message       TEXT,
        rows_inserted INTEGER
    );
    """
    with conn.cursor() as cur:
        cur.execute(createSql)
    conn.commit()


def logRun(conn, fileName, status, message, rowsInserted=0):
    ensureLogTable(conn)

    insertSql = f"""
    INSERT INTO {LOG_TABLE_NAME} (file_name, status, message, rows_inserted)
    VALUES (%s, %s, %s, %s);
    """
    with conn.cursor() as cur:
        cur.execute(insertSql, (fileName, status, message, rowsInserted))
    conn.commit()


# ---------- MAIN DATA TABLE HELPERS ----------

def ensureDataTable(conn):
    createSql = f"""
    CREATE TABLE IF NOT EXISTS {DATA_TABLE_NAME} (
        id SERIAL PRIMARY KEY,
        snapshot_date          DATE,
        hour                   INTEGER,
        time_block             TEXT,
        purchase_bid_mw        NUMERIC,
        sell_bid_mw            NUMERIC,
        mcv_mw                 NUMERIC,
        final_scheduled_mw     NUMERIC,
        mcp_rs_per_mwh         NUMERIC,
        file_name              TEXT,
        loaded_at              TIMESTAMPTZ DEFAULT NOW()
    );
    """
    with conn.cursor() as cur:
        cur.execute(createSql)
    conn.commit()


def insertMarketSnapshot(conn, df, fileName):
    """
    Insert first (main) table from Excel.
    Expects df to have columns:
    date, hour, time_block, purchase_bid_mw, sell_bid_mw,
    mcv_mw, final_scheduled_mw, mcp_rs_per_mwh
    """
    if df.empty:
        return 0

    ensureDataTable(conn)

    insertSql = f"""
    INSERT INTO {DATA_TABLE_NAME} (
        snapshot_date,
        hour,
        time_block,
        purchase_bid_mw,
        sell_bid_mw,
        mcv_mw,
        final_scheduled_mw,
        mcp_rs_per_mwh,
        file_name
    )
    VALUES %s
    """

    rows = []
    for _, row in df.iterrows():
        dateVal = None if pd.isna(row["date"]) else row["date"].date()
        hourVal = None if pd.isna(row["hour"]) else int(row["hour"])

        rows.append((
            dateVal,
            hourVal,
            row["time_block"],
            row["purchase_bid_mw"],
            row["sell_bid_mw"],
            row["mcv_mw"],
            row["final_scheduled_mw"],
            row["mcp_rs_per_mwh"],
            fileName
        ))

    with conn.cursor() as cur:
        execute_values(cur, insertSql, rows)

    conn.commit()
    return len(rows)


# ---------- SUMMARY TABLE HELPERS (SECOND TABLE) ----------

def ensureSummaryTable(conn):
    createSql = f"""
    CREATE TABLE IF NOT EXISTS {SUMMARY_TABLE_NAME} (
        id SERIAL PRIMARY KEY,
        snapshot_date          DATE,
        summary_type           TEXT,
        purchase_bid_mwh       NUMERIC,
        sell_bid_mwh           NUMERIC,
        mcv_mwh                NUMERIC,
        final_scheduled_mwh    NUMERIC,
        mcp_rs_per_mwh         NUMERIC,
        file_name              TEXT,
        loaded_at              TIMESTAMPTZ DEFAULT NOW()
    );
    """
    with conn.cursor() as cur:
        cur.execute(createSql)
    conn.commit()


def insertMarketSnapshotSummary(conn, df, fileName):
    """
    Insert second (summary) table from Excel.
    Expects df to have columns:
    date, summary_type, purchase_bid_mwh, sell_bid_mwh,
    mcv_mwh, final_scheduled_mwh, mcp_rs_per_mwh
    """
    if df.empty:
        return 0

    ensureSummaryTable(conn)

    insertSql = f"""
    INSERT INTO {SUMMARY_TABLE_NAME} (
        snapshot_date,
        summary_type,
        purchase_bid_mwh,
        sell_bid_mwh,
        mcv_mwh,
        final_scheduled_mwh,
        mcp_rs_per_mwh,
        file_name
    )
    VALUES %s
    """

    rows = []
    for _, row in df.iterrows():
        dateVal = None if pd.isna(row["date"]) else row["date"].date()

        rows.append((
            dateVal,
            row["summary_type"],
            row["purchase_bid_mwh"],
            row["sell_bid_mwh"],
            row["mcv_mwh"],
            row["final_scheduled_mwh"],
            row["mcp_rs_per_mwh"],
            fileName
        ))

    with conn.cursor() as cur:
        execute_values(cur, insertSql, rows)

    conn.commit()
    return len(rows)

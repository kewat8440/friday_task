import configparser
from pathlib import Path

from db import (
    getConnection,
    logRun,
    insertMarketSnapshot,
    insertMarketSnapshotSummary,
)
from excelentry import (
    findExcelInFolder,
    parseMarketSnapshotExcel,
    parseMarketSnapshotSummaryExcel,
)


# ---------- CONFIG FOR FOLDERS ----------
config = configparser.ConfigParser()
configPath = Path(__file__).with_name("config.ini")
config.read(configPath)

inputFolderConfig = config["Folders"]["input_folder"].strip()


def main():
    conn = None
    excelPath = None
    fileName = None

    try:
        # ---------- Establish DB Connection ONCE ----------
        conn = getConnection()

        baseDir = Path(__file__).resolve().parent
        excelFolderPath = (baseDir / inputFolderConfig).resolve()

        print(f"Looking for Excel in: {excelFolderPath}")
        excelPath = findExcelInFolder(str(excelFolderPath))

        # ---------- No Excel Found ----------
        if not excelPath:
            logRun(conn, None, "error",
                   f"No Excel file found in folder: {excelFolderPath}", 0)
            print(f"ERROR: No Excel file found in folder: {excelFolderPath}")
            return

        fileName = excelPath.split("/")[-1].split("\\")[-1]
        print(f"Found Excel file: {fileName}")

        # ---------- Parse FIRST table ----------
        dfMain = parseMarketSnapshotExcel(excelPath)
        print(f"Parsed {len(dfMain)} rows from main table.")

        # ---------- Parse SECOND table ----------
        dfSummary = parseMarketSnapshotSummaryExcel(excelPath)
        print(f"Parsed {len(dfSummary)} rows from summary table.")

        # ---------- Insert BOTH tables (using same DB conn) ----------
        rowsInsertedMain = insertMarketSnapshot(conn, dfMain, fileName)
        rowsInsertedSummary = insertMarketSnapshotSummary(conn, dfSummary, fileName)

        totalRows = rowsInsertedMain + rowsInsertedSummary

        message = (
            f"Inserted {rowsInsertedMain} rows into {config['Table_Names']['data_table']} and "
            f"{rowsInsertedSummary} rows into {config['Table_Names']['summary_table']} "
            f"from {fileName}."
        )

        # ---------- Log Success ----------
        logRun(conn, fileName, "success", message, totalRows)
        print(message)

    except Exception as e:
        errorMessage = f"Run failed: {e}"
        print(errorMessage)

        # ---------- Log Error ----------
        try:
            if conn:
                logRun(
                    conn,
                    fileName if fileName else None,
                    "error",
                    errorMessage,
                    0
                )
        except Exception as logErr:
            print(f"Additionally failed to log error: {logErr}")

    finally:
        # ---------- CLOSE DB CONNECTION ONCE ----------
        if conn:
            conn.close()


if __name__ == "__main__":
    main()

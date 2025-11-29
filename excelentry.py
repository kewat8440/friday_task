import os
import pandas as pd


def findExcelInFolder(folderPath):
    """
    Returns full path of the first Excel file found in the folder,
    or None if no Excel file present.
    """
    excelExtensions = (".xls", ".xlsx", ".xlsm")

    if not os.path.isdir(folderPath):
        return None

    for file in os.listdir(folderPath):
        if file.lower().endswith(excelExtensions):
            return os.path.join(folderPath, file)

    return None


def parseMarketSnapshotExcel(excelPath):
    """
    Parse the FIRST (main) table:

    Header row like:
    Date | Hour | Time Block | Purchase Bid (MW) | Sell Bid (MW) |
    MCV (MW) | Final Scheduled Volume (MW) | MCP (Rs/MWh) *
    """
    dfRaw = pd.read_excel(excelPath, header=None)

    headerRowIndex = None
    for idx in range(len(dfRaw)):
        firstCell = str(dfRaw.iloc[idx, 0]).strip().lower()
        secondCell = str(dfRaw.iloc[idx, 1]).strip().lower()
        if firstCell == "date" and secondCell == "hour":
            headerRowIndex = idx
            break

    if headerRowIndex is None:
        raise ValueError("Could not find main header row with 'Date' and 'Hour'.")

    dataStartIndex = headerRowIndex + 1
    df = dfRaw.iloc[dataStartIndex:].copy()
    df.columns = dfRaw.iloc[headerRowIndex]

    df = df.rename(columns={
        "Date": "date",
        "Hour": "hour",
        "Time Block": "time_block",
        "Purchase Bid (MW)": "purchase_bid_mw",
        "Sell Bid (MW)": "sell_bid_mw",
        "MCV (MW)": "mcv_mw",
        "Final Scheduled Volume (MW)": "final_scheduled_mw",
        "MCP (Rs/MWh) *": "mcp_rs_per_mwh"
    })

    expectedCols = [
        "date",
        "hour",
        "time_block",
        "purchase_bid_mw",
        "sell_bid_mw",
        "mcv_mw",
        "final_scheduled_mw",
        "mcp_rs_per_mwh",
    ]
    df = df[expectedCols]

    # Type conversions
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["hour"] = pd.to_numeric(df["hour"], errors="coerce")

    for col in [
        "purchase_bid_mw",
        "sell_bid_mw",
        "mcv_mw",
        "final_scheduled_mw",
        "mcp_rs_per_mwh",
    ]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Remove non-data rows
    df = df[df["date"].notna()]
    df = df[df["hour"].notna()]
    df = df[df["time_block"].notna()]

    df = df.reset_index(drop=True)
    return df


def parseMarketSnapshotSummaryExcel(excelPath):
    """
    Parse the SECOND (summary) table.

    Header row exactly like:
    Date | Summary | Purchase Bid | Sell Bid | MCV | Final Scheduled Volume | MCP (Rs/MWh) *
    (possibly with extra spaces around names)
    """
    dfRaw = pd.read_excel(excelPath, header=None)

    # 1) Find header row: first col "Date", second col "Summary"
    headerRowIndex = None
    for idx in range(len(dfRaw)):
        firstCell = str(dfRaw.iloc[idx, 0]).strip().lower()
        secondCell = str(dfRaw.iloc[idx, 1]).strip().lower()
        if firstCell == "date" and secondCell == "summary":
            headerRowIndex = idx
            break

    if headerRowIndex is None:
        # No summary table found → empty df with proper columns
        return pd.DataFrame(columns=[
            "date",
            "summary_type",
            "purchase_bid_mwh",
            "sell_bid_mwh",
            "mcv_mwh",
            "final_scheduled_mwh",
            "mcp_rs_per_mwh",
        ])

    # 2) Data starts after header row
    dataStartIndex = headerRowIndex + 1
    df = dfRaw.iloc[dataStartIndex:].copy()

    # Take header row, convert to string and strip spaces
    headerRow = dfRaw.iloc[headerRowIndex].astype(str).map(lambda x: x.strip())
    df.columns = headerRow

    # 3) Strip spaces from all column names to fix "Purchase Bid  " etc.
    df.columns = [str(c).strip() for c in df.columns]

    # 4) Rename to our internal names
    df = df.rename(columns={
        "Date": "date",
        "Summary": "summary_type",
        "Purchase Bid": "purchase_bid_mwh",
        "Sell Bid": "sell_bid_mwh",
        "MCV": "mcv_mwh",
        "Final Scheduled Volume": "final_scheduled_mwh",
        "MCP (Rs/MWh) *": "mcp_rs_per_mwh",
    })

    expectedCols = [
        "date",
        "summary_type",
        "purchase_bid_mwh",
        "sell_bid_mwh",
        "mcv_mwh",
        "final_scheduled_mwh",
        "mcp_rs_per_mwh",
    ]

    # Keep only our expected columns
    df = df[expectedCols]

    # 5) Type conversions
    df["date"] = pd.to_datetime(df["date"], errors="coerce", dayfirst=True)

    for col in [
        "purchase_bid_mwh",
        "sell_bid_mwh",
        "mcv_mwh",
        "final_scheduled_mwh",
        "mcp_rs_per_mwh",
    ]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 6) Drop empty / non-data rows
    df = df[df["date"].notna()]
    df = df[df["summary_type"].notna()]

    df = df.reset_index(drop=True)
    return df


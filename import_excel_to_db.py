
"""
import_excel_to_db.py  (v2 - with proper type handling)
=========================================================
Imports W:\\test1\\UnshippedDTFOrders_11042026_031657.xlsx
  Sheet "DTF Orders"        -> tblCustomOrder
  Sheet "DTF Order Details" -> tblCustomOrderDetails
Safe to re-run: skips rows where primary key already exists.
"""

import pandas as pd
import pyodbc
from datetime import datetime

EXCEL_FILE = r"W:\test1\UnshippedDTFOrders_11042026_031657.xlsx"
DB_CONN = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost\\SQLEXPRESS;"
    "DATABASE=dbAmazonCustomOrders;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

# Columns that must be stored as datetime (not string)
DATETIME_COLS_ORDERS = {
    'ConvertedPurchaseDate', 'CustomOrderDetailsGetTime', 'AlertEmailProcessTime',
    'DateAdd', 'ShippedStatusSetTime', 'NotesUpdateTime', 'ConvertedShipByDate',
    'UpdatedDate', 'ItemUpdatedTime'
}
BOOL_COLS_ORDERS = {
    'IsCustomOrderDetailsGet', 'IsShipped', 'IsNotesUpdated', 'IsItemTypeGet'
}
INT_COLS_ORDERS = {'Quantity'}

DATETIME_COLS_DETAILS = {'DateAdd', 'ProcessTime'}
BOOL_COLS_DETAILS = {
    'IsFrontPSDDownload', 'IsBackPSDDownload', 'IsPocketPSDDownload',
    'IsSleevePSDDownload', 'IsAdditionalPSDDownload', 'IsFrontLocation',
    'IsBackLocation', 'IsPocketLocation', 'IsSleeveLocation', 'IsOrderClick',
    'IsOrderProcess', 'IsDesignComplete', 'IsOrderItemIdUpdated'
}

def safe_datetime(val):
    if val is None or val == '' or (isinstance(val, float) and __import__('math').isnan(val)):
        return None
    if isinstance(val, str):
        for fmt in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try:
                return datetime.strptime(val.strip(), fmt)
            except ValueError:
                continue
        return None
    return val

def safe_bool(val):
    if val is None or val == '' or (isinstance(val, float) and __import__('math').isnan(val)):
        return None
    if isinstance(val, str):
        return val.strip().lower() in ('true', '1', 'yes')
    return bool(val)

def safe_int(val):
    if val is None or val == '' or (isinstance(val, float) and __import__('math').isnan(val)):
        return None
    try:
        return int(float(val))
    except:
        return None

def clean_val(col, val, dt_cols, bool_cols, int_cols=set()):
    if col in dt_cols:
        return safe_datetime(val)
    if col in bool_cols:
        return safe_bool(val)
    if col in int_cols:
        return safe_int(val)
    # Generic None handling
    if val is None:
        return None
    if isinstance(val, float):
        import math
        if math.isnan(val):
            return None
    if isinstance(val, str) and val.strip() == '':
        return None
    return val

def import_sheet(conn, df, table, pk_col, dt_cols, bool_cols, int_cols=set()):
    cursor = conn.cursor()
    cursor.execute(f"SELECT [{pk_col}] FROM [{table}]")
    existing = {str(r[0]) for r in cursor.fetchall()}
    print(f"  Existing rows in {table}: {len(existing)}")

    cols = list(df.columns)
    ph   = ", ".join(["?" for _ in cols])
    cn   = ", ".join([f"[{c}]" for c in cols])
    sql  = f"INSERT INTO [{table}] ({cn}) VALUES ({ph})"

    inserted = skipped = errors = 0
    for i, row in df.iterrows():
        pk_val = str(row[pk_col]) if pd.notna(row[pk_col]) else None
        if pk_val and pk_val in existing:
            skipped += 1
            continue
        values = [clean_val(c, row[c], dt_cols, bool_cols, int_cols) for c in cols]
        try:
            cursor.execute(sql, values)
            inserted += 1
        except Exception as e:
            errors += 1
            if errors <= 3:  # show first 3 errors only
                print(f"  ERROR row {i} ({pk_val}): {e}")
            elif errors == 4:
                print(f"  ... (suppressing further error messages)")

    conn.commit()
    print(f"  Result: {inserted} inserted | {skipped} skipped (already exist) | {errors} errors")
    return inserted, skipped, errors

def main():
    print(f"Loading: {EXCEL_FILE}")
    df_orders  = pd.read_excel(EXCEL_FILE, sheet_name="DTF Orders",        dtype=str)
    df_details = pd.read_excel(EXCEL_FILE, sheet_name="DTF Order Details", dtype=str)
    df_orders  = df_orders.where(pd.notna(df_orders), None)
    df_details = df_details.where(pd.notna(df_details), None)
    print(f"  Orders: {len(df_orders)} rows | Details: {len(df_details)} rows")

    conn = pyodbc.connect(DB_CONN)
    print("Connected to database OK\n")

    print("--- Importing tblCustomOrder ---")
    import_sheet(conn, df_orders, "tblCustomOrder",
                 "idCustomOrder", DATETIME_COLS_ORDERS, BOOL_COLS_ORDERS, INT_COLS_ORDERS)

    print("\n--- Importing tblCustomOrderDetails ---")
    import_sheet(conn, df_details, "tblCustomOrderDetails",
                 "idCustomOrderDetails", DATETIME_COLS_DETAILS, BOOL_COLS_DETAILS)

    conn.close()
    print("\nImport complete.")

if __name__ == "__main__":
    main()

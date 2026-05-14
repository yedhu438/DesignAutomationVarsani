"""
export_premium_fonts.py
Query local DB for up to 5 premium-font orders and export full details to Excel.

Run:
    python export_premium_fonts.py
Output:
    C:\Varsany\Output\premium_font_orders_YYYYMMDD_HHMMSS.xlsx
"""

import ast
import json
import os
import sys
from datetime import datetime

import pyodbc
import pandas as pd

# ─── CONFIG ───────────────────────────────────────────────────────────────────

DB_CONNECTION = (
    "DRIVER={SQL Server};"
    "SERVER=localhost\\SQLEXPRESS;"
    "DATABASE=dbAmazonCustomOrders;"
    "Trusted_Connection=yes;"
)

OUTPUT_DIR  = r"C:\Varsany\Output"
OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    f"premium_font_orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
)

LIMIT = 5  # number of orders to export

PREMIUM_FONT_KEYS = {
    "smartkids", "colorfulblocks", "paintsplashesrainbow",
    "wavemermaid", "refractionray", "camoblock", "spiderweb",
    "cozywinter", "soccerarmy", "tropicalflower", "vinylfont",
}

# ─── FONT HELPERS ─────────────────────────────────────────────────────────────

def _parse_fonts_json(raw):
    """Return (font_name, is_premium) from a Fonts column value."""
    if not raw:
        return "Arial", False
    s = str(raw).strip()
    if s.startswith("{"):
        d = None
        try:
            d = json.loads(s)
        except Exception:
            try:
                d = ast.literal_eval(s)
            except Exception:
                pass
        if d is not None:
            premium = (d.get("PremiumFont") or "").strip()
            if premium and premium.lower() not in ("no", "none", "false", ""):
                return premium, True
            return (d.get("NormalFont") or "Arial").strip(), False
    # Plain string
    plain = s.lower().replace(" ", "")
    return s, plain in PREMIUM_FONT_KEYS


def _is_premium(raw):
    return _parse_fonts_json(raw)[1]


def _premium_zones(row, has_db_cols):
    """Return comma-separated list of zones that use a premium font."""
    zones = []
    for zone in ("Front", "Back", "Pocket", "Sleeve"):
        # Check dedicated DB column first
        if has_db_cols:
            val = row.get(f"{zone}PremiumFont") or ""
            if str(val).strip().lower() in ("yes", "1", "true"):
                zones.append(zone)
                continue
        # Fall back to parsing Fonts JSON
        if _is_premium(row.get(f"{zone}Fonts")):
            zones.append(zone)
    return ", ".join(zones)


def _font_label(raw):
    """Return 'FontName  [PREMIUM]' or just 'FontName'."""
    name, is_prem = _parse_fonts_json(raw)
    return f"{name}  [PREMIUM]" if is_prem else name

# ─── DATABASE ─────────────────────────────────────────────────────────────────

def _connect():
    return pyodbc.connect(DB_CONNECTION)


def fetch_premium_orders(limit=LIMIT):
    conn = _connect()
    cur  = conn.cursor()

    # Detect whether dedicated premium-font columns exist on this DB
    has_db_cols = False
    try:
        cur.execute("SELECT TOP 0 FrontPremiumFont FROM tblCustomOrderDetails")
        has_db_cols = True
    except Exception:
        conn.close()
        conn = _connect()
        cur  = conn.cursor()

    premium_select = ""
    premium_where  = ""
    if has_db_cols:
        premium_select = (
            ", d.FrontPremiumFont, d.BackPremiumFont"
            ", d.PocketPremiumFont, d.SleevePremiumFont"
        )
        premium_where = """
            AND (
                LOWER(ISNULL(d.FrontPremiumFont,''))  IN ('yes','1','true') OR
                LOWER(ISNULL(d.BackPremiumFont,''))   IN ('yes','1','true') OR
                LOWER(ISNULL(d.PocketPremiumFont,'')) IN ('yes','1','true') OR
                LOWER(ISNULL(d.SleevePremiumFont,'')) IN ('yes','1','true')
            )"""

    # When no dedicated DB columns, filter in SQL via JSON LIKE patterns
    json_where = ""
    if not has_db_cols:
        json_where = """
            AND (
                (ISNULL(d.FrontFonts,'')  LIKE '%"PremiumFont":"%'
                 AND ISNULL(d.FrontFonts,'')  NOT LIKE '%"PremiumFont":""%'
                 AND ISNULL(d.FrontFonts,'')  NOT LIKE '%"PremiumFont":"No"%') OR
                (ISNULL(d.BackFonts,'')   LIKE '%"PremiumFont":"%'
                 AND ISNULL(d.BackFonts,'')   NOT LIKE '%"PremiumFont":""%'
                 AND ISNULL(d.BackFonts,'')   NOT LIKE '%"PremiumFont":"No"%') OR
                (ISNULL(d.PocketFonts,'') LIKE '%"PremiumFont":"%'
                 AND ISNULL(d.PocketFonts,'') NOT LIKE '%"PremiumFont":""%'
                 AND ISNULL(d.PocketFonts,'') NOT LIKE '%"PremiumFont":"No"%') OR
                (ISNULL(d.SleeveFonts,'') LIKE '%"PremiumFont":"%'
                 AND ISNULL(d.SleeveFonts,'') NOT LIKE '%"PremiumFont":""%'
                 AND ISNULL(d.SleeveFonts,'') NOT LIKE '%"PremiumFont":"No"%')
            )"""

    sql = f"""
        SELECT TOP {limit}
            o.OrderID,
            o.SKU,
            o.ItemType,
            o.Quantity,
            o.IsShipped,
            d.idCustomOrderDetails,
            d.PrintLocation,
            d.IsDesignComplete,
            d.IsOrderProcess,
            d.FrontText,   d.FrontFonts,   d.FrontColours,
            d.FrontImage,  d.FrontPreviewImage,
            d.BackText,    d.BackFonts,    d.BackColours,
            d.BackImage,   d.BackPreviewImage,
            d.PocketText,  d.PocketFonts,  d.PocketColours,
            d.PocketImage, d.PocketPreviewImage,
            d.SleeveText,  d.SleeveImage,  d.SleevePreviewImage
            {premium_select}
        FROM tblCustomOrder o
        JOIN tblCustomOrderDetails d ON o.idCustomOrder = d.idCustomOrder
        WHERE 1=1
          {premium_where}
          {json_where}
        ORDER BY o.DateAdd ASC
    """
    cur.execute(sql)
    cols = [c[0] for c in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    conn.close()

    # When no dedicated DB column, filter by parsing Fonts JSON in Python
    if not has_db_cols:
        filtered = []
        for row in rows:
            if any(_is_premium(row.get(f"{z}Fonts")) for z in ("Front", "Back", "Pocket", "Sleeve")):
                filtered.append(row)
        rows = filtered

    return rows[:limit], has_db_cols

# ─── EXPORT ───────────────────────────────────────────────────────────────────

def export(rows, has_db_cols):
    if not rows:
        print("No premium font orders found in the database.")
        return None

    df = pd.DataFrame(rows)

    # Add helper columns at the front
    df.insert(0, "PremiumZones",
              df.apply(lambda r: _premium_zones(r, has_db_cols), axis=1))

    # Replace raw Fonts column values with readable labels
    for zone in ("Front", "Back", "Pocket", "Sleeve"):
        col = f"{zone}Fonts"
        if col in df.columns:
            df[col] = df[col].apply(lambda v: _font_label(v) if pd.notna(v) and v else "")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Premium Font Orders")

        ws = writer.sheets["Premium Font Orders"]

        # Auto-fit column widths (capped at 60)
        for col_cells in ws.columns:
            width = max((len(str(c.value or "")) for c in col_cells), default=10)
            ws.column_dimensions[col_cells[0].column_letter].width = min(width + 2, 60)

        # Bold header row
        from openpyxl.styles import Font as XLFont
        for cell in ws[1]:
            cell.font = XLFont(bold=True)

    print(f"\nExported {len(df)} premium font order(s): {OUTPUT_FILE}")
    return OUTPUT_FILE

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print(f"Connecting to DB…")
    try:
        rows, has_db_cols = fetch_premium_orders(LIMIT)
    except pyodbc.Error as e:
        print(f"DB connection failed: {e}")
        sys.exit(1)

    print(f"Found {len(rows)} premium font order(s).\n")
    for i, r in enumerate(rows, 1):
        zones = _premium_zones(r, has_db_cols)
        fonts = [_parse_fonts_json(r.get(f"{z}Fonts"))[0]
                 for z in ("Front", "Back", "Pocket", "Sleeve")
                 if _is_premium(r.get(f"{z}Fonts"))]
        print(f"  {i}. OrderID={r['OrderID']}  SKU={r['SKU']}  "
              f"Zones=[{zones}]  Fonts={fonts}")

    export(rows, has_db_cols)


if __name__ == "__main__":
    main()

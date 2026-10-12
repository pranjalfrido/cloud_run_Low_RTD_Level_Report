"""
Frido Low RTD Level Report Generator
Reads data from the same pre-computed JSON the dashboard serves:
  https://frido-navigator.vercel.app/inv-data-7d.json

This guarantees numbers match the dashboard export exactly.
Filter: RTD Level = Low (same as dashboard inventory_detail export)

Run: python generate_low_rtd_report.py
"""

import json
import math
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import xlsxwriter

# ── Config ────────────────────────────────────────────────────────────────────

INV_JSON_URL = "https://frido-navigator.vercel.app/inv-data-7d.json"
OUTPUT_DIR   = Path(__file__).parent

LOCATION_ORDER = ["PNQ", "GGN", "BLR", "MUM", "KOL", "HYD", "CHN"]

# ── Fetch JSON ────────────────────────────────────────────────────────────────

def fetch_inv_json():
    print(f"Fetching dashboard JSON: {INV_JSON_URL}")
    with urllib.request.urlopen(INV_JSON_URL, timeout=60) as r:
        data = json.load(r)
    print(f"  asOf          : {data.get('asOf', '?')}")
    print(f"  snapshotUpdated: {data.get('lastSnapshotUpdated', '?')}")
    print(f"  avgSaleWindow : {data.get('avgSaleWindow', '?')}")
    print(f"  Total SKUs    : {len(data.get('skus', []))}")
    return data

# ── Build Detail Rows ─────────────────────────────────────────────────────────

def regular_locations(s):
    """
    Mirrors regularLocations() in InventoryHealthPage.jsx:
    If facilities[] exists, aggregate only facilityType=Regular per location,
    then merge avgSale/doi/stockStatus from s.locations[].
    This is what the dashboard CSV export uses.
    """
    facilities = s.get("facilities", [])
    if not facilities:
        return s.get("locations", [])

    by_loc = {}
    for f in facilities:
        if f.get("facilityType") != "Regular":
            continue
        loc = f.get("location", "")
        if loc not in LOCATION_ORDER:
            continue
        if loc not in by_loc:
            by_loc[loc] = {"location": loc, "totalInvt": 0, "rtdInvt": 0,
                           "rawInvt": 0, "rawBlockedInvt": 0,
                           "avgSale": 0, "doi": 0, "stockStatus": ""}
        acc = by_loc[loc]
        acc["totalInvt"]      += f.get("totalInvt",      0) or 0
        acc["rtdInvt"]        += f.get("rtdInvt",        0) or 0
        acc["rawInvt"]        += f.get("rawInvt",        0) or 0
        acc["rawBlockedInvt"] += f.get("rawBlockedInvt", 0) or 0

    # Merge avgSale, doi, stockStatus from s.locations (same as JS)
    for l in s.get("locations", []):
        loc = l.get("location", "")
        if loc in by_loc:
            acc = by_loc[loc]
            acc["avgSale"]     = l.get("avgSale", 0) or 0
            total = acc["totalInvt"]
            denom = math.ceil(acc["avgSale"] or 0)
            if total > 0 and denom == 0:
                acc["doi"] = None
            elif denom > 0:
                acc["doi"] = math.floor(total / denom)
            else:
                acc["doi"] = 0
            # stockStatus
            doi = acc["doi"]
            avg = acc["avgSale"]
            if total == 0:
                acc["stockStatus"] = "Out of Stock"
            elif avg == 0:
                acc["stockStatus"] = "No Demand"
            elif doi is None:
                acc["stockStatus"] = "No Demand"
            elif doi <= 2:
                acc["stockStatus"] = "Critical"
            elif doi <= 15:
                acc["stockStatus"] = "Low"
            elif doi <= 45:
                acc["stockStatus"] = "Sufficient"
            else:
                acc["stockStatus"] = "Excess"

    return list(by_loc.values())


def build_detail_rows(data):
    skus = data.get("skus", [])
    detail_rows = []

    for s in skus:
        if s.get("rtdLevel") != "Low":
            continue

        # SKU-level all-FC totals (across all locations from s.locations[])
        all_fc_total  = round(s.get("totalInvt",      0) or 0)
        all_fc_rtd    = round(s.get("rtdInvt",        0) or 0)
        all_fc_raw    = round(s.get("rawInvt",        0) or 0)
        all_fc_blocked= round(s.get("rawBlockedInvt", 0) or 0)
        all_fc_avg    = s.get("avgSale", 0) or 0
        rtd_level     = s.get("rtdLevel", "")

        for loc in regular_locations(s):
            location    = loc.get("location", "")
            total_invt  = loc.get("totalInvt",      0) or 0
            avg_sale    = loc.get("avgSale",         0) or 0

            if total_invt == 0 and avg_sale == 0:
                continue

            detail_rows.append({
                "category":        s.get("category",    ""),
                "sub_category":    s.get("subCategory", ""),
                "product_id":      s.get("sku",         ""),
                "location":        location,
                "rtd_invt":        round(loc.get("rtdInvt",        0) or 0),
                "raw_invt":        round(loc.get("rawInvt",        0) or 0),
                "raw_blocked":     round(loc.get("rawBlockedInvt", 0) or 0),
                "total_invt":      round(total_invt),
                "avg_sale":        avg_sale,
                "doi":             loc.get("doi", 0) or 0,
                "stock_status":    loc.get("stockStatus", ""),
                "website_status":  s.get("websiteStatus", ""),
                # SKU-level all-FC fields for location sheets
                "all_fc_total":    all_fc_total,
                "all_fc_rtd":      all_fc_rtd,
                "all_fc_raw":      all_fc_raw,
                "all_fc_blocked":  all_fc_blocked,
                "all_fc_avg_sale": all_fc_avg,
                "rtd_level":       rtd_level,
            })

    return detail_rows

# ── Excel Writer ──────────────────────────────────────────────────────────────

def safe_int(val):
    try:
        if val is None:
            return 0
        f = float(val)
        return 0 if (f != f) else int(f)
    except (TypeError, ValueError):
        return 0

def write_excel(detail_rows, report_date, output_path):
    date_str = report_date.strftime("%d %b %Y")

    NAV = "#1A3A5C"
    GRY = "#F5F5F5"
    WHT = "#FFFFFF"
    RED = "#F4CCCC"
    ORA = "#FCE5CD"
    YEL = "#FFF2CC"
    GRN = "#D9EAD3"
    BLU = "#CFE2F3"

    wb = xlsxwriter.Workbook(str(output_path))
    wb.set_properties({"title": f"Frido Low RTD Level Report {date_str}"})

    def fmt(bold=False, bg=WHT, fc="#222222", num_fmt=None, align="left",
            border=1, wrap=False, size=10):
        d = {
            "font_name": "Aptos Narrow", "font_size": size, "font_color": fc,
            "bg_color": bg, "bold": bold, "align": align, "valign": "vcenter",
            "border": border, "border_color": "#CCCCCC", "text_wrap": wrap,
        }
        if num_fmt:
            d["num_format"] = num_fmt
        return wb.add_format(d)

    F = {
        "hdr":    fmt(bold=True, bg=NAV, fc=WHT, align="center"),
        "txt":    fmt(),
        "txt_g":  fmt(bg=GRY),
        "ctr":    fmt(align="center"),
        "ctr_g":  fmt(bg=GRY, align="center"),
        "num":    fmt(num_fmt="#,##0", align="right"),
        "num_g":  fmt(bg=GRY, num_fmt="#,##0", align="right"),
        "num_b":  fmt(bold=True, num_fmt="#,##0", align="right"),
        "num_bg": fmt(bold=True, bg=GRY, num_fmt="#,##0", align="right"),
        "tot_t":  fmt(bold=True, bg=BLU, fc=NAV),
        "tot_n":  fmt(bold=True, bg=BLU, fc=NAV, num_fmt="#,##0", align="right"),
        "s_crit": fmt(bg=RED, align="center", bold=True),
        "s_low":  fmt(bg=ORA, align="center", bold=True),
        "s_suff": fmt(bg=GRN, align="center", bold=True),
        "s_exc":  fmt(bg=YEL, align="center", bold=True),
        "s_other":fmt(bg=GRY, align="center"),
    }

    def status_fmt(s):
        return {"Critical": F["s_crit"], "Low": F["s_low"],
                "Sufficient": F["s_suff"], "Excess": F["s_exc"]}.get(s, F["s_other"])

    COLS = [
        ("Category",               "category",       18, "txt"),
        ("Sub-category",           "sub_category",   28, "txt"),
        ("Product ID",             "product_id",     16, "txt"),
        ("Location",               "location",        9, "ctr"),
        ("RTD Inventory",          "rtd_invt",       12, "num"),
        ("RAW Inventory",          "raw_invt",       12, "num"),
        ("RAW Blocked Inventory",  "raw_blocked",    18, "num"),
        ("Total Inventory",        "total_invt",     14, "num_b"),
        ("Avg Sale (B2C)",         "avg_sale",       13, "num"),
        ("DOI",                    "doi",             7, "num"),
        ("Stock Status",           "stock_status",   13, "status"),
    ]

    print("  Writing inventory_detail sheet...")
    ws = wb.add_worksheet("inventory_detail")
    ws.hide_gridlines(2)
    ws.set_zoom(90)
    ws.set_tab_color(NAV)
    ws.freeze_panes(1, 3)
    ws.set_row(0, 22)

    for c, (h, _, w, _) in enumerate(COLS):
        ws.write(0, c, h, F["hdr"])
        ws.set_column(c, c, w)

    df = pd.DataFrame(detail_rows)
    detail_df = (
        df.sort_values(["category", "sub_category", "product_id", "location"])
        .reset_index(drop=True)
    )

    for i, (_, row) in enumerate(detail_df.iterrows()):
        alt = (i % 2 == 1)
        for c, (_, col, _, ftype) in enumerate(COLS):
            val = row[col]
            if ftype == "txt":
                ws.write(i+1, c, str(val or ""), F["txt_g"] if alt else F["txt"])
            elif ftype == "ctr":
                ws.write(i+1, c, str(val or ""), F["ctr_g"] if alt else F["ctr"])
            elif ftype == "num_b":
                ws.write(i+1, c, safe_int(val), F["num_bg"] if alt else F["num_b"])
            elif ftype == "status":
                ws.write(i+1, c, str(val or ""), status_fmt(str(val or "")))
            else:
                ws.write(i+1, c, safe_int(val), F["num_g"] if alt else F["num"])

    tr = len(detail_df) + 1
    ws.write(tr, 0, "TOTAL", F["tot_t"])
    for c, (_, col, _, ftype) in enumerate(COLS[1:], start=1):
        if ftype in ("num", "num_b"):
            try:
                ws.write(tr, c, round(detail_df[col].sum()), F["tot_n"])
            except Exception:
                ws.write(tr, c, "", F["tot_t"])
        else:
            ws.write(tr, c, "", F["tot_t"])

    # ── Sheet1: Pivot ─────────────────────────────────────────────────────────
    print("  Writing Sheet1 (pivot)...")

    # Filters: Total Invt > 0, Avg Sale > 0, Website = Live, DOI not in (0,1)
    piv_df = detail_df[
        (detail_df["total_invt"] > 0) &
        (detail_df["avg_sale"]   > 0) &
        (detail_df["website_status"] == "Live") &
        (detail_df["doi"] > 1)
    ].copy()

    LOCS = [l for l in LOCATION_ORDER if l in piv_df["location"].unique()]
    VALUES = [
        ("RTD Invt",          "rtd_invt"),
        ("RAW Invt",          "raw_invt"),
        ("RAW Blocked Invt",  "raw_blocked"),
        ("Regional Avg Sale", "avg_sale"),
    ]

    # Build pivot: rows=product_id, cols=location × value
    pivot = {}
    for _, row in piv_df.iterrows():
        pid = row["product_id"]
        loc = row["location"]
        if pid not in pivot:
            pivot[pid] = {}
        for vname, vcol in VALUES:
            key = (loc, vname)
            pivot[pid][key] = pivot[pid].get(key, 0) + row[vcol]

    sorted_skus = sorted(pivot.keys())

    # Total columns at the end
    TOT_COLS = [
        ("Total RTD Invt",          "rtd_invt"),
        ("Total RAW Invt",          "raw_invt"),
        ("Total RAW Blocked Invt",  "raw_blocked"),
        ("Total Regional Avg Sale", "avg_sale"),
    ]

    # Filter labels (rows 0-3)
    FLT = [
        ("Total Inventory",  "(excluding 0)"),
        ("Avg Sale (B2C)",   "(excluding 0)"),
        ("DOI",              "(excluding 0 and 1)"),
        ("Website Status",   "Live"),
    ]

    ws2 = wb.add_worksheet("Sheet1")
    ws2.hide_gridlines(2)
    ws2.set_zoom(85)
    ws2.set_tab_color(NAV)

    flt_lbl = fmt(bold=True, bg=GRY)
    flt_val = fmt(bg=GRY)
    loc_hdr = fmt(bold=True, bg=NAV, fc=WHT, align="center")
    val_hdr = fmt(bold=True, bg="#2C5282", fc=WHT, align="center", size=9)
    row_lbl = fmt(bold=True)
    row_g   = fmt(bg=GRY, bold=True)
    num_f   = fmt(num_fmt="#,##0", align="right")
    num_fg  = fmt(bg=GRY, num_fmt="#,##0", align="right")
    dash_f  = fmt(align="center")
    dash_fg = fmt(bg=GRY, align="center")
    tot_lbl = fmt(bold=True, bg=BLU, fc=NAV)
    tot_num = fmt(bold=True, bg=BLU, fc=NAV, num_fmt="#,##0", align="right")
    tot_dsh = fmt(bold=True, bg=BLU, fc=NAV, align="center")

    # Filter label rows (0-3)
    for fi, (lbl, val) in enumerate(FLT):
        ws2.write(fi, 0, lbl, flt_lbl)
        ws2.write(fi, 1, val, flt_val)

    # Col layout: col 0 = Product ID, then per location 4 value cols, then 4 totals
    COL_START = 1  # data starts at col 1

    # Row 5 = location group headers
    # Row 6 = value sub-headers
    # Row 7+ = data

    HDR_ROW  = 5
    SUBHDR_ROW = 6
    DATA_ROW = 7

    ws2.write(HDR_ROW, 0, "Row Labels", loc_hdr)
    ws2.set_column(0, 0, 22)

    col = COL_START
    loc_col_start = {}
    for loc in LOCS:
        loc_col_start[loc] = col
        ws2.merge_range(HDR_ROW, col, HDR_ROW, col + len(VALUES) - 1, loc, loc_hdr)
        for vi, (vname, _) in enumerate(VALUES):
            ws2.write(SUBHDR_ROW, col + vi, vname, val_hdr)
            ws2.set_column(col + vi, col + vi, 14)
        col += len(VALUES)

    # Total cols
    tot_col_start = col
    for ti, (tname, _) in enumerate(TOT_COLS):
        ws2.write(HDR_ROW, col + ti, tname, loc_hdr)
        ws2.write(SUBHDR_ROW, col + ti, "", val_hdr)
        ws2.set_column(col + ti, col + ti, 16)

    # Data rows
    for ri, pid in enumerate(sorted_skus):
        r = DATA_ROW + ri
        alt = ri % 2 == 1
        ws2.write(r, 0, pid, row_g if alt else row_lbl)

        col = COL_START
        for loc in LOCS:
            for vname, vcol in VALUES:
                val = pivot[pid].get((loc, vname), 0)
                if val == 0:
                    ws2.write(r, col, "-", dash_fg if alt else dash_f)
                else:
                    ws2.write(r, col, val, num_fg if alt else num_f)
                col += 1

        # Totals
        for ti, (_, vcol) in enumerate(TOT_COLS):
            total = sum(pivot[pid].get((loc, dict(VALUES)[vcol] if vcol in dict(VALUES) else ""), 0)
                        for loc in LOCS
                        for vname2, vcol2 in VALUES if vcol2 == vcol
                        for _ in [pivot[pid].get((loc, vname2), 0)])
            # simpler:
            total = sum(pivot[pid].get((loc, vname), 0) for loc in LOCS for vname, vc in VALUES if vc == vcol)
            ws2.write(r, tot_col_start + ti, total if total else "-",
                      tot_num if total else tot_dsh)

    # Grand total row
    tr2 = DATA_ROW + len(sorted_skus)
    ws2.write(tr2, 0, "Grand Total", tot_lbl)
    col = COL_START
    for loc in LOCS:
        for vname, vcol in VALUES:
            total = sum(pivot[pid].get((loc, vname), 0) for pid in sorted_skus)
            ws2.write(tr2, col, total if total else "-", tot_num if total else tot_dsh)
            col += 1
    for ti, (_, vcol) in enumerate(TOT_COLS):
        total = sum(pivot[pid].get((loc, vname), 0)
                    for pid in sorted_skus for loc in LOCS
                    for vname, vc in VALUES if vc == vcol)
        ws2.write(tr2, tot_col_start + ti, total if total else "-",
                  tot_num if total else tot_dsh)

    ws2.freeze_panes(DATA_ROW, 1)

    # ── Location Sheets (one per WH) ──────────────────────────────────────────
    # Group definitions: (group_label, bg_color, cols)
    # Col indices: 0-2 = identity, 3-6 = This FC, 7-10 = All FCs, 11 = Status
    GRP_IDENTITY = "#1A3A5C"   # dark navy  — SKU/Product/Category
    GRP_THIS_FC  = "#1F6B8E"   # steel blue — This FC columns
    GRP_ALL_FC   = "#B45309"   # warm amber — All FCs columns
    GRP_STATUS   = "#374151"   # slate      — RTD Level

    LOC_COLS = [
        # (header, field, width, ftype, group_bg)
        ("SKU",                       "product_id",      18, "txt",   GRP_IDENTITY),
        ("Product Name",              "sub_category",    32, "txt",   GRP_IDENTITY),
        ("Category",                  "category",        14, "txt",   GRP_IDENTITY),
        ("RTD Invt",                  "rtd_invt",        11, "num",   GRP_THIS_FC),
        ("RAW Invt",                  "raw_invt",        11, "num",   GRP_THIS_FC),
        ("RAW Blocked Invt",          "raw_blocked",     15, "num",   GRP_THIS_FC),
        ("Regional Avg Sale",         "avg_sale",        15, "num",   GRP_THIS_FC),
        ("All FCs Total Invt",        "all_fc_total",    16, "num_b", GRP_ALL_FC),
        ("All FCs RTD Invt",          "all_fc_rtd",      14, "num",   GRP_ALL_FC),
        ("All FCs Avg Sale",          "all_fc_avg_sale", 15, "num",   GRP_ALL_FC),
        ("All FCs DOI",               "all_fc_doi",      12, "doi",   GRP_ALL_FC),
        ("Overall RTD Level",         "rtd_level",       15, "rtd",   GRP_STATUS),
    ]

    # Group header spans: (label, start_col, end_col, bg)
    GRP_SPANS = [
        ("",          0, 2, GRP_IDENTITY),
        ("This FC",   3, 6, GRP_THIS_FC),
        ("All FCs",   7, 10, GRP_ALL_FC),
        ("Status",   11, 11, GRP_STATUS),
    ]

    def grp_hdr_fmt(bg):
        return fmt(bold=True, bg=bg, fc=WHT, align="center", size=10)

    def col_hdr_fmt(bg):
        return fmt(bold=True, bg=bg, fc=WHT, align="center", size=9)

    # DOI colour bands
    def doi_fmt(doi_val):
        if doi_val is None or doi_val == 0:
            return fmt(bg=RED, align="right", bold=True, num_fmt="#,##0")
        elif doi_val <= 2:
            return fmt(bg=RED, align="right", bold=True, num_fmt="#,##0")
        elif doi_val <= 15:
            return fmt(bg=ORA, align="right", bold=True, num_fmt="#,##0")
        elif doi_val <= 45:
            return fmt(bg=YEL, align="right", num_fmt="#,##0")
        else:
            return fmt(bg=GRN, align="right", num_fmt="#,##0")

    rtd_low_fmt  = fmt(bg=RED, align="center", bold=True)
    rtd_suff_fmt = fmt(bg=GRN, align="center", bold=True)

    for loc_name in LOCATION_ORDER:
        loc_rows = [
            r for r in detail_rows
            if r["location"] == loc_name
            and r["total_invt"] > 0
            and r["avg_sale"] > 0
            and r["website_status"] == "Live"
            and r["doi"] > 1
        ]
        if not loc_rows:
            continue

        print(f"  Writing sheet: {loc_name} ({len(loc_rows)} SKUs)...")
        ws_loc = wb.add_worksheet(loc_name)
        ws_loc.hide_gridlines(2)
        ws_loc.set_zoom(90)
        ws_loc.set_tab_color(GRP_THIS_FC)
        ws_loc.freeze_panes(2, 3)   # freeze 2 header rows + 3 identity cols
        ws_loc.set_row(0, 18)       # group header row
        ws_loc.set_row(1, 22)       # col header row

        # Set column widths
        for c, (_, _, w, _, _) in enumerate(LOC_COLS):
            ws_loc.set_column(c, c, w)

        # Row 0: group header spans
        for lbl, c_start, c_end, bg in GRP_SPANS:
            if c_start == c_end:
                ws_loc.write(0, c_start, lbl, grp_hdr_fmt(bg))
            else:
                ws_loc.merge_range(0, c_start, 0, c_end, lbl, grp_hdr_fmt(bg))

        # Row 1: column headers
        for c, (h, _, _, _, bg) in enumerate(LOC_COLS):
            ws_loc.write(1, c, h, col_hdr_fmt(bg))

        # Sort by category, product name
        loc_df = pd.DataFrame(loc_rows).sort_values(
            ["category", "sub_category", "product_id"]
        ).reset_index(drop=True)

        # Compute All FCs DOI per row
        def compute_all_fc_doi(row):
            rtd = row["all_fc_rtd"]
            avg = row["all_fc_avg_sale"]
            if avg == 0:
                return None
            return math.floor(rtd / avg) if rtd > 0 else 0

        loc_df["all_fc_doi"] = loc_df.apply(compute_all_fc_doi, axis=1)

        for i, (_, row) in enumerate(loc_df.iterrows()):
            alt = i % 2 == 1
            for c, (_, col, _, ftype, grp_bg) in enumerate(LOC_COLS):
                val = row[col]
                if ftype == "txt":
                    ws_loc.write(i+2, c, str(val or ""), F["txt_g"] if alt else F["txt"])
                elif ftype == "num":
                    ws_loc.write(i+2, c, safe_int(val), F["num_g"] if alt else F["num"])
                elif ftype == "num_b":
                    ws_loc.write(i+2, c, safe_int(val), F["num_bg"] if alt else F["num_b"])
                elif ftype == "doi":
                    ws_loc.write(i+2, c,
                                 safe_int(val) if val is not None else "-",
                                 doi_fmt(val))
                elif ftype == "rtd":
                    ws_loc.write(i+2, c, str(val or ""),
                                 rtd_low_fmt if val == "Low" else rtd_suff_fmt)

        # Totals row
        tr_loc = len(loc_df) + 2
        ws_loc.write(tr_loc, 0, "TOTAL", F["tot_t"])
        for c, (_, col, _, ftype, _) in enumerate(LOC_COLS[1:], start=1):
            if ftype in ("num", "num_b"):
                try:
                    ws_loc.write(tr_loc, c, round(loc_df[col].sum()), F["tot_n"])
                except Exception:
                    ws_loc.write(tr_loc, c, "", F["tot_t"])
            else:
                ws_loc.write(tr_loc, c, "", F["tot_t"])

    wb.close()
    print(f"  Excel written: {output_path}")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    today = datetime.now(timezone.utc).date()

    print(f"\n{'='*60}")
    print(f"  Frido Low RTD Level Report -- {today.strftime('%d %b %Y')}")
    print(f"{'='*60}\n")

    data = fetch_inv_json()

    print("\nBuilding detail rows (RTD Level = Low)...")
    detail_rows = build_detail_rows(data)
    unique_skus = len({r["product_id"] for r in detail_rows})
    print(f"  {unique_skus} Low RTD SKUs  ({len(detail_rows)} location-rows)")

    if not detail_rows:
        print("  No Low RTD SKUs -- no report generated.")
        return None, {}

    output_filename = f"Low_RTD_Level_{today.strftime('%d_%b_%Y')}.xlsx"
    output_path     = OUTPUT_DIR / output_filename

    print(f"\nWriting Excel: {output_filename}")
    write_excel(detail_rows, today, output_path)

    stats = {
        "sku_count":   unique_skus,
        "row_count":   len(detail_rows),
        "report_date": today.strftime("%d %b %Y"),
        "as_of":       data.get("asOf", ""),
    }

    print(f"\n{'='*60}")
    print(f"  DONE -- {unique_skus} SKUs  |  {len(detail_rows)} location rows")
    print(f"  Saved: {output_path}")
    print(f"{'='*60}\n")

    return str(output_path), stats


if __name__ == "__main__":
    main()

"""
Core business logic for the FA Gain/Loss reclassification pipeline.

Pipeline:
  1. Load Retirement_Data + Cost_Center_Master from the same Excel workbook.
  2. Merge, deduplicate columns, filter to F620/F630 functional areas.
  3. Build two balanced JE line groups (accounts 7615000 and 6645000).
  4. Write a styled Excel output file.
"""

import glob
import os
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill

# ── Constants ─────────────────────────────────────────────────────────────────

ACCOUNT_FA_GAIN_LOSS = 7615000   # Debit/Credit follows Net Reclass sign
ACCOUNT_INTERCO      = 6645000   # Reversed sign; includes customer & product
INTERCO_CUSTOMER     = "1000509"
FUNCTIONAL_AREAS     = ["F620", "F630"]
HEADER_COLOR         = "#D8E4BC"  # Light green


# ── Step helpers ──────────────────────────────────────────────────────────────

def find_excel_file(folder_path: str) -> str:
    """Return the first .xlsx or .xls file found in *folder_path*."""
    for pattern in ("*.xlsx", "*.xls"):
        matches = glob.glob(os.path.join(folder_path, pattern))
        if matches:
            return matches[0]
    raise FileNotFoundError(f"No Excel file found in: {folder_path}")


def build_je_item_text(df: pd.DataFrame) -> pd.Series:
    """
    Derive JE Item Text (max 50 chars) from each row:
      - If Cockpit Req No is present → 'reclass_gain_loss_<int(req_no)>'
      - Otherwise                    → 'reclass_gain_loss_<Item Text>'
    """
    raw = np.where(
        df["Cockpit Req No"].notna(),
        "reclass_gain_loss_" + df["Cockpit Req No"].fillna(0).astype(int).astype(str),
        "reclass_gain_loss_" + df["Item Text"].fillna("").astype(str),
    )
    return pd.Series(raw, index=df.index).str[:50]


def load_and_prepare(input_file: str) -> pd.DataFrame:
    """
    Read both sheets, merge on Costcenter, filter to target functional areas,
    and attach JE Item Text.
    """
    df = pd.read_excel(input_file, sheet_name="Retirement_Data")
    df["Net Reclass Amount"] = df["Gain"] + df["Loss"]

    cc_master = pd.read_excel(input_file, sheet_name="Cost_Center_Master")
    df = df.merge(cc_master, on="Costcenter", how="left")

    # After merge, columns shared by both sheets get _x / _y suffixes.
    # Keep _x (original data), drop _y (master duplicates).
    df = df.drop(columns=[c for c in df.columns if c.endswith("_y")])
    df = df.rename(columns={c: c[:-2] for c in df.columns if c.endswith("_x")})

    df = df[df["Functional Area"].isin(FUNCTIONAL_AREAS)].copy()
    df["JE Item Text"] = build_je_item_text(df)
    return df


def build_group1(df: pd.DataFrame) -> pd.DataFrame:
    """
    Account 7615000 — one line per (Company, Date, JE Text, PC, CC, FA).
    Positive Net Reclass → Debit; Negative → Credit.
    """
    keys = [
        "Company Code", "Posting Date", "JE Item Text",
        "Profit Center", "Costcenter", "Functional Area",
    ]
    g = df.groupby(keys, dropna=False)[["Net Reclass Amount"]].sum().reset_index()
    g["account_number"] = ACCOUNT_FA_GAIN_LOSS
    g["Transaction Currency Debit"]  = g["Net Reclass Amount"].where(g["Net Reclass Amount"] > 0, 0)
    g["Transaction Currency Credit"] = g["Net Reclass Amount"].where(g["Net Reclass Amount"] < 0, 0)
    return g


def build_group2(df: pd.DataFrame) -> pd.DataFrame:
    """
    Account 6645000 — grouped at Profit Center level (no Costcenter).
    Sign is reversed: Negative Net Reclass → Debit; Positive → Credit.
    """
    keys = [
        "Company Code", "Posting Date", "JE Item Text",
        "Profit Center", "Functional Area",
    ]
    g = df.groupby(keys, dropna=False)[["Net Reclass Amount"]].sum().reset_index()
    g["Costcenter"] = ""
    g["account_number"] = ACCOUNT_INTERCO
    g["Transaction Currency Debit"]  = g["Net Reclass Amount"].where(g["Net Reclass Amount"] < 0, 0)
    g["Transaction Currency Credit"] = g["Net Reclass Amount"].where(g["Net Reclass Amount"] > 0, 0)
    return g


def build_output_df(combined: pd.DataFrame) -> pd.DataFrame:
    """Map the combined groups to the final JE upload column layout."""
    is_interco = combined["account_number"] == ACCOUNT_INTERCO

    return pd.DataFrame({
        "*Company Code (4)":                combined["Company Code"],
        "*Journal Entry Type (2)":          "SA",
        "*Journal Entry Date":              combined["Posting Date"].dt.strftime("%Y%m%d"),
        "*Posting Date":                    combined["Posting Date"].dt.strftime("%Y%m%d"),
        "*Transaction Currency (3)":        "USD",
        "*IC Service Transaction Type (5)": 0,
        "Document Header Text (25)":        "XXX.GF.WW.RC FA GAIN LOSS",
        "Ledger Group (4)":                 "-",
        "Company Code (4)":                 combined["Company Code"],
        "G/L Account (10)":                 combined["account_number"],
        "Transaction Currency Debit":       combined["Transaction Currency Debit"].abs(),
        "Transaction Currency Credit":      combined["Transaction Currency Credit"].abs(),
        "Item Text (50)":                   combined["JE Item Text"],
        "Amount in Company Code Currency":  "-",
        "Cost Center (10)":                 combined["Costcenter"],
        "Profit Center (10)":               combined["Profit Center"],
        "Trading Partner (6)":              "-",
        "Customer (10)":                    np.where(is_interco, INTERCO_CUSTOMER, ""),
        "Product number (40)":              np.where(
                                                is_interco,
                                                "PC_" + combined["Profit Center"].astype(str) + "_Dummy",
                                                "",
                                            ),
    })


def _apply_header_style(worksheet) -> None:
    fill = PatternFill(start_color=HEADER_COLOR, end_color=HEADER_COLOR, fill_type="solid")
    font = Font(color="000000", bold=True)
    for cell in worksheet[1]:
        cell.fill = fill
        cell.font = font


# ── Public entry point ────────────────────────────────────────────────────────

def process_reclass(input_folder: str, output_folder: str) -> tuple[str, str]:
    """
    Full pipeline: read → transform → write.

    Returns:
        (output_path, output_filename)
    """
    input_file = find_excel_file(input_folder)
    df = load_and_prepare(input_file)

    group1 = build_group1(df)
    group2 = build_group2(df)

    combined = pd.concat([group1, group2], ignore_index=True)
    combined = combined.sort_values(
        by=[
            "Company Code", "Posting Date", "JE Item Text", "Profit Center",
            "Costcenter", "Net Reclass Amount", "Functional Area", "account_number",
        ],
        ascending=True,
        ignore_index=True,
    )

    output_df = build_output_df(combined)

    os.makedirs(output_folder, exist_ok=True)
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    input_base  = os.path.splitext(os.path.basename(input_file))[0]
    output_filename = f"JE_{input_base}_{timestamp}.xlsx"
    output_path     = os.path.join(output_folder, output_filename)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False, sheet_name="JE_Output")
        _apply_header_style(writer.sheets["JE_Output"])

    return output_path, output_filename

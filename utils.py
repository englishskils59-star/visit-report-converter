# utils.py
# WDI Visit Analytics Engine
# Utility functions: Arabic text handling, column validation, shared helpers

import pandas as pd
import numpy as np
from datetime import datetime, date
import re

# ─────────────────────────────────────────────
# REQUIRED COLUMNS
# ─────────────────────────────────────────────

REQUIRED_COLUMNS = [
    "Year",
    "Month",
    "Visit Date",
    "Customer Name",
    "Customer Category",
    "Governorate",
    "District",
    "Visit Notes",
    "Total Visit Flag",
    "Unique Customer Flag",
    "Sales Rep Name",
    "Current Customer",
    "Target Customer",
    "Potential Customer",
    "New Customer",
    "Not Interested Customer",
    "Former Customer",
]

# ─────────────────────────────────────────────
# COLUMN ALIASES (Arabic or alternate names)
# ─────────────────────────────────────────────

COLUMN_ALIASES = {
    "السنة": "Year",
    "الشهر": "Month",
    "تاريخ الزيارة": "Visit Date",
    "اسم العميل": "Customer Name",
    "فئة العميل": "Customer Category",
    "المحافظة": "Governorate",
    "المنطقة": "District",
    "ملاحظات الزيارة": "Visit Notes",
    "إجمالي الزيارات": "Total Visit Flag",
    "عميل فريد": "Unique Customer Flag",
    "اسم المندوب": "Sales Rep Name",
    "عميل حالي": "Current Customer",
    "عميل مستهدف": "Target Customer",
    "عميل محتمل": "Potential Customer",
    "عميل جديد": "New Customer",
    "غير مهتم": "Not Interested Customer",
    "عميل سابق": "Former Customer",
}

# ─────────────────────────────────────────────
# CUSTOMER STATUS LABELS
# ─────────────────────────────────────────────

STATUS_LABELS = {
    "current": "Current Customer",
    "potential": "Potential Customer",
    "target": "Target Customer",
    "new": "New Customer",
    "former": "Former Customer",
    "not_interested": "Not Interested",
    "unclassified": "Unclassified",
}

STATUS_COLORS = {
    "Current Customer":    "#70AD47",
    "Potential Customer":  "#2E75B6",
    "Target Customer":     "#FFC000",
    "New Customer":        "#1F4E79",
    "Former Customer":     "#A9A9A9",
    "Not Interested":      "#C00000",
    "Unclassified":        "#D9D9D9",
}

# ─────────────────────────────────────────────
# ARABIC TEXT HELPERS
# ─────────────────────────────────────────────

import re

def normalize_arabic(text: str) -> str:

    if not isinstance(text, str):
        return ""

    text = text.strip()

    # Alef normalization
    text = re.sub(r"[إأآا]", "ا", text)

    # Ta Marbuta
    text = re.sub(r"ة", "ه", text)

    # Ya
    text = re.sub(r"[يى]", "ي", text)

    # Hamza forms
    text = re.sub(r"ؤ", "و", text)
    text = re.sub(r"ئ", "ي", text)

    # Tatweel
    text = re.sub(r"ـ", "", text)

    # Remove extra spaces
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def contains_arabic(text: str) -> bool:
    """Return True if text contains Arabic characters."""
    if not isinstance(text, str):
        return False
    return bool(re.search(r"[\u0600-\u06FF]", text))


def safe_str(val) -> str:
    """Convert any value to a clean string; return empty string for NaN/None."""
    if val is None:
        return ""
    if isinstance(val, float) and np.isnan(val):
        return ""
    return str(val).strip()


# ─────────────────────────────────────────────
# FILE LOADING
# ─────────────────────────────────────────────

def load_excel(uploaded_file) -> tuple[pd.DataFrame | None, str]:
    """
    Load an uploaded Excel file into a DataFrame.
    Returns (df, error_message). error_message is empty string on success.
    """
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        # Rename columns using alias map
        df = df.rename(columns=COLUMN_ALIASES)
        return df, ""
    except Exception as e:
        return None, f"Failed to read file: {e}"


# ─────────────────────────────────────────────
# COLUMN VALIDATION
# ─────────────────────────────────────────────

def validate_columns(df: pd.DataFrame) -> tuple[bool, list[str], list[str]]:
    """
    Validate that required columns are present.
    Returns (is_valid, missing_cols, present_cols).
    """
    present = list(df.columns)
    missing = [c for c in REQUIRED_COLUMNS if c not in present]
    is_valid = len(missing) == 0
    return is_valid, missing, present


# ─────────────────────────────────────────────
# DATE PARSING
# ─────────────────────────────────────────────

def parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Parse and clean the Visit Date column."""
    if "Visit Date" not in df.columns:
        return df
    df = df.copy()
    df["Visit Date"] = pd.to_datetime(df["Visit Date"], errors="coerce")
    return df


# ─────────────────────────────────────────────
# DATA CLEANING
# ─────────────────────────────────────────────

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply baseline cleaning:
    - Strip string columns
    - Fill NaN in text columns with empty string
    - Parse dates
    - Sort by Visit Date ascending
    """
    df = df.copy()

    text_cols = [
        "Customer Name", "Customer Category", "Governorate",
        "District", "Visit Notes", "Sales Rep Name",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(safe_str)

    flag_cols = [
        "Current Customer", "Target Customer", "Potential Customer",
        "New Customer", "Not Interested Customer", "Former Customer",
        "Total Visit Flag", "Unique Customer Flag",
    ]
    for col in flag_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df = parse_dates(df)
    df = df.sort_values("Visit Date", ascending=True).reset_index(drop=True)
    return df


# ─────────────────────────────────────────────
# SUMMARY STATS
# ─────────────────────────────────────────────

def basic_stats(df: pd.DataFrame) -> dict:
    """Return a dict of basic dataset statistics."""
    stats = {
        "total_records": len(df),
        "unique_customers": df["Customer Name"].nunique() if "Customer Name" in df.columns else 0,
        "unique_reps": df["Sales Rep Name"].nunique() if "Sales Rep Name" in df.columns else 0,
        "date_range_start": df["Visit Date"].min() if "Visit Date" in df.columns else None,
        "date_range_end": df["Visit Date"].max() if "Visit Date" in df.columns else None,
        "governorates": df["Governorate"].nunique() if "Governorate" in df.columns else 0,
    }
    return stats


# ─────────────────────────────────────────────
# DATE HELPERS
# ─────────────────────────────────────────────

def days_since(last_date, reference_date=None) -> int | None:
    """Return number of days between last_date and reference_date (today if None)."""
    if reference_date is None:
        reference_date = pd.Timestamp(datetime.today().date())
    if pd.isnull(last_date):
        return None
    delta = reference_date - pd.Timestamp(last_date)
    return delta.days


def months_list(df: pd.DataFrame) -> list[str]:
    """Return sorted list of 'YYYY-MM' strings from Visit Date column."""
    if "Visit Date" not in df.columns:
        return []
    dates = df["Visit Date"].dropna()
    months = dates.dt.to_period("M").astype(str).unique().tolist()
    return sorted(months)


# ─────────────────────────────────────────────
# NUMBER FORMATTING
# ─────────────────────────────────────────────

def fmt_number(n) -> str:
    """Format an integer with comma separators."""
    try:
        return f"{int(n):,}"
    except Exception:
        return str(n)


def fmt_pct(n, decimals=1) -> str:
    """Format a float as a percentage string."""
    try:
        return f"{float(n):.{decimals}f}%"
    except Exception:
        return "—"
# ─────────────────────────────────────────────
# Customer Name Cleaning
# ─────────────────────────────────────────────
REMOVE_WORDS = [
    "الحاج",
    "الحاجه",
    "م/",
    "ا/",
    "أ/",
    "د/",
    "مزرعه",
    "مزرعة",
    "عنبر",
    "مكتب",
    "فرع"
]

def normalize_customer_name(name):

    name = normalize_arabic(name)

    for word in REMOVE_WORDS:
        name = name.replace(word, "")

    name = re.sub(r"\s+", " ", name)

    return name.strip()


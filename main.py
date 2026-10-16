"""
Data Cleaning & Reporting Automation
Run: python main.py
This script cleans data/data file and writes cleaned data plus Excel and chart reports.
"""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

ROOT = Path(__file__).resolve().parent
DATA_PATH = ROOT / "data" / "sales_data_raw.csv"
OUTPUT_DIR = ROOT / "outputs"
REPORT_DIR = ROOT / "reports"

def load_data(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    return pd.read_csv(path)

def clean_data(df: pd.DataFrame):
    original_rows = len(df)
    df = df.copy()
    # Remove fully duplicated rows
    before_duplicates = len(df)
    df = df.drop_duplicates()
    duplicates_removed = before_duplicates - len(df)

    # Standardize text fields
    for col in ["Order_ID", "Customer", "Region", "Product", "Salesperson", "Email"]:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()

    # Standardize region names
    region_map = {
        "north": "North", "south": "South", "east": "East", "west": "West"
    }
    df["Region"] = df["Region"].str.lower().map(region_map).fillna("Unknown")

    # Convert data types; invalid dates become missing
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Unit_Price"] = pd.to_numeric(df["Unit_Price"], errors="coerce")

    # Flag invalid values before replacing missing values
    invalid_quantity_count = int((df["Quantity"] <= 0).sum())
    df.loc[df["Quantity"] <= 0, "Quantity"] = pd.NA
    df.loc[df["Unit_Price"] < 0, "Unit_Price"] = pd.NA

    missing_before = int(df.isna().sum().sum())

    # Fill numeric missing values with median, dates with mode, text with "Not Provided"
    for col in ["Quantity", "Unit_Price"]:
        median_value = df[col].median()
        df[col] = df[col].fillna(median_value if pd.notna(median_value) else 0)
    if df["Date"].notna().any():
        date_mode = df["Date"].mode().iloc[0]
        df["Date"] = df["Date"].fillna(date_mode)
    else:
        df["Date"] = df["Date"].fillna(pd.Timestamp.today().normalize())

    for col in ["Customer", "Email", "Salesperson"]:
        df[col] = df[col].replace({"<NA>": pd.NA, "nan": pd.NA}).fillna("Not Provided")

    # Clean data should have positive numeric values
    df["Quantity"] = df["Quantity"].astype(int)
    df["Unit_Price"] = df["Unit_Price"].round(2)
    df["Sales"] = df["Quantity"] * df["Unit_Price"]
    df["Date"] = df["Date"].dt.strftime("%Y-%m-%d")

    stats = {
        "Rows in raw data": original_rows,
        "Duplicate rows removed": duplicates_removed,
        "Invalid/non-positive quantities corrected": invalid_quantity_count,
        "Missing cells filled or corrected": missing_before,
        "Rows in cleaned data": len(df),
        "Total sales (₹)": round(float(df["Sales"].sum()), 2),
    }
    return df, stats

def create_reports(cleaned: pd.DataFrame, stats: dict):
    OUTPUT_DIR.mkdir(exist_ok=True)
    REPORT_DIR.mkdir(exist_ok=True)
    cleaned.to_csv(OUTPUT_DIR / "cleaned_sales_data.csv", index=False)

    # Summary tables
    region_summary = (cleaned.groupby("Region", as_index=False)
                      .agg(Orders=("Order_ID", "count"), Total_Sales=("Sales", "sum"))
                      .sort_values("Total_Sales", ascending=False))
    product_summary = (cleaned.groupby("Product", as_index=False)
                        .agg(Units_Sold=("Quantity", "sum"), Total_Sales=("Sales", "sum"))
                        .sort_values("Total_Sales", ascending=False))
    daily_summary = (cleaned.groupby("Date", as_index=False)
                     .agg(Total_Sales=("Sales", "sum"))
                     .sort_values("Date"))

    stats_df = pd.DataFrame({"Metric": list(stats.keys()), "Value": list(stats.values())})
    with pd.ExcelWriter(OUTPUT_DIR / "automated_sales_report.xlsx", engine="openpyxl") as writer:
        stats_df.to_excel(writer, sheet_name="Data Quality", index=False)
        region_summary.to_excel(writer, sheet_name="Region Summary", index=False)
        product_summary.to_excel(writer, sheet_name="Product Summary", index=False)
        daily_summary.to_excel(writer, sheet_name="Daily Sales", index=False)
        cleaned.to_excel(writer, sheet_name="Cleaned Data", index=False)

    # Sales by region chart
    plt.figure(figsize=(8, 5))
    plt.bar(region_summary["Region"], region_summary["Total_Sales"])
    plt.title("Total Sales by Region")
    plt.xlabel("Region")
    plt.ylabel("Sales (₹)")
    plt.tight_layout()
    plt.savefig(REPORT_DIR / "sales_by_region.png", dpi=160)
    plt.close()

    # Sales by product chart
    plt.figure(figsize=(8, 5))
    plt.bar(product_summary["Product"], product_summary["Total_Sales"])
    plt.title("Total Sales by Product")
    plt.xlabel("Product")
    plt.ylabel("Sales (₹)")
    plt.xticks(rotation=20, ha="right")
    plt.tight_layout()
    plt.savefig(REPORT_DIR / "sales_by_product.png", dpi=160)
    plt.close()

def main():
    raw = load_data(DATA_PATH)
    cleaned, stats = clean_data(raw)
    create_reports(cleaned, stats)
    print("Automation completed successfully.")
    print(f"Cleaned CSV: {OUTPUT_DIR / 'cleaned_sales_data.csv'}")
    print(f"Excel report: {OUTPUT_DIR / 'automated_sales_report.xlsx'}")
    print(f"Charts: {REPORT_DIR}")

if __name__ == "__main__":
    main()

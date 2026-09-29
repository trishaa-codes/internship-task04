"""
Data Cleaning & Reporting Automation
------------------------------------
Place a CSV named raw_sales_data.csv inside the data/ folder and run:

    python automate_report.py

The script cleans text, numeric values and dates, removes duplicate Order_IDs,
fills missing numeric/date values, calculates Revenue, creates summary tables,
charts and an Excel report.
"""

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
REPORTS = BASE / "reports"
CHARTS = BASE / "charts"
REPORTS.mkdir(exist_ok=True)
CHARTS.mkdir(exist_ok=True)

input_file = DATA / "raw_sales_data.csv"
output_file = DATA / "cleaned_sales_data.csv"
excel_file = REPORTS / "automated_data_cleaning_report.xlsx"

df = pd.read_csv(input_file)
original_rows = len(df)

# Standardize text fields
text_cols = ["Customer_Name", "Email", "City", "Product", "Category", "Payment_Status"]
for col in text_cols:
    df[col] = df[col].astype("string").str.strip()

df["Customer_Name"] = df["Customer_Name"].str.replace(r"\s+", " ", regex=True)
df["City"] = df["City"].str.title()
df["Category"] = df["Category"].str.title()
df["Payment_Status"] = df["Payment_Status"].str.title()

# Convert numeric/currency fields
df["Unit_Price"] = (
    df["Unit_Price"].astype("string")
    .str.replace("₹", "", regex=False)
    .str.replace(",", "", regex=False)
    .str.strip()
)
df["Unit_Price"] = pd.to_numeric(df["Unit_Price"], errors="coerce")
df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")

# Convert dates and make invalid dates missing
df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")

# Remove duplicate orders
duplicates_removed = int(df.duplicated(subset=["Order_ID"], keep="first").sum())
df = df.drop_duplicates(subset=["Order_ID"], keep="first").copy()

# Fill missing values using robust defaults
missing_values_filled = int(df.isna().sum().sum())
df["Quantity"] = df["Quantity"].fillna(df["Quantity"].median()).round().astype(int)
df["Unit_Price"] = df["Unit_Price"].fillna(df["Unit_Price"].median())
df["Order_Date"] = df["Order_Date"].fillna(df["Order_Date"].dropna().median())

# Derived metric
df["Revenue"] = df["Quantity"] * df["Unit_Price"]
df.to_csv(output_file, index=False)

# Summaries
city = df.groupby("City", as_index=False).agg(
    Orders=("Order_ID", "count"), Revenue=("Revenue", "sum")
).sort_values("Revenue", ascending=False)

category = df.groupby("Category", as_index=False).agg(
    Orders=("Order_ID", "count"), Revenue=("Revenue", "sum")
).sort_values("Revenue", ascending=False)

payment = df["Payment_Status"].value_counts().rename_axis(
    "Payment_Status"
).reset_index(name="Orders")

daily = df.groupby(df["Order_Date"].dt.date, as_index=False).agg(
    Orders=("Order_ID", "count"), Revenue=("Revenue", "sum")
)

# Excel report
kpi = pd.DataFrame({
    "Metric": [
        "Original Rows", "Duplicate Rows Removed", "Missing Values Filled",
        "Final Rows", "Total Revenue", "Average Order Value"
    ],
    "Value": [
        original_rows, duplicates_removed, missing_values_filled,
        len(df), df["Revenue"].sum(), df["Revenue"].mean()
    ]
})

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    kpi.to_excel(writer, sheet_name="Dashboard", index=False)
    city.to_excel(writer, sheet_name="City Summary", index=False)
    category.to_excel(writer, sheet_name="Category Summary", index=False)
    payment.to_excel(writer, sheet_name="Payment Summary", index=False)
    daily.to_excel(writer, sheet_name="Daily Trend", index=False)
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)

# Visual summaries
plt.figure(figsize=(9, 5))
plt.bar(city["City"], city["Revenue"])
plt.title("Revenue by City")
plt.xticks(rotation=45, ha="right")
plt.ylabel("Revenue")
plt.tight_layout()
plt.savefig(CHARTS / "revenue_by_city.png", dpi=180)
plt.close()

plt.figure(figsize=(9, 5))
plt.plot(daily["Order_Date"].astype(str), daily["Revenue"], marker="o")
plt.title("Daily Revenue Trend")
plt.xticks(rotation=45, ha="right")
plt.ylabel("Revenue")
plt.tight_layout()
plt.savefig(CHARTS / "daily_revenue_trend.png", dpi=180)
plt.close()

print("Automation completed successfully.")
print(f"Original rows: {original_rows}")
print(f"Duplicate rows removed: {duplicates_removed}")
print(f"Final rows: {len(df)}")
print(f"Total revenue: ₹{df['Revenue'].sum():,.2f}")
print(f"Excel report: {excel_file}")

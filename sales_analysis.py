"""
Module 4 Capstone - Global Sales Data Cleaning, Exploration and Aggregation
===========================================================================
Input : Sample_data.xlsx  (700 transactions, Sep 2013 - Dec 2014)
Output: cleaned CSV, summary tables (CSV) and charts (PNG) in the output folder

Usage : python sales_analysis.py [path_to_excel] [output_folder]
Needs : pandas, openpyxl, matplotlib
"""
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")                      # no display needed
import matplotlib.pyplot as plt
import pandas as pd

INPUT = sys.argv[1] if len(sys.argv) > 1 else "Sample_data.xlsx"
OUT = Path(sys.argv[2] if len(sys.argv) > 2 else "analysis_output")
OUT.mkdir(parents=True, exist_ok=True)
pd.options.display.float_format = "{:,.2f}".format
pd.options.display.width = 200

# ---------------------------------------------------------------
# 1. LOAD
# ---------------------------------------------------------------
df = pd.read_excel(INPUT)
print("Raw shape:", df.shape)

# ---------------------------------------------------------------
# 2. CLEANING
# ---------------------------------------------------------------
# 2.1 Trim header whitespace (the raw file has " Sales" with a leading space)
df.columns = df.columns.str.strip()

# 2.2 Trim text values and make sure labels are consistent
for col in ["Segment", "Country", "Product", "Discount Band", "Month Name"]:
    df[col] = df[col].astype("string").str.strip()

# 2.3 Missing values
print("\nMissing values before cleaning:\n", df.isna().sum()[df.isna().sum() > 0])
# Discount Band is blank only where Discounts == 0, i.e. "no discount" (not missing data)
blank_band = df["Discount Band"].isna()
assert (df.loc[blank_band, "Discounts"] == 0).all(), "Blank band rows should have zero discount"
df["Discount Band"] = df["Discount Band"].fillna("None")
assert df.isna().sum().sum() == 0, "Unexpected missing values remain"

# 2.4 Data types
df["Date"] = pd.to_datetime(df["Date"])
money_cols = ["Gross Sales", "Discounts", "Sales", "COGS", "Profit"]
df[money_cols + ["Units Sold"]] = df[money_cols + ["Units Sold"]].astype(float)
for col in ["Segment", "Country", "Product", "Month Name"]:
    df[col] = df[col].astype("category")
df["Discount Band"] = pd.Categorical(
    df["Discount Band"], categories=["None", "Low", "Medium", "High"], ordered=True)

# 2.5 Duplicates
print("Duplicate rows:", df.duplicated().sum())
df = df.drop_duplicates()

# 2.6 Logic / integrity checks (tolerance of 1 dollar for rounding)
checks = {
    "Gross Sales - Discounts = Sales": (df["Gross Sales"] - df["Discounts"] - df["Sales"]).abs() > 1,
    "Sales - COGS = Profit":           (df["Sales"] - df["COGS"] - df["Profit"]).abs() > 1,
    "Units x Sale Price = Gross Sales": (df["Units Sold"] * df["Sale Price"] - df["Gross Sales"]).abs() > 1,
    "Month Number matches Date":       df["Date"].dt.month != df["Month Number"],
    "Year matches Date":               df["Date"].dt.year != df["Year"],
}
print("\nIntegrity checks (number of failing rows):")
for name, bad in checks.items():
    print(f"  {name}: {int(bad.sum())}")

# 2.7 Flags (kept in the data, not removed)
print("\nRows with fractional Units Sold:", int((df["Units Sold"] % 1 != 0).sum()))
print("Rows with negative Profit      :", int((df["Profit"] < 0).sum()),
      "-> segments:", df.loc[df["Profit"] < 0, "Segment"].unique().tolist())

# 2.8 Derived columns
df["Profit Margin"] = df["Profit"] / df["Sales"]
df["Period"] = df["Date"].dt.to_period("M").astype(str)

df.to_csv(OUT / "cleaned_data.csv", index=False)

# ---------------------------------------------------------------
# 3. EXPLORATION
# ---------------------------------------------------------------
print("\nDate range:", df["Date"].min().date(), "to", df["Date"].max().date())
print("\nDescriptive statistics:\n", df[["Units Sold"] + money_cols].describe().T)

# ---------------------------------------------------------------
# 4. KEY METRICS
# ---------------------------------------------------------------
total = df[["Units Sold"] + money_cols].sum()
kpis = {
    "Total Revenue (Sales)": total["Sales"],
    "Gross Sales": total["Gross Sales"],
    "Total Discounts": total["Discounts"],
    "Total COGS": total["COGS"],
    "Total Profit": total["Profit"],
    "Units Sold": total["Units Sold"],
    "Profit Margin": total["Profit"] / total["Sales"],
    "Average Discount Rate": total["Discounts"] / total["Gross Sales"],
}
print("\nKEY METRICS")
for k, v in kpis.items():
    print(f"  {k:<24}", f"{v:.1%}" if "Margin" in k or "Rate" in k else f"{v:,.0f}")
pd.Series(kpis, name="Value").to_csv(OUT / "kpis.csv")

# ---------------------------------------------------------------
# 5. AGGREGATIONS
# ---------------------------------------------------------------
def summarise(by):
    g = (df.groupby(by, observed=True)
           .agg(Units=("Units Sold", "sum"), Gross_Sales=("Gross Sales", "sum"),
                Discounts=("Discounts", "sum"), Revenue=("Sales", "sum"),
                COGS=("COGS", "sum"), Profit=("Profit", "sum"),
                Transactions=("Sales", "size")))
    g["Margin"] = g["Profit"] / g["Revenue"]
    g["Discount_Rate"] = g["Discounts"] / g["Gross_Sales"]
    g["Pct_Revenue"] = g["Revenue"] / g["Revenue"].sum()
    g["Pct_Profit"] = g["Profit"] / g["Profit"].sum()
    return g.sort_values("Revenue", ascending=False)

by_segment = summarise("Segment")
by_country = summarise("Country")
by_product = summarise("Product")
by_band = summarise("Discount Band").sort_index()
by_seg_cty = df.pivot_table(index="Segment", columns="Country", values="Profit",
                            aggfunc="sum", observed=True, margins=True, margins_name="Total")
monthly = summarise("Period").sort_index()
loss_rows = df[df["Profit"] < 0].groupby(["Segment", "Country"], observed=True)["Profit"].agg(["count", "sum"])

tables = {"by_segment": by_segment, "by_country": by_country, "by_product": by_product,
          "by_discount_band": by_band, "profit_segment_x_country": by_seg_cty,
          "monthly_trend": monthly, "loss_making_rows": loss_rows}
for name, t in tables.items():
    t.to_csv(OUT / f"{name}.csv")
    print(f"\n--- {name.upper().replace('_', ' ')} ---\n", t)

# Segment x discount band (shows where discounting destroys margin)
seg_band = (df.groupby(["Segment", "Discount Band"], observed=True)
              .agg(Revenue=("Sales", "sum"), Profit=("Profit", "sum")))
seg_band["Margin"] = seg_band["Profit"] / seg_band["Revenue"]
print("\n--- SEGMENT x DISCOUNT BAND ---\n", seg_band)
seg_band.to_csv(OUT / "segment_x_discount_band.csv")

# Like-for-like growth: Sep-Dec 2014 vs Sep-Dec 2013 (2013 only has Sep-Dec)
sd = df[df["Month Number"] >= 9].groupby("Year")[["Sales", "Profit"]].sum()
growth = sd.loc[2014] / sd.loc[2013] - 1
print("\nSep-Dec growth 2014 vs 2013:  Revenue {:.1%}  Profit {:.1%}".format(
    growth["Sales"], growth["Profit"]))

# ---------------------------------------------------------------
# 6. CHARTS
# ---------------------------------------------------------------
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9,
                     "axes.spines.top": False, "axes.spines.right": False})
NAVY, TEAL, RED = "#1F3864", "#2E9CCA", "#C0392B"

def money_bar(table, title, fname):
    fig, ax = plt.subplots(figsize=(6.4, 3.2))
    t = table[["Revenue", "Profit"]] / 1e6
    t.plot(kind="bar", ax=ax, color=[NAVY, TEAL], width=0.75)
    ax.set_title(title, fontweight="bold", loc="left"); ax.set_xlabel("")
    ax.set_ylabel("USD millions"); ax.axhline(0, color="black", lw=0.6)
    plt.xticks(rotation=20, ha="right"); ax.legend(frameon=False)
    fig.tight_layout(); fig.savefig(OUT / fname, dpi=200); plt.close(fig)

money_bar(by_segment, "Revenue and Profit by Segment", "chart_segment.png")
money_bar(by_country, "Revenue and Profit by Country", "chart_country.png")

fig, ax = plt.subplots(figsize=(6.4, 3.0))
vals = by_band["Margin"] * 100
ax.bar(vals.index.astype(str), vals.values, color=[TEAL, TEAL, NAVY, RED])
for i, v in enumerate(vals.values):
    ax.text(i, v + 0.4, f"{v:.1f}%", ha="center", fontsize=9)
ax.set_title("Profit Margin by Discount Band", fontweight="bold", loc="left")
ax.set_ylabel("Margin (%)"); fig.tight_layout()
fig.savefig(OUT / "chart_discount_band.png", dpi=200); plt.close(fig)

fig, ax = plt.subplots(figsize=(6.4, 3.0))
ax.plot(monthly.index, monthly["Revenue"] / 1e6, marker="o", color=NAVY, label="Revenue")
ax.plot(monthly.index, monthly["Profit"] / 1e6, marker="o", color=TEAL, label="Profit")
ax.set_title("Monthly Revenue and Profit", fontweight="bold", loc="left")
ax.set_ylabel("USD millions"); plt.xticks(rotation=60, ha="right", fontsize=7)
ax.legend(frameon=False); fig.tight_layout()
fig.savefig(OUT / "chart_monthly.png", dpi=200); plt.close(fig)

print(f"\nDone. Files saved in: {OUT.resolve()}")

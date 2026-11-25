"""
Mutual Fund Overview & Insights
Cleaning, scoring and Top 30 shortlisting of mutual funds (Pandas, scikit-learn).

Usage:
    python mutual_fund_analysis.py --data comprehensive_mutual_funds_data.csv --out outputs

Outputs: clean_funds.csv, Top30_Funds.xlsx, insights.txt and PNG charts.
"""
import argparse
import os

import numpy as np
import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.preprocessing import MinMaxScaler
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

WEIGHTS = {"return": 0.45, "expense": 0.25, "age": 0.15, "consistency": 0.15}
IDEAL_AGE = 7
ALT_WEIGHTS = [
    (0.35, 0.30, 0.20, 0.15),
    (0.60, 0.20, 0.10, 0.10),
    (0.40, 0.40, 0.10, 0.10),
]
TEXT_COLS = ["sortino", "alpha", "sd", "beta", "sharpe"]


def load_data(path):
    if not os.path.exists(path):
        try:
            from google.colab import files
            print(f"{path} not found. Please upload it.")
            files.upload()
        except ImportError:
            raise SystemExit(f"File not found: {path}")
    df = pd.read_csv(path)
    print(f"Loaded {df.shape[0]} rows x {df.shape[1]} columns")
    return df


def clean(df):
    df = df.copy()
    print("\n[Stage 1] Missing values before cleaning:")
    print(df.isna().sum()[df.isna().sum() > 0].to_string())

    for c in TEXT_COLS:
        df[c] = pd.to_numeric(df[c].replace("-", np.nan), errors="coerce")

    for c in ["scheme_name", "fund_manager", "amc_name", "category", "sub_category"]:
        df[c] = df[c].astype(str).str.strip()

    before = len(df)
    df = df.dropna(subset=["returns_3yr"]).reset_index(drop=True)
    print(f"Dropped {before - len(df)} rows with missing returns_3yr -> {len(df)} rows left")

    print(f"Unique scheme names: {df['scheme_name'].nunique()} of {len(df)} rows")
    print(f"Exact duplicate rows: {df.duplicated().sum()}")

    df.insert(0, "fund_id", df.index + 1)
    return df


def describe(df, out):
    num = ["returns_1yr", "returns_3yr", "expense_ratio",
           "fund_age_yr", "fund_size_cr", "risk_level"]
    summary = df[num].agg(["mean", "median", "min", "max", "std"]).T
    summary["mode"] = df[num].mode().iloc[0]
    print("\n[Stage 2] Summary statistics:")
    print(summary.round(2).to_string())

    fig, axes = plt.subplots(2, 3, figsize=(15, 8))
    for ax, c in zip(axes.ravel(), num):
        sns.histplot(df[c], kde=True, ax=ax)
        ax.set_title(c)
    plt.tight_layout()
    plt.savefig(f"{out}/stage2_distributions.png", dpi=120)
    plt.close()

    plt.figure(figsize=(8, 2.5))
    sns.boxplot(x=df["returns_1yr"])
    plt.title("1-year return (outlier check)")
    plt.tight_layout()
    plt.savefig(f"{out}/stage2_returns_1yr_boxplot.png", dpi=120)
    plt.close()

    print("\nTop 5 1-year returns (outlier check):")
    print(df.nlargest(5, "returns_1yr")[["scheme_name", "category",
          "returns_1yr", "returns_3yr"]].to_string(index=False))
    return summary


def score_funds(df, weights=None):
    w = weights or (WEIGHTS["return"], WEIGHTS["expense"],
                    WEIGHTS["age"], WEIGHTS["consistency"])
    return (w[0] * df["ret3_n"]
            + w[1] * (1 - df["exp_n"])
            + w[2] * df["age_score"]
            + w[3] * df["consistency"])


def normalise_and_rank(df):
    df = df.copy()
    scaler = MinMaxScaler()
    df[["ret3_n", "exp_n"]] = scaler.fit_transform(df[["returns_3yr", "expense_ratio"]])

    span = max(IDEAL_AGE - df["fund_age_yr"].min(), df["fund_age_yr"].max() - IDEAL_AGE)
    df["age_score"] = 1 - (df["fund_age_yr"] - IDEAL_AGE).abs() / span
    df["consistency"] = (df["returns_1yr"] > 0).astype(int)

    df["score"] = score_funds(df)
    df = df.sort_values("score", ascending=False).reset_index(drop=True)
    df.insert(0, "rank", df.index + 1)

    print("\n[Stage 3/4] Weights:", WEIGHTS, "| ideal age:", IDEAL_AGE)
    print("Correlation of score with its components:")
    print(df[["score", "ret3_n", "exp_n", "age_score", "consistency"]]
          .corr()["score"].drop("score").round(2).to_string())

    base = set(df.head(30)["fund_id"])
    print("\nSensitivity: Top-30 overlap with other weightings")
    for w in ALT_WEIGHTS:
        alt = set(df.loc[score_funds(df, w).nlargest(30).index, "fund_id"])
        print(f"  weights {w}: {len(base & alt)} of 30 stay")
    return df


def export_top30(df, out):
    cols = ["rank", "scheme_name", "category", "sub_category", "amc_name",
            "risk_level", "fund_age_yr", "expense_ratio",
            "returns_1yr", "returns_3yr", "score"]
    top = df.head(30)[cols].copy()
    top["score"] = top["score"].round(4)
    path = f"{out}/Top30_Funds.xlsx"
    top.to_excel(path, index=False, sheet_name="Top 30")

    wb = load_workbook(path)
    ws = wb.active
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "A2"
    for col in ws.columns:
        width = max(len(str(c.value or "")) for c in col) + 2
        ws.column_dimensions[col[0].column_letter].width = min(45, width)
    wb.save(path)

    print("\n[Stage 5] Top 30 traits")
    print("Category:", top["category"].value_counts().to_dict())
    print("Risk level:", top["risk_level"].value_counts().to_dict())
    print("Top AMCs:", top["amc_name"].value_counts().head(4).to_dict())
    print("Top sub-categories:", top["sub_category"].value_counts().head(3).to_dict())

    plt.figure(figsize=(8, 5))
    top.groupby("category")["score"].count().sort_values().plot.barh(color="#1f6feb")
    plt.title("Top 30 funds by category")
    plt.tight_layout()
    plt.savefig(f"{out}/stage5_top30_by_category.png", dpi=120)
    plt.close()
    return top


def insights(df, out):
    lines = []
    aum = df.groupby("category")["fund_size_cr"].sum().sort_values(ascending=False)
    lines.append("1. AUM by fund type (Rs crore): " + aum.round(0).astype(int).to_dict().__str__())
    lines.append(f"   Largest: {aum.index[0]}, ahead of {aum.index[1]} by Rs {aum.iloc[0]-aum.iloc[1]:,.0f} crore")

    mgr = df.groupby("fund_manager")["fund_size_cr"].sum().nlargest(3)
    lines.append("2. Top fund managers by AUM: " + mgr.round(0).astype(int).to_dict().__str__())

    exp_cat = df.groupby("category")["expense_ratio"].mean().round(2)
    sub = df.groupby("sub_category")["expense_ratio"].agg(["mean", "count"]).round(2).sort_values("mean")
    lines.append("3. Avg expense ratio by category: " + exp_cat.to_dict().__str__())
    lines.append("   Cheapest sub-categories:\n" + sub.head(3).to_string())

    best = df.loc[df["returns_1yr"].idxmax()]
    lines.append(f"4. Best 1-year return: {best['scheme_name']} at {best['returns_1yr']}% "
                 f"(its 3-year return is {best['returns_3yr']}%)")

    lines.append(f"5. Average min SIP: {df['min_sip'].mean():.2f} (median {df['min_sip'].median():.0f}); "
                 f"average min lumpsum: {df['min_lumpsum'].mean():,.0f} "
                 f"(median {df['min_lumpsum'].median():,.0f}, mode {df['min_lumpsum'].mode()[0]:,.0f})")

    cat = df.groupby("category").agg(avg_3yr_return=("returns_3yr", "mean"),
                                     avg_risk_level=("risk_level", "mean"),
                                     avg_std_dev=("sd", "mean")).round(2)
    lines.append("6. Return vs risk by category:\n" + cat.sort_values("avg_3yr_return", ascending=False).to_string())

    amc = df.groupby("amc_name").agg(avg_3yr=("returns_3yr", "mean"),
                                     aum=("fund_size_cr", "sum"),
                                     funds=("fund_id", "count")).round(2)
    lines.append("   Top AMCs by average 3-year return (note the fund counts):\n"
                 + amc.nlargest(10, "avg_3yr").to_string())

    text = "\n".join(lines)
    print("\n[Stage 7] Insights\n" + text)
    with open(f"{out}/insights.txt", "w", encoding="utf-8") as f:
        f.write(text)

    cat["avg_3yr_return"].sort_values().plot.barh(color="#2da44e", figsize=(8, 4))
    plt.title("Average 3-year return by category (%)")
    plt.tight_layout()
    plt.savefig(f"{out}/stage7_returns_by_category.png", dpi=120)
    plt.close()


def main():
    ap = argparse.ArgumentParser(description="Mutual fund analysis")
    ap.add_argument("--data", default="comprehensive_mutual_funds_data.csv")
    ap.add_argument("--out", default="outputs")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    raw = load_data(args.data)
    df = clean(raw)
    describe(df, args.out)
    ranked = normalise_and_rank(df)
    export_top30(ranked, args.out)
    insights(ranked, args.out)

    ranked.drop(columns=["ret3_n", "exp_n", "age_score", "consistency", "score", "rank"]) \
          .sort_values("fund_id").to_csv(f"{args.out}/clean_funds.csv", index=False)
    print(f"\nDone. Files saved in '{args.out}/'")


if __name__ == "__main__":
    main()

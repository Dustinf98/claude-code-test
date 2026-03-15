import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import pandas as pd


def _debit_series(df):
    """Return a numeric Series of debit amounts, excluding Payment rows."""
    df = df.copy()
    df["Debit"] = pd.to_numeric(df["Debit"], errors="coerce").fillna(0)
    if "Category" in df.columns:
        df = df[df["Category"] != "Payment"]
    return df


def category_bar(df):
    df = _debit_series(df)
    totals = df.groupby("Category")["Debit"].sum().sort_values(ascending=False)
    totals = totals[totals > 0]

    fig, ax = plt.subplots(figsize=(9, 5))
    bars = ax.bar(totals.index, totals.values, color="steelblue")
    ax.set_title("Spending by Category")
    ax.set_ylabel("Amount ($)")
    ax.set_xlabel("Category")
    plt.xticks(rotation=35, ha="right")

    for bar, val in zip(bars, totals.values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                f"${val:.0f}", ha="center", va="bottom", fontsize=8)

    fig.tight_layout()
    return fig


def monthly_trend(all_df):
    df = _debit_series(all_df)
    df["YearMonth"] = pd.to_datetime(df["Date"]).dt.to_period("M").astype(str)
    totals = df.groupby("YearMonth")["Debit"].sum().sort_index()

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(totals.index, totals.values, marker="o", color="steelblue", linewidth=2)
    ax.set_title("Monthly Spending Trend")
    ax.set_ylabel("Total Spent ($)")
    ax.set_xlabel("Month")
    plt.xticks(rotation=35, ha="right")

    for x, y in zip(totals.index, totals.values):
        ax.annotate(f"${y:.0f}", (x, y), textcoords="offset points",
                    xytext=(0, 8), ha="center", fontsize=8)

    fig.tight_layout()
    return fig


def top_merchants(df, n=10):
    df = _debit_series(df)
    totals = df.groupby("Description")["Debit"].sum().nlargest(n).sort_values()

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.barh(totals.index, totals.values, color="steelblue")
    ax.set_title(f"Top {n} Merchants by Spending")
    ax.set_xlabel("Amount ($)")

    for i, val in enumerate(totals.values):
        ax.text(val + 0.5, i, f"${val:.0f}", va="center", fontsize=8)

    fig.tight_layout()
    return fig

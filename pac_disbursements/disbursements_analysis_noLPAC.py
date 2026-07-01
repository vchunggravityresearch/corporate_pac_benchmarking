"""
analyze_pac_disbursements.py

Takes one or more *cleaned* PAC disbursement CSVs (the output of
clean_pac_disbursements.py -- must have committee_name, disbursement_amount,
disbursement_year, disbursement_quarter, and candidate_party columns) and
analyzes how disbursements differ:

  1. By party, across all companies (which company's PAC leans most
     Democratic/Republican/Independent in its giving, both in dollars and
     in share of their total disbursements).
  2. By year, across all companies (total giving trend per company per
     year, plus an all-companies-combined total per year).
  3. By quarter, across all companies (same, broken out by year-quarter so
     seasonality/cyclicality across election cycles is visible).

Multiple cleaned files can be selected at once (e.g. one per company/PAC)
-- they get stacked together and the company identifier used for grouping
is each row's own committee_name column (e.g. "ABBVIE POLITICAL ACTION
COMMITTEE"), not the filename, so it works correctly even if a filename
doesn't clearly indicate which PAC it is.

OUTPUTS
-------
Everything is written to an output folder (default: "pac_disbursement_analysis"
next to the first input file):
  - by_party.csv              company x party $ totals, transaction counts, % of company total
  - by_year.csv                company x year $ totals, transaction counts
  - by_quarter.csv             company x year-quarter $ totals, transaction counts
  - totals_by_party.csv        all-companies-combined $ totals by party
  - totals_by_year.csv         all-companies-combined $ totals by year
  - totals_by_quarter.csv      all-companies-combined $ totals by year-quarter
  - chart_by_party.png         stacked bar: $ disbursed by company, broken out by party
  - chart_by_year.png          grouped bar: $ disbursed by company, broken out by year
  - chart_by_quarter.png       line chart: $ disbursed over time (year-quarter), one line per company

USAGE
-----
    # No arguments needed -- a file picker opens, supporting multi-select
    python analyze_pac_disbursements.py

    # Or specify files directly:
    python analyze_pac_disbursements.py \
        --input Abbvie_cleaned.csv Pfizer_cleaned.csv Comcast_cleaned.csv

    # Send output somewhere specific instead of the default folder:
    python analyze_pac_disbursements.py \
        --input Abbvie_cleaned.csv Pfizer_cleaned.csv \
        --output-dir "/path/to/some/folder"
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

REQUIRED_COLUMNS = [
    "committee_name",
    "disbursement_amount",
    "disbursement_year",
    "disbursement_quarter",
    "candidate_party",
]


# ---------------------------------------------------------------------------
# File selection helpers (same pattern as clean_pac_disbursements.py)
# ---------------------------------------------------------------------------
def pick_input_files() -> list[str]:
    """Open a native file-picker dialog that allows selecting multiple
    cleaned disbursement CSVs at once. Falls back to a typed,
    comma-separated prompt if no GUI is available."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        paths = filedialog.askopenfilenames(
            title="Select one or more CLEANED PAC disbursement CSVs to analyze",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        root.destroy()
        return list(paths)
    except Exception as e:
        print(f"  [info] file picker unavailable ({e}); falling back to typed input", file=sys.stderr)
        typed = input("Path(s) to cleaned CSV(s), comma-separated: ").strip()
        return [p.strip().strip('"') for p in typed.split(",") if p.strip()]


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------
def load_and_combine(input_paths: list[str]) -> pd.DataFrame:
    frames = []
    for p in input_paths:
        df = pd.read_csv(p, low_memory=False)
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            print(
                f"  [skip] {p} is missing required column(s) {missing} -- "
                "is this a CLEANED file from clean_pac_disbursements.py?",
                file=sys.stderr,
            )
            continue
        df["disbursement_amount"] = pd.to_numeric(df["disbursement_amount"], errors="coerce")
        frames.append(df)

    if not frames:
        print("No usable input files (none had the required columns). Exiting.", file=sys.stderr)
        sys.exit(1)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.dropna(subset=["disbursement_amount"])

    n_companies = combined["committee_name"].nunique()
    print(f"  [info] loaded {len(combined):,} rows across {len(frames)} file(s), {n_companies} distinct committee(s)", file=sys.stderr)
    return combined


# ---------------------------------------------------------------------------
# Analysis: by party
# ---------------------------------------------------------------------------
def analyze_by_party(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    grouped = (
        df.groupby(["committee_name", "candidate_party"])["disbursement_amount"]
        .agg(total_amount="sum", transaction_count="count")
        .reset_index()
    )
    company_totals = grouped.groupby("committee_name")["total_amount"].sum()
    grouped["pct_of_company_total"] = grouped.apply(
        lambda r: round(100 * r["total_amount"] / company_totals[r["committee_name"]], 1)
        if company_totals[r["committee_name"]] else 0,
        axis=1,
    )
    grouped = grouped.sort_values(["committee_name", "total_amount"], ascending=[True, False])

    totals_by_party = (
        df.groupby("candidate_party")["disbursement_amount"]
        .agg(total_amount="sum", transaction_count="count")
        .reset_index()
        .sort_values("total_amount", ascending=False)
    )
    grand_total = totals_by_party["total_amount"].sum()
    totals_by_party["pct_of_grand_total"] = round(100 * totals_by_party["total_amount"] / grand_total, 1)

    return grouped, totals_by_party


# ---------------------------------------------------------------------------
# Analysis: by year
# ---------------------------------------------------------------------------
def analyze_by_year(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    yr = df.dropna(subset=["disbursement_year"]).copy()
    yr["disbursement_year"] = yr["disbursement_year"].astype(int)

    grouped = (
        yr.groupby(["committee_name", "disbursement_year"])["disbursement_amount"]
        .agg(total_amount="sum", transaction_count="count")
        .reset_index()
        .sort_values(["committee_name", "disbursement_year"])
    )

    totals_by_year = (
        yr.groupby("disbursement_year")["disbursement_amount"]
        .agg(total_amount="sum", transaction_count="count")
        .reset_index()
        .sort_values("disbursement_year")
    )

    return grouped, totals_by_year


# ---------------------------------------------------------------------------
# Analysis: by quarter (year-quarter, so it stays chronological across years)
# ---------------------------------------------------------------------------
def analyze_by_quarter(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    q = df.dropna(subset=["disbursement_year", "disbursement_quarter"]).copy()
    q["disbursement_year"] = q["disbursement_year"].astype(int)
    q["year_quarter"] = q["disbursement_year"].astype(str) + " " + q["disbursement_quarter"].astype(str)

    grouped = (
        q.groupby(["committee_name", "disbursement_year", "disbursement_quarter", "year_quarter"])[
            "disbursement_amount"
        ]
        .agg(total_amount="sum", transaction_count="count")
        .reset_index()
        .sort_values(["committee_name", "disbursement_year", "disbursement_quarter"])
    )

    totals_by_quarter = (
        q.groupby(["disbursement_year", "disbursement_quarter", "year_quarter"])["disbursement_amount"]
        .agg(total_amount="sum", transaction_count="count")
        .reset_index()
        .sort_values(["disbursement_year", "disbursement_quarter"])
    )

    return grouped, totals_by_quarter


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
PARTY_COLORS = {
    "Democratic": "#1f77b4",
    "Republican": "#d62728",
    "Independent": "#2ca02c",
    "NON-CANDIDATE / LEADERSHIP PAC": "#7f7f7f",
}


def make_charts(
    by_party: pd.DataFrame,
    totals_by_party: pd.DataFrame,
    by_year: pd.DataFrame,
    totals_by_year: pd.DataFrame,
    by_quarter: pd.DataFrame,
    totals_by_quarter: pd.DataFrame,
    output_dir: Path,
) -> None:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker

    plt.rcParams.update({"figure.autolayout": True})
    dollar_fmt = mticker.FuncFormatter(lambda x, _: f"${x:,.0f}")
    pct_fmt = mticker.FuncFormatter(lambda x, _: f"{x:.0f}%")

    # ---- helpers ----
    def party_color_list(parties):
        default = plt.rcParams["axes.prop_cycle"].by_key()["color"]
        out, idx = [], 0
        for p in parties:
            if p in PARTY_COLORS:
                out.append(PARTY_COLORS[p])
            else:
                out.append(default[idx % len(default)])
                idx += 1
        return out

    # =========================================================
    # PER-COMPANY CHARTS
    # =========================================================

    # --- Chart 1: by party, stacked bars per company ---
    pivot_party = by_party.pivot_table(
        index="committee_name", columns="candidate_party", values="total_amount", fill_value=0
    )
    colors = party_color_list(pivot_party.columns.tolist())
    fig, ax = plt.subplots(figsize=(max(8, len(pivot_party) * 1.3), 6))
    pivot_party.plot(kind="bar", stacked=True, ax=ax, color=colors)
    ax.set_title("Disbursements by Company — Broken Out by Party")
    ax.set_xlabel("Company / PAC")
    ax.set_ylabel("Total Disbursed ($)")
    ax.yaxis.set_major_formatter(dollar_fmt)
    ax.legend(title="Party", bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.xticks(rotation=30, ha="right")
    fig.savefig(output_dir / "chart_by_party.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    # --- Chart 2: by year, grouped bars per company ---
    pivot_year = by_year.pivot_table(
        index="disbursement_year", columns="committee_name", values="total_amount", fill_value=0
    )
    fig, ax = plt.subplots(figsize=(max(8, len(pivot_year) * 1.3), 6))
    pivot_year.plot(kind="bar", ax=ax)
    ax.set_title("Disbursements by Company — Broken Out by Year")
    ax.set_xlabel("Year")
    ax.set_ylabel("Total Disbursed ($)")
    ax.yaxis.set_major_formatter(dollar_fmt)
    ax.legend(title="Company / PAC", bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.xticks(rotation=0)
    fig.savefig(output_dir / "chart_by_year.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    # --- Chart 3: by quarter, one line per company over time ---
    pivot_quarter = by_quarter.pivot_table(
        index="year_quarter", columns="committee_name", values="total_amount", fill_value=0
    )
    order = (
        by_quarter[["disbursement_year", "disbursement_quarter", "year_quarter"]]
        .drop_duplicates()
        .sort_values(["disbursement_year", "disbursement_quarter"])["year_quarter"]
        .tolist()
    )
    pivot_quarter = pivot_quarter.reindex(order)

    fig, ax = plt.subplots(figsize=(max(10, len(order) * 0.6), 6))
    for company in pivot_quarter.columns:
        ax.plot(pivot_quarter.index, pivot_quarter[company], marker="o", label=company)
    ax.set_title("Disbursements Over Time by Quarter — by Company")
    ax.set_xlabel("Year-Quarter")
    ax.set_ylabel("Total Disbursed ($)")
    ax.yaxis.set_major_formatter(dollar_fmt)
    ax.legend(title="Company / PAC", bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.xticks(rotation=45, ha="right")
    fig.savefig(output_dir / "chart_by_quarter.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    # =========================================================
    # SUMMARY CHARTS  (all companies combined)
    # =========================================================

    # --- Summary Chart 1: party — side-by-side $ bar + % donut ---
    parties = totals_by_party["candidate_party"].tolist()
    amounts = totals_by_party["total_amount"].tolist()
    pcts = totals_by_party["pct_of_grand_total"].tolist()
    s_colors = party_color_list(parties)

    fig, (ax_bar, ax_pie) = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle("All-PACs Combined — Disbursements by Party", fontsize=13, fontweight="bold")

    ax_bar.barh(parties[::-1], amounts[::-1], color=s_colors[::-1])
    ax_bar.set_xlabel("Total Disbursed ($)")
    ax_bar.xaxis.set_major_formatter(dollar_fmt)
    ax_bar.set_title("Dollar totals")
    for i, (amt, pct) in enumerate(zip(amounts[::-1], pcts[::-1])):
        ax_bar.text(amt * 1.01, i, f"  {pct:.1f}%", va="center", fontsize=9)

    wedges, texts, autotexts = ax_pie.pie(
        amounts, labels=None, colors=s_colors,
        autopct="%1.1f%%", startangle=140, pctdistance=0.75,
        wedgeprops={"linewidth": 0.8, "edgecolor": "white"},
    )
    for at in autotexts:
        at.set_fontsize(9)
    ax_pie.legend(wedges, parties, title="Party", loc="lower center",
                  bbox_to_anchor=(0.5, -0.18), ncol=2, fontsize=8)
    ax_pie.set_title("Share of total giving")

    fig.savefig(output_dir / "summary_chart_by_party.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    # --- Summary Chart 2: total disbursements by year (bar) ---
    years = totals_by_year["disbursement_year"].astype(int).tolist()
    yr_amounts = totals_by_year["total_amount"].tolist()

    fig, ax = plt.subplots(figsize=(max(7, len(years) * 1.0), 5))
    fig.suptitle("All-PACs Combined — Total Disbursements by Year", fontsize=13, fontweight="bold")
    bars = ax.bar([str(y) for y in years], yr_amounts, color="#5c85d6", edgecolor="white")
    ax.set_xlabel("Year")
    ax.set_ylabel("Total Disbursed ($)")
    ax.yaxis.set_major_formatter(dollar_fmt)
    ax.set_title("Dollar totals across all companies")
    for bar, amt in zip(bars, yr_amounts):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() * 1.01,
                f"${amt:,.0f}", ha="center", va="bottom", fontsize=8)
    plt.xticks(rotation=0)
    fig.savefig(output_dir / "summary_chart_by_year.png", dpi=150, bbox_inches="tight")
    plt.close(fig)

    # --- Summary Chart 3: total disbursements by quarter (line + area) ---
    q_order = (
        totals_by_quarter[["disbursement_year", "disbursement_quarter", "year_quarter"]]
        .drop_duplicates()
        .sort_values(["disbursement_year", "disbursement_quarter"])["year_quarter"]
        .tolist()
    )
    q_amounts = (
        totals_by_quarter.set_index("year_quarter")
        .reindex(q_order)["total_amount"]
        .tolist()
    )

    fig, ax = plt.subplots(figsize=(max(10, len(q_order) * 0.7), 5))
    fig.suptitle("All-PACs Combined — Total Disbursements by Quarter", fontsize=13, fontweight="bold")
    ax.fill_between(range(len(q_order)), q_amounts, alpha=0.15, color="#5c85d6")
    ax.plot(range(len(q_order)), q_amounts, marker="o", color="#5c85d6", linewidth=2)
    ax.set_xticks(range(len(q_order)))
    ax.set_xticklabels(q_order, rotation=45, ha="right")
    ax.set_xlabel("Year-Quarter")
    ax.set_ylabel("Total Disbursed ($)")
    ax.yaxis.set_major_formatter(dollar_fmt)
    ax.set_title("Dollar totals across all companies")
    # annotate every other tick to avoid crowding
    for i, (lbl, amt) in enumerate(zip(q_order, q_amounts)):
        if i % 2 == 0:
            ax.annotate(f"${amt:,.0f}", (i, amt), textcoords="offset points",
                        xytext=(0, 8), ha="center", fontsize=7.5)
    fig.savefig(output_dir / "summary_chart_by_quarter.png", dpi=150, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        nargs="*",
        default=None,
        help="Path(s) to one or more CLEANED PAC disbursement CSVs. If omitted, a file-picker dialog opens (supports multi-select).",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Folder to save analysis outputs into. Defaults to a 'pac_disbursement_analysis' subfolder created next to the first input file.",
    )
    parser.add_argument(
        "--no-charts",
        action="store_true",
        help="Skip generating PNG charts (only write the CSV summary tables).",
    )
    args = parser.parse_args()

    input_paths = args.input if args.input else pick_input_files()
    if not input_paths:
        print("No input file(s) selected. Exiting.", file=sys.stderr)
        sys.exit(1)

    missing_inputs = [p for p in input_paths if not Path(p).exists()]
    if missing_inputs:
        print(f"Input file(s) not found: {missing_inputs}", file=sys.stderr)
        sys.exit(1)

    output_dir = Path(args.output_dir) if args.output_dir else (
        Path(input_paths[0]).resolve().parent / "pac_disbursement_analysis"
    )
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {output_dir}", file=sys.stderr)

    df = load_and_combine(input_paths)

    by_party, totals_by_party = analyze_by_party(df)
    by_year, totals_by_year = analyze_by_year(df)
    by_quarter, totals_by_quarter = analyze_by_quarter(df)

    by_party.to_csv(output_dir / "by_party.csv", index=False)
    totals_by_party.to_csv(output_dir / "totals_by_party.csv", index=False)
    by_year.to_csv(output_dir / "by_year.csv", index=False)
    totals_by_year.to_csv(output_dir / "totals_by_year.csv", index=False)
    by_quarter.drop(columns=["year_quarter"]).to_csv(output_dir / "by_quarter.csv", index=False)
    totals_by_quarter.drop(columns=["year_quarter"]).to_csv(output_dir / "totals_by_quarter.csv", index=False)

    if not args.no_charts:
        try:
            make_charts(by_party, totals_by_party, by_year, totals_by_year, by_quarter, totals_by_quarter, output_dir)
        except ImportError:
            print(
                "  [warn] matplotlib not installed -- skipping charts (pip install matplotlib --break-system-packages). "
                "CSV summaries were still written.",
                file=sys.stderr,
            )

    print("\n=== Totals by party (all companies combined) ===")
    print(totals_by_party.to_string(index=False))
    print("\n=== Totals by year (all companies combined) ===")
    print(totals_by_year.to_string(index=False))
    print("\n=== Totals by year-quarter (all companies combined) ===")
    print(totals_by_quarter.to_string(index=False))
    print(f"\nFull breakdowns and charts written to: {output_dir}")


if __name__ == "__main__":
    main()
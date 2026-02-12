from __future__ import annotations

import re
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


EXCEL_FILE = Path("Grad Program Exit Survey Data (2).xlsx")
OUTPUT_DIR = Path("outputs")


def detect_sheet_and_year(excel_path: Path) -> tuple[str, str]:
    xls = pd.ExcelFile(excel_path, engine="openpyxl")
    preferred = [s for s in xls.sheet_names if "exit survey" in s.lower()]
    sheet_name = preferred[0] if preferred else xls.sheet_names[0]

    year_match = re.search(r"(20\d{2})", sheet_name)
    year = year_match.group(1) if year_match else "unknown_year"
    return sheet_name, year


def extract_rank_columns(columns: list[str], group: str) -> list[str]:
    group_lower = group.lower()
    selected = []

    for col in columns:
        if not isinstance(col, str):
            continue
        col_lower = col.lower()
        has_ranks = "ranks" in col_lower
        has_rank_marker = (" - rank" in col_lower) or col_lower.endswith("- rank")
        in_group = group_lower in col_lower
        is_did_not_take = "did not take" in col_lower

        if has_ranks and has_rank_marker and in_group and not is_did_not_take:
            selected.append(col)

    return selected


def course_name_from_header(header: str) -> str:
    parts = re.split(r"\s-\sRank\s*$", header, flags=re.IGNORECASE)
    raw = parts[0] if parts else header
    raw = re.sub(r"\bRanks\b", "", raw, flags=re.IGNORECASE)
    return " ".join(raw.split()).strip(" -")


def build_ranking(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    if not columns:
        return pd.DataFrame(columns=["course_name", "n_responses", "avg_rank", "median_rank", "borda_points"])

    k = len(columns)
    records = []

    for col in columns:
        numeric_ranks = pd.to_numeric(df[col], errors="coerce")
        valid_for_borda = numeric_ranks.where((numeric_ranks >= 1) & (numeric_ranks <= k))

        n_responses = int(numeric_ranks.notna().sum())
        avg_rank = float(numeric_ranks.mean()) if n_responses else float("nan")
        median_rank = float(numeric_ranks.median()) if n_responses else float("nan")
        borda_points = float((k - valid_for_borda + 1).sum(skipna=True))

        records.append(
            {
                "course_name": course_name_from_header(col),
                "n_responses": n_responses,
                "avg_rank": avg_rank,
                "median_rank": median_rank,
                "borda_points": borda_points,
            }
        )

    ranking = pd.DataFrame(records)
    ranking = ranking.sort_values(
        by=["avg_rank", "median_rank", "n_responses"],
        ascending=[True, True, False],
        na_position="last",
    ).reset_index(drop=True)

    return ranking


def save_plot(ranking: pd.DataFrame, group: str, year: str, out_path: Path) -> None:
    if ranking.empty:
        return

    fig_height = max(4, len(ranking) * 0.45)
    fig, ax = plt.subplots(figsize=(12, fig_height))

    y_positions = list(range(len(ranking)))
    bars = ax.barh(y_positions, ranking["avg_rank"], color="#4C78A8")
    ax.set_yticks(y_positions)
    ax.set_yticklabels(ranking["course_name"])
    ax.invert_yaxis()

    ax.set_xlabel("Average Rank (lower is better)")
    ax.set_title(f"{group} Course Ranking by Average Rank ({year})")

    xmax = ranking["avg_rank"].max(skipna=True)
    if pd.notna(xmax):
        ax.set_xlim(0, xmax * 1.25)

    for bar, avg, n_resp in zip(bars, ranking["avg_rank"], ranking["n_responses"]):
        if pd.isna(avg):
            continue
        ax.text(
            bar.get_width() + 0.03,
            bar.get_y() + bar.get_height() / 2,
            f"avg={avg:.2f}, n={int(n_resp)}",
            va="center",
            fontsize=9,
        )

    fig.tight_layout()
    fig.savefig(out_path, dpi=200)
    plt.close(fig)


def process_group(df: pd.DataFrame, group: str, year: str) -> None:
    rank_cols = extract_rank_columns([str(c) for c in df.columns], group=group)
    print(f"{group} rank columns found: {len(rank_cols)}")

    if not rank_cols:
        print(f"No {group} ranking columns detected; skipping outputs for this group.")
        return

    ranking = build_ranking(df, rank_cols)
    ranking_path = OUTPUT_DIR / f"{group.lower()}_ranking_{year}.csv"
    figure_path = OUTPUT_DIR / f"{group.lower()}_ranking_{year}.png"

    ranking.to_csv(ranking_path, index=False)
    save_plot(ranking, group, year, figure_path)

    print(f"Top 10 {group} items:")
    print(ranking[["course_name", "avg_rank", "n_responses"]].head(10).to_string(index=False))


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    sheet_name, year = detect_sheet_and_year(EXCEL_FILE)
    print(f"Detected sheet: {sheet_name}")
    print(f"Detected year: {year}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, engine="openpyxl")

    process_group(df, group="CORE", year=year)
    process_group(df, group="ELECTIVE", year=year)


if __name__ == "__main__":
    main()

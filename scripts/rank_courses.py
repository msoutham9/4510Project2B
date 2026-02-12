from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
import pandas as pd

EXCEL_FILE = Path("Grad Program Exit Survey Data (2).xlsx")
OUTPUT_DIR = Path("outputs")
FIGURE_PATH = OUTPUT_DIR / "rank_order.png"


def detect_sheet(excel_path: Path) -> str:
    xls = pd.ExcelFile(excel_path, engine="openpyxl")
    sheet_names = xls.sheet_names

    preferred = [name for name in sheet_names if "exit survey" in name.lower()]
    if preferred:
        return preferred[0]

    year_like = [name for name in sheet_names if re.search(r"20\d{2}", name)]
    if year_like:
        return year_like[0]

    return sheet_names[0]


def detect_year(sheet_name: str) -> str:
    match = re.search(r"(20\d{2})", sheet_name)
    return match.group(1) if match else sheet_name


def clean_item_name(header: str) -> str:
    item = re.sub(r"\s*-\s*Rank\s*$", "", header, flags=re.IGNORECASE)
    item = re.sub(r"\bRanks\b", "", item, flags=re.IGNORECASE)
    item = re.sub(r"\s+", " ", item).strip(" -")
    return item


def detect_rank_columns(columns: Iterable[object]) -> list[str]:
    rank_columns: list[str] = []
    for col in columns:
        if not isinstance(col, str):
            continue
        lower = col.lower()
        has_ranks = "ranks" in lower
        has_rank = ("- rank" in lower) or (" - rank" in lower)
        if has_ranks and has_rank and "did not take" not in lower:
            rank_columns.append(col)
    return rank_columns


def infer_group_name(header: str) -> str:
    lowered = header.lower()
    if "core" in lowered:
        return "core"
    if "elective" in lowered:
        return "elective"
    if "program" in lowered:
        return "program"
    if "course" in lowered:
        return "course"
    return "all"


def compute_ranking_table(df: pd.DataFrame, rank_columns: list[str]) -> pd.DataFrame:
    records: list[dict[str, float | int | str]] = []
    k = len(rank_columns)

    for col in rank_columns:
        series = df[col].replace({"Did not take": pd.NA})
        numeric = pd.to_numeric(series, errors="coerce")
        valid = numeric.where((numeric >= 1) & (numeric <= k))

        n_responses = int(numeric.notna().sum())
        avg_rank = float(numeric.mean()) if n_responses else float("nan")
        median_rank = float(numeric.median()) if n_responses else float("nan")
        borda_points = float((k - valid + 1).sum(skipna=True))

        records.append(
            {
                "item_name": clean_item_name(col),
                "n_responses": n_responses,
                "avg_rank": avg_rank,
                "median_rank": median_rank,
                "borda_points": borda_points,
            }
        )

    ranking = pd.DataFrame.from_records(records)
    ranking = ranking.sort_values(
        by=["avg_rank", "median_rank", "n_responses", "item_name"],
        ascending=[True, True, False, True],
        na_position="last",
        kind="mergesort",
    ).reset_index(drop=True)
    ranking["rank_position"] = ranking.index + 1
    return ranking


def save_group_csvs(df: pd.DataFrame, rank_columns: list[str]) -> list[tuple[str, pd.DataFrame, Path]]:
    grouped: dict[str, list[str]] = {}
    for col in rank_columns:
        group = infer_group_name(col)
        grouped.setdefault(group, []).append(col)

    # Use grouped output only when there is more than one meaningful group.
    non_all_groups = [name for name in grouped if name != "all"]
    use_grouped = len(non_all_groups) >= 2

    outputs: list[tuple[str, pd.DataFrame, Path]] = []
    if use_grouped:
        for group_name in sorted(non_all_groups):
            table = compute_ranking_table(df, grouped[group_name])
            out_path = OUTPUT_DIR / f"rank_order_{group_name}.csv"
            table.to_csv(out_path, index=False)
            outputs.append((group_name, table, out_path))
    else:
        table = compute_ranking_table(df, rank_columns)
        out_path = OUTPUT_DIR / "rank_order.csv"
        table.to_csv(out_path, index=False)
        outputs.append(("all", table, out_path))

    return outputs


def save_figure(group_tables: list[tuple[str, pd.DataFrame, Path]], sheet_name: str, year: str) -> None:
    n_groups = len(group_tables)
    fig_height = max(4.5, 0.45 * sum(len(table) for _, table, _ in group_tables) + (n_groups - 1))

    fig, axes = plt.subplots(n_groups, 1, figsize=(13, fig_height), squeeze=False)
    axes_flat = axes.flatten()

    for ax, (group_name, table, _) in zip(axes_flat, group_tables):
        y_positions = list(range(len(table)))
        bars = ax.barh(y_positions, table["avg_rank"], color="#4C78A8")
        ax.set_yticks(y_positions)
        ax.set_yticklabels(table["item_name"])
        ax.invert_yaxis()

        label = "Average Rank (lower is better)"
        ax.set_xlabel(label)
        ax.set_title(f"{group_name.upper()} ranking" if group_name != "all" else "Ranking")

        xmax = table["avg_rank"].max(skipna=True)
        if pd.notna(xmax):
            ax.set_xlim(0, xmax * 1.25)

        for bar, avg, n_resp in zip(bars, table["avg_rank"], table["n_responses"]):
            if pd.isna(avg):
                continue
            ax.text(
                bar.get_width() + 0.02,
                bar.get_y() + bar.get_height() / 2,
                f"avg={avg:.2f}, n={int(n_resp)}",
                va="center",
                fontsize=8.5,
            )

    fig.suptitle(f"Rank order of courses/programs ({year} | {sheet_name})", fontsize=13)
    fig.tight_layout(rect=[0, 0, 1, 0.97])
    fig.savefig(FIGURE_PATH, dpi=220)
    plt.close(fig)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    sheet_name = detect_sheet(EXCEL_FILE)
    year = detect_year(sheet_name)
    print(f"Detected sheet name: {sheet_name}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, engine="openpyxl")
    rank_columns = detect_rank_columns(df.columns)
    print(f"Number of rank columns found: {len(rank_columns)}")

    if not rank_columns:
        raise ValueError("No rank columns were detected using the required header rules.")

    group_tables = save_group_csvs(df, rank_columns)

    overall = compute_ranking_table(df, rank_columns)
    print("Top 10 ranked items (item_name, avg_rank, n_responses):")
    print(overall[["item_name", "avg_rank", "n_responses"]].head(10).to_string(index=False))

    for _, _, path in group_tables:
        print(f"Saved CSV: {path}")

    save_figure(group_tables, sheet_name=sheet_name, year=year)
    print(f"Saved figure: {FIGURE_PATH}")


if __name__ == "__main__":
    main()

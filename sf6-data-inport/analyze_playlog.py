from __future__ import annotations

import argparse
import codecs
import csv
import math
from pathlib import Path
from typing import Any

HEADER_ROW_INDEX = 2
DATA_START_ROW_INDEX = 4
MASTER_LP_THRESHOLD = 25000
EXCLUDED_FEATURES = {
    "No",
    "データ取得日",
    "プレイヤー名",
    "リーグポイント",
    "ランク",
    "MR",
}


def detect_text_encoding(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(codecs.BOM_UTF8):
        return "utf-8-sig"
    try:
        raw.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        return "cp932"


def load_playlog_rows(path: Path) -> tuple[list[str], list[dict[str, Any]]]:
    encoding = detect_text_encoding(path)
    with path.open("r", encoding=encoding, newline="") as file:
        raw_rows = list(csv.reader(file))

    if len(raw_rows) <= DATA_START_ROW_INDEX:
        raise ValueError("CSVのデータ行が不足しています")

    group_labels = raw_rows[1][:]
    item_labels = raw_rows[HEADER_ROW_INDEX][:]
    columns = build_column_names(group_labels, item_labels)

    rows: list[dict[str, Any]] = []
    for raw_row in raw_rows[DATA_START_ROW_INDEX:]:
        if not raw_row or not any(cell.strip() for cell in raw_row):
            continue

        padded = raw_row + [""] * (len(columns) - len(raw_row))
        row = {column: padded[index].strip() for index, column in enumerate(columns)}
        row["No"] = parse_numeric(row["No"])
        row["リーグポイント"] = parse_numeric(row["リーグポイント"])
        row["MR"] = parse_numeric(row["MR"])
        rows.append(row)

    return columns, rows


def build_column_names(group_labels: list[str], item_labels: list[str]) -> list[str]:
    counts: dict[str, int] = {}
    for item in item_labels[1:]:
        key = item.strip()
        counts[key] = counts.get(key, 0) + 1

    columns = ["No"]
    for index in range(1, len(item_labels)):
        item = item_labels[index].strip()
        group = group_labels[index].strip() if index < len(group_labels) else ""

        if counts.get(item, 0) > 1 and group and group != "プレイヤー名情報":
            columns.append(f"{group}_{item}")
        else:
            columns.append(item)

    return columns


def parse_numeric(value: str) -> float | None:
    text = value.strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1]
    try:
        return float(text)
    except ValueError:
        return None


def get_feature_names(columns: list[str]) -> list[str]:
    return [column for column in columns if column not in EXCLUDED_FEATURES]


def mean(values: list[float]) -> float:
    return sum(values) / len(values)


def sample_std(values: list[float]) -> float:
    if len(values) < 2:
        return 0.0
    avg = mean(values)
    variance = sum((value - avg) ** 2 for value in values) / (len(values) - 1)
    return math.sqrt(variance)


def pearson_correlation(xs: list[float], ys: list[float]) -> float:
    if len(xs) != len(ys) or len(xs) < 2:
        return 0.0

    mean_x = mean(xs)
    mean_y = mean(ys)
    std_x = sample_std(xs)
    std_y = sample_std(ys)
    if std_x == 0 or std_y == 0:
        return 0.0

    covariance = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys)) / (len(xs) - 1)
    return covariance / (std_x * std_y)


def simple_regression(xs: list[float], ys: list[float]) -> tuple[float, float, float]:
    mean_x = mean(xs)
    mean_y = mean(ys)
    variance_x = sum((x - mean_x) ** 2 for x in xs)
    if variance_x == 0:
        return 0.0, mean_y, 0.0

    covariance_xy = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys))
    slope = covariance_xy / variance_x
    intercept = mean_y - slope * mean_x
    correlation = pearson_correlation(xs, ys)
    return slope, intercept, correlation**2


def analyze_segment(
    rows: list[dict[str, Any]],
    feature_names: list[str],
    target_name: str,
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []

    for feature in feature_names:
        xs: list[float] = []
        ys: list[float] = []
        for row in rows:
            x = parse_numeric(str(row.get(feature, "")))
            y = row.get(target_name)
            if x is None or not isinstance(y, (int, float)):
                continue
            xs.append(float(x))
            ys.append(float(y))

        if len(xs) < 3:
            continue

        correlation = pearson_correlation(xs, ys)
        slope, intercept, r_squared = simple_regression(xs, ys)
        results.append(
            {
                "feature": feature,
                "n": len(xs),
                "mean": mean(xs),
                "correlation": correlation,
                "slope": slope,
                "intercept": intercept,
                "r_squared": r_squared,
            }
        )

    return sorted(results, key=lambda item: abs(item["correlation"]), reverse=True)


def build_progress_score(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    submaster_targets = [
        float(row["リーグポイント"]) for row in rows if is_submaster_row(row) and isinstance(row["リーグポイント"], (int, float))
    ]
    master_targets = [
        float(row["MR"]) for row in rows if is_master_row(row) and isinstance(row["MR"], (int, float))
    ]

    submaster_mean = mean(submaster_targets) if submaster_targets else 0.0
    submaster_std = sample_std(submaster_targets)
    master_mean = mean(master_targets) if master_targets else 0.0
    master_std = sample_std(master_targets)

    scored_rows: list[dict[str, Any]] = []
    for row in rows:
        scored = dict(row)
        if is_submaster_row(row) and isinstance(row["リーグポイント"], (int, float)):
            if submaster_std == 0:
                scored["進捗スコア"] = 0.0
            else:
                scored["進捗スコア"] = (float(row["リーグポイント"]) - submaster_mean) / submaster_std
        elif is_master_row(row) and isinstance(row["MR"], (int, float)):
            if master_std == 0:
                scored["進捗スコア"] = 0.0
            else:
                scored["進捗スコア"] = (float(row["MR"]) - master_mean) / master_std
        else:
            scored["進捗スコア"] = None
        scored_rows.append(scored)

    return scored_rows


def is_submaster_row(row: dict[str, Any]) -> bool:
    lp = row.get("リーグポイント")
    return isinstance(lp, (int, float)) and lp < MASTER_LP_THRESHOLD


def is_master_row(row: dict[str, Any]) -> bool:
    lp = row.get("リーグポイント")
    mr = row.get("MR")
    return isinstance(lp, (int, float)) and lp >= MASTER_LP_THRESHOLD and isinstance(mr, (int, float))


def print_section(title: str, results: list[dict[str, Any]], top_n: int = 8) -> None:
    print(f"\n[{title}]")
    if not results:
        print("分析に十分なデータがありません")
        return

    print("上位の正相関")
    for item in [result for result in results if result["correlation"] > 0][:top_n]:
        print(
            f"  + {item['feature']}: r={item['correlation']:.3f}, "
            f"slope={item['slope']:.3f}, R^2={item['r_squared']:.3f}, n={item['n']}"
        )

    print("上位の負相関")
    negatives = [result for result in results if result["correlation"] < 0]
    negatives = sorted(negatives, key=lambda item: item["correlation"])[:top_n]
    for item in negatives:
        print(
            f"  - {item['feature']}: r={item['correlation']:.3f}, "
            f"slope={item['slope']:.3f}, R^2={item['r_squared']:.3f}, n={item['n']}"
        )


def write_results_csv(path: Path, results: list[dict[str, Any]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(
            file,
            fieldnames=["feature", "n", "correlation", "slope", "intercept", "r_squared"],
        )
        writer.writeheader()
        writer.writerows(results)


def main() -> None:
    parser = argparse.ArgumentParser(description="sf6-playlog-out.csv の要因分析を行う")
    parser.add_argument("--input", default="sf6-playlog-out.csv", help="分析対象CSV")
    parser.add_argument(
        "--out-dir",
        default="analysis-output",
        help="分析結果CSVの出力先フォルダ",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    columns, rows = load_playlog_rows(input_path)
    feature_names = get_feature_names(columns)

    submaster_rows = [row for row in rows if is_submaster_row(row)]
    master_rows = [row for row in rows if is_master_row(row)]
    progress_rows = build_progress_score(rows)

    submaster_results = analyze_segment(submaster_rows, feature_names, "リーグポイント")
    master_results = analyze_segment(master_rows, feature_names, "MR")
    progress_results = analyze_segment(progress_rows, feature_names, "進捗スコア")

    print(f"総件数: {len(rows)}")
    print(f"LP 25000 未満: {len(submaster_rows)}件")
    print(f"LP 25000 以上: {len(master_rows)}件")
    print("注意: 件数が少ないため、この結果は探索的な相関分析として解釈してください")

    print_section("LP要因分析（LP < 25000）", submaster_results)
    print_section("MR要因分析（LP >= 25000）", master_results)
    print_section("進捗スコア要因分析（LP/MR統合）", progress_results)

    write_results_csv(out_dir / "lp_factors.csv", submaster_results)
    write_results_csv(out_dir / "mr_factors.csv", master_results)
    write_results_csv(out_dir / "progress_factors.csv", progress_results)

    print(f"\n結果CSVを出力しました: {out_dir}")


if __name__ == "__main__":
    main()

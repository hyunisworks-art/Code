from __future__ import annotations

import argparse
import csv
from pathlib import Path

import playlog


DATA_START_ROW_INDEX = 4


def normalize_numeric_text(value: str) -> str:
    return value.strip().replace(",", "")


def choose_first_non_empty(row: dict[str, str], columns: list[str]) -> str:
    for column in columns:
        value = row.get(column, "")
        if value and value.strip():
            return value.strip()
    return ""


def detect_csv_encoding(path: Path) -> str:
    return playlog.detect_text_encoding(path)


def read_ranking_rows(path: Path) -> list[dict[str, str]]:
    encoding = detect_csv_encoding(path)
    with path.open("r", encoding=encoding, newline="") as file:
        return list(csv.DictReader(file))


def read_existing_keys(path: Path) -> set[tuple[str, str, str, str]]:
    encoding = detect_csv_encoding(path)
    keys: set[tuple[str, str, str, str]] = set()
    with path.open("r", encoding=encoding, newline="") as file:
        rows = list(csv.reader(file))

    for row in rows[DATA_START_ROW_INDEX:]:
        if len(row) < 6:
            continue
        date = row[1].strip()
        player = row[2].strip()
        lp = normalize_numeric_text(row[3])
        mr = normalize_numeric_text(row[5])
        if date and player:
            keys.add((date, player, lp, mr))
    return keys


def resolve_rank(lp_text: str, mr_text: str, lp_master: list[dict[str, object]], mr_master: list[dict[str, object]]) -> str:
    rank = playlog.resolve_rank(lp_text, lp_master)
    lp_value = playlog.parse_lp_value(lp_text)
    if lp_value is not None and lp_value >= 25000:
        rank = playlog.resolve_master_rank(lp_text, mr_text, mr_master)
    return rank


def build_playlog_row(no: int, date: str, player: str, lp: str, rank: str, mr: str) -> list[str]:
    row = [""] * 40
    row[0] = str(no)
    row[1] = date
    row[2] = player
    row[3] = lp
    row[4] = rank
    row[5] = mr
    return row


def main() -> None:
    parser = argparse.ArgumentParser(description="ランキングCSVを sf6-playlog-out.csv に追記する")
    parser.add_argument("--ranking-csv", default="ranking-output/master_p1-p3.csv", help="入力ランキングCSV")
    parser.add_argument("--output", default="sf6-playlog-out.csv", help="追記先 playlog CSV")
    parser.add_argument("--date", default="", help="データ取得日 (既定: 今日)")
    parser.add_argument("--dry-run", action="store_true", help="追記せず件数のみ表示")
    args = parser.parse_args()

    ranking_csv_path = Path(args.ranking_csv)
    output_path = Path(args.output)

    if not ranking_csv_path.exists():
        raise FileNotFoundError(f"ランキングCSVが見つかりません: {ranking_csv_path}")
    if not output_path.exists():
        raise FileNotFoundError(f"出力CSVが見つかりません: {output_path}")

    lp_master = playlog.load_lp_master(playlog.get_default_lp_master_path())
    mr_master = playlog.load_mr_master(playlog.get_default_mr_master_path())

    date = playlog.get_today_date() if not args.date.strip() else args.date.strip()
    ranking_rows = read_ranking_rows(ranking_csv_path)
    existing_keys = read_existing_keys(output_path)

    appended_count = 0
    skipped_count = 0
    next_no = playlog.get_next_no(output_path)

    player_columns = [
        "fighter_banner_info.personal_info.fighter_id",
        "fighter_banner_info.main_circle.leader.fighter_id",
    ]
    lp_columns = [
        "fighter_banner_info.favorite_character_league_info.league_point",
        "league_point",
    ]
    mr_columns = [
        "rating",
        "fighter_banner_info.favorite_character_league_info.master_rating",
    ]

    for ranking_row in ranking_rows:
        player = choose_first_non_empty(ranking_row, player_columns)
        lp_text = normalize_numeric_text(choose_first_non_empty(ranking_row, lp_columns))
        mr_text = normalize_numeric_text(choose_first_non_empty(ranking_row, mr_columns))

        if not player or not lp_text:
            skipped_count += 1
            continue

        if not mr_text:
            mr_text = "0"

        key = (date, player, lp_text, mr_text)
        if key in existing_keys:
            skipped_count += 1
            continue

        rank = resolve_rank(lp_text, mr_text, lp_master, mr_master)
        row = build_playlog_row(next_no, date, player, lp_text, rank, mr_text)

        if not args.dry_run:
            playlog.append_csv_row(output_path, row)

        existing_keys.add(key)
        next_no += 1
        appended_count += 1

    print(f"appended={appended_count}")
    print(f"skipped={skipped_count}")
    if args.dry_run:
        print("dry_run=true")


if __name__ == "__main__":
    main()
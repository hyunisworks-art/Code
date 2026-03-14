"""SF6 Buckler ランキング＋プロフィール実績を一括取得して sf6-playlog-out.csv に書き込む。

動作:
  1. ランキングページを取得してプレイヤー一覧 (short_id, LP, MR) を収集
  2. 各プレイヤーのプロフィールページから battle_stats を取得
  3. sf6-playlog-out.csv に書き込む
     - 実績列が空の既存行 → インプレース更新（取得のたびに即保存）
     - 未登録の行 → 末尾に追加（取得のたびに即保存）
     - 実績列が埋まっている行 → スキップ

使い方:
    python collect_playlog.py
    python collect_playlog.py --end-page 5
    python collect_playlog.py --end-page 3 --limit 10 --dry-run
    python collect_playlog.py --start-page 1 --end-page 466913 --page-step 46691 --dry-run
    python collect_playlog.py --start-page 1 --end-page 466913 --page-step 46691 --random-start-offset --dry-run
"""
from __future__ import annotations

import argparse
import csv
import io
import random
import time
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError

import playlog as _playlog
import scrape_rankings as _ranking


DEFAULT_OUTPUT       = "sf6-playlog-out.csv"
DEFAULT_RANKING_TYPE = "master"
DEFAULT_START_PAGE   = 1
DEFAULT_END_PAGE     = 3
DEFAULT_DELAY        = 1.5
MIN_PAGE_STEP        = 1000
MIN_ALLOWED_LP       = 9000
DATA_START_ROW_INDEX = 4   # 行 0〜3 はヘッダ

# ランキング flatten 後の列名
_COL_SHORT_ID = "fighter_banner_info.personal_info.short_id"
_COL_PLAYER   = "fighter_banner_info.personal_info.fighter_id"
_COL_LP       = "fighter_banner_info.favorite_character_league_info.league_point"
_COL_MR       = "fighter_banner_info.favorite_character_league_info.master_rating"

# ---------------------------------------------------------------------------
# battle_stats → playlog 列インデックス マッピング
# ---------------------------------------------------------------------------
# %値(0.0〜1.0) → "XX.XX%" 形式
_PCT_MAP: dict[int, str] = {
    6:  "gauge_rate_drive_guard",            # ドライブパリィ%
    7:  "gauge_rate_drive_impact",           # ドライブインパクト%
    8:  "gauge_rate_drive_arts",             # オーバードライブアーツ%
    9:  "gauge_rate_drive_rush_from_parry",  # パリィドライブラッシュ%
    10: "gauge_rate_drive_rush_from_cancel", # キャンセルドライブラッシュ%
    11: "gauge_rate_drive_reversal",         # ドライブリバーサル%
    12: "gauge_rate_drive_other",            # ダメージ/その他%
    24: "gauge_rate_sa_lv1",                 # SA Lv1%
    25: "gauge_rate_sa_lv2",                 # SA Lv2%
    26: "gauge_rate_sa_lv3",                 # SA Lv3%
    27: "gauge_rate_ca",                     # CA%
}

# 数値（平均回数・秒数・累計整数）→ そのまま文字列
_NUM_MAP: dict[int, str] = {
    13: "drive_reversal",
    14: "drive_parry",
    15: "throw_drive_parry",
    16: "received_throw_drive_parry",
    17: "just_parry",
    18: "drive_impact",
    19: "punish_counter",
    20: "drive_impact_to_drive_impact",
    21: "received_drive_impact",
    22: "received_punish_counter",
    23: "received_drive_impact_to_drive_impact",
    28: "stun",
    29: "received_stun",
    30: "throw_count",
    31: "received_throw_count",
    32: "throw_tech",
    33: "corner_time",
    34: "cornered_time",
    35: "rank_match_play_count",
    36: "casual_match_play_count",
    37: "custom_room_match_play_count",
    38: "battle_hub_match_play_count",
    39: "total_all_character_play_point",
}


# ---------------------------------------------------------------------------
# 書式変換
# ---------------------------------------------------------------------------

def _fmt_pct(val: Any) -> str:
    """API の 0.0〜1.0 値 → '9.01%' 形式。"""
    if val is None or val == "":
        return ""
    try:
        return f"{float(val) * 100:.2f}%"
    except (ValueError, TypeError):
        return str(val)


def _fmt_num(val: Any) -> str:
    return "" if val is None else str(val)


# ---------------------------------------------------------------------------
# CSV 入出力
# ---------------------------------------------------------------------------

def _read_all_rows(output_path: Path) -> tuple[list[list[str]], str]:
    """CSV 全行をリスト形式で返す。戻り値: (rows, encoding)"""
    encoding = _playlog.detect_text_encoding(output_path)
    raw, _ = _playlog.read_text(output_path)
    return list(csv.reader(io.StringIO(raw))), encoding


def _rewrite_csv(output_path: Path, rows: list[list[str]], encoding: str) -> None:
    """CSV 全行を上書き保存する。cp932 で表現できない文字があれば utf-8-sig に昇格。"""
    flat = "\r\n".join(",".join(r) for r in rows)
    if encoding == "cp932" and not _playlog.can_encode_text(flat, "cp932"):
        encoding = "utf-8-sig"
    with output_path.open("w", encoding=encoding, newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)


def _build_target_pages(
    start_page: int,
    end_page: int,
    page_step: int | None,
    random_start_offset: bool,
    random_seed: int | None,
) -> tuple[list[int], int]:
    """取得対象ページ一覧と、start_page からのオフセットを返す。"""
    if page_step is None:
        return list(range(start_page, end_page + 1)), 0

    if page_step < MIN_PAGE_STEP:
        raise ValueError(f"--page-step は {MIN_PAGE_STEP} 以上で指定してください")

    offset = 0
    if random_start_offset:
        max_offset = min(page_step - 1, end_page - start_page)
        if max_offset > 0:
            rng = random.Random(random_seed)
            offset = rng.randint(0, max_offset)

    first_page = start_page + offset
    return list(range(first_page, end_page + 1, page_step)), offset


def _has_missing_stats(row: list[str]) -> bool:
    """col 6〜39 に欠損がある行を欠損扱いにする。"""
    if len(row) < 40:
        return True
    return any(not row[idx].strip() for idx in range(6, 40))


def _parse_int_field(value: str) -> int | None:
    try:
        return int(value.strip().replace(",", ""))
    except (ValueError, AttributeError):
        return None


def _row_quality(row: list[str]) -> tuple[int, int, int]:
    """重複行でどちらを残すかの優先度（高いほど優先）。"""
    filled = sum(1 for idx in range(6, 40) if idx < len(row) and row[idx].strip())

    lp = _parse_int_field(row[3]) if len(row) > 3 else None

    mr = _parse_int_field(row[5]) if len(row) > 5 else None

    return (filled, lp if lp is not None else -1, mr if mr is not None else -1)


def _cleanup_output_csv(output_path: Path) -> tuple[int, int, int]:
    """同日重複・欠損・低LP行を除去し、必要ならCSVを書き直す。"""
    all_rows, encoding = _read_all_rows(output_path)
    header_rows = all_rows[:DATA_START_ROW_INDEX]
    data_rows = all_rows[DATA_START_ROW_INDEX:]

    cleaned_rows: list[list[str]] = []
    seen_same_day: dict[tuple[str, str], int] = {}
    removed_missing = 0
    removed_low_lp = 0
    removed_duplicate = 0

    for source_row in data_rows:
        row = list(source_row)
        while len(row) < 40:
            row.append("")

        if _has_missing_stats(row):
            removed_missing += 1
            continue

        lp = _parse_int_field(row[3]) if len(row) > 3 else None
        if lp is None or lp < MIN_ALLOWED_LP:
            removed_low_lp += 1
            continue

        date = row[1].strip() if len(row) > 1 else ""
        player = row[2].strip() if len(row) > 2 else ""

        if date and player:
            key = (date, player)
            prev_idx = seen_same_day.get(key)
            if prev_idx is not None:
                prev_row = cleaned_rows[prev_idx]
                if _row_quality(row) >= _row_quality(prev_row):
                    cleaned_rows[prev_idx] = row
                removed_duplicate += 1
                continue
            seen_same_day[key] = len(cleaned_rows)

        cleaned_rows.append(row)

    for idx, row in enumerate(cleaned_rows, start=1):
        row[0] = str(idx)

    changed = (
        (removed_missing > 0)
        or (removed_low_lp > 0)
        or (removed_duplicate > 0)
        or (len(cleaned_rows) != len(data_rows))
    )
    if changed:
        _rewrite_csv(output_path, header_rows + cleaned_rows, encoding)

    return removed_missing, removed_low_lp, removed_duplicate


# ---------------------------------------------------------------------------
# ランキング取得
# ---------------------------------------------------------------------------

def _fetch_ranking_entries(
    ranking_type: str,
    start_page: int,
    end_page: int,
    page_step: int | None,
    random_start_offset: bool,
    random_seed: int | None,
    locale: str,
    delay: float,
    timeout: int,
    cookie: str,
) -> list[dict[str, str]]:
    pages, offset = _build_target_pages(
        start_page=start_page,
        end_page=end_page,
        page_step=page_step,
        random_start_offset=random_start_offset,
        random_seed=random_seed,
    )
    first_url = _ranking.build_ranking_page_url(ranking_type, pages[0], locale)
    html_headers = _ranking.make_headers(cookie=cookie, referer=first_url)
    try:
        html = _ranking.fetch_text(first_url, html_headers, timeout)
    except HTTPError as exc:
        if exc.code == 403:
            raise PermissionError(
                "ランキングHTMLが 403 で拒否されました。Cookie を --cookie-file で渡してください。"
            ) from exc
        raise
    build_id = _ranking.get_build_id(html)

    entries: list[dict[str, str]] = []
    for idx, page in enumerate(pages):
        page_url = _ranking.build_ranking_page_url(ranking_type, page, locale)
        api_url  = _ranking.build_next_data_url(build_id, ranking_type, page, locale)
        headers  = _ranking.make_headers(cookie=cookie, referer=page_url)

        try:
            data = _ranking.fetch_json(api_url, headers, timeout)
        except HTTPError as exc:
            if exc.code == 403:
                raise PermissionError(
                    f"ランキングJSON p{page} が 403 で拒否されました。Cookie を確認してください。"
                ) from exc
            raise
        except URLError as exc:
            raise ConnectionError(f"ランキングJSON p{page} の取得に失敗しました: {exc}") from exc

        page_props    = data.get("pageProps", {})
        payload       = _ranking.get_ranking_payload(page_props, ranking_type)
        ranking_items = payload.get("ranking_fighter_list", [])

        for item in ranking_items:
            if not isinstance(item, dict):
                continue
            flat     = _ranking.flatten_item(item)
            short_id = flat.get(_COL_SHORT_ID, "").strip()
            player   = flat.get(_COL_PLAYER, "").strip()
            lp       = flat.get(_COL_LP, "").strip().replace(",", "")
            mr       = flat.get(_COL_MR, "").strip().replace(",", "")
            lp_value = _parse_int_field(lp)
            if short_id and player and lp_value is not None and lp_value >= MIN_ALLOWED_LP:
                entries.append({"short_id": short_id, "player": player, "lp": str(lp_value), "mr": mr or "0"})

        if idx < len(pages) - 1:
            time.sleep(delay)

    if page_step is None:
        print(f"  ランキング取得完了: {len(entries)} 件 ({start_page}〜{end_page}ページ)")
    else:
        rand_info = ""
        if random_start_offset:
            rand_info = f", random_offset={offset}, seed={random_seed}"
        print(
            f"  ランキング取得完了: {len(entries)} 件 "
            f"(start={start_page}, end={end_page}, step={page_step}, 取得ページ数={len(pages)}{rand_info})"
        )
    return entries


# ---------------------------------------------------------------------------
# プロフィール（battle_stats）取得
# ---------------------------------------------------------------------------

def _fetch_battle_stats(short_id: str, cookie: str, timeout: int) -> dict[str, Any]:
    url     = f"{_ranking.BASE_URL}/profile/{short_id}"
    headers = _ranking.make_headers(cookie=cookie, referer=url)
    try:
        html = _ranking.fetch_text(url, headers, timeout)
    except HTTPError as exc:
        if exc.code == 403:
            raise PermissionError(
                f"short_id={short_id}: プロフィールが 403 で拒否されました。Cookie を確認してください。"
            ) from exc
        raise
    next_data  = _ranking.extract_next_data(html)
    page_props = next_data.get("props", {}).get("pageProps", {})
    if page_props.get("common", {}).get("statusCode") == 403:
        raise PermissionError(f"short_id={short_id}: pageProps.statusCode=403")
    return page_props.get("play", {}).get("battle_stats", {})


# ---------------------------------------------------------------------------
# 40列 playlog 行を組み立てる
# ---------------------------------------------------------------------------

def _apply_stats_to_row(row: list[str], stats: dict[str, Any]) -> list[str]:
    """既存 row の col 6〜39 を battle_stats で上書きして返す（長さ 40 に揃える）。"""
    row = list(row)
    while len(row) < 40:
        row.append("")
    for col_idx, key in _PCT_MAP.items():
        row[col_idx] = _fmt_pct(stats.get(key, ""))
    for col_idx, key in _NUM_MAP.items():
        row[col_idx] = _fmt_num(stats.get(key, ""))
    return row


def _build_new_row(
    no: int,
    date: str,
    player: str,
    lp: str,
    mr: str,
    rank: str,
    stats: dict[str, Any],
) -> list[str]:
    row = [""] * 40
    row[0] = str(no)
    row[1] = date
    row[2] = player
    row[3] = lp
    row[4] = rank
    row[5] = mr
    return _apply_stats_to_row(row, stats)


def _resolve_rank(lp: str, mr: str, lp_master: list, mr_master: list) -> str:
    rank   = _playlog.resolve_rank(lp, lp_master)
    lp_val = _playlog.parse_lp_value(lp)
    if lp_val is not None and lp_val >= 25000:
        rank = _playlog.resolve_master_rank(lp, mr, mr_master)
    return rank


# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------

def collect(
    ranking_type: str,
    start_page: int,
    end_page: int,
    page_step: int | None,
    random_start_offset: bool,
    random_seed: int | None,
    locale: str,
    delay: float,
    timeout: int,
    cookie: str,
    output_path: Path,
    limit: int | None,
    dry_run: bool,
) -> None:
    lp_master = _playlog.load_lp_master(_playlog.get_default_lp_master_path())
    mr_master = _playlog.load_mr_master(_playlog.get_default_mr_master_path())
    date = _playlog.get_today_date()

    # ── 1. ランキング取得
    if page_step is None:
        print(f"[1/3] ランキング取得 ({ranking_type}, {start_page}〜{end_page}ページ)")
    else:
        rand_label = ""
        if random_start_offset:
            rand_label = f", random-start-offset=True, seed={random_seed}"
        print(
            f"[1/3] ランキング取得 ({ranking_type}, start={start_page}, end={end_page}, step={page_step}{rand_label})"
        )
    entries = _fetch_ranking_entries(
        ranking_type,
        start_page,
        end_page,
        page_step,
        random_start_offset,
        random_seed,
        locale,
        delay,
        timeout,
        cookie,
    )
    if limit is not None:
        entries = entries[:limit]
        print(f"  --limit {limit} 適用: {len(entries)} 件に絞り込み")

    # ── 2. 既存 CSV を読み込んで振り分け
    if not output_path.exists():
        raise FileNotFoundError(f"出力CSVが見つかりません: {output_path}")

    all_rows, encoding = _read_all_rows(output_path)
    header_rows = all_rows[:DATA_START_ROW_INDEX]
    data_rows   = all_rows[DATA_START_ROW_INDEX:]

    # (player, lp) → data_rows のインデックス（重複時は後ろを優先）
    player_lp_to_idx: dict[tuple[str, str], int] = {}
    for idx, row in enumerate(data_rows):
        if len(row) < 4:
            continue
        p  = row[2].strip()
        lp = row[3].strip().replace(",", "")
        if p and lp:
            player_lp_to_idx[(p, lp)] = idx

    to_update: list[tuple[int, dict[str, str]]] = []  # (data_rows_idx, entry)
    to_insert: list[dict[str, str]] = []
    skip_count = 0
    seen_keys: set[tuple[str, str]] = set()  # 同一エントリの重複処理を防ぐ

    for entry in entries:
        key = (entry["player"], entry["lp"])
        if key in seen_keys:
            continue
        seen_keys.add(key)

        if key in player_lp_to_idx:
            idx = player_lp_to_idx[key]
            row = data_rows[idx]
            has_stats = len(row) >= 40 and row[6].strip() != ""
            if has_stats:
                skip_count += 1
            else:
                to_update.append((idx, entry))
        else:
            to_insert.append(entry)

    print(
        f"[2/3] 振り分け: 更新={len(to_update)}件 / 新規追加={len(to_insert)}件 / スキップ={skip_count}件"
    )

    if dry_run:
        print("\n--dry-run 指定のため書き込みは行いません。")
        if to_update:
            print("  ▼ 更新予定（実績空欄の既存行）:")
            for idx, e in to_update:
                print(f"    No={data_rows[idx][0]}  {e['player']}  LP={e['lp']}")
        if to_insert:
            print("  ▼ 新規追加予定:")
            for e in to_insert:
                print(f"    {e['short_id']}  {e['player']}  LP={e['lp']}")
        return

    if not to_update and not to_insert:
        removed_missing, removed_low_lp, removed_duplicate = _cleanup_output_csv(output_path)
        print("処理対象がないため終了します。")
        print(
            f"整形: 同日重複削除={removed_duplicate}  欠損行削除={removed_missing}  LP{MIN_ALLOWED_LP}未満削除={removed_low_lp}"
        )
        print(f"出力: {output_path}")
        return

    total = len(to_update) + len(to_insert)
    updated_count  = 0
    inserted_count = 0
    error_count    = 0
    counter        = 0

    print(f"[3/3] プロフィール取得・書き込み (更新{len(to_update)}件 + 新規{len(to_insert)}件)")

    # ── 3a. 更新（実績が空の既存行）→ 1件ごとに CSV を即保存
    for idx, entry in to_update:
        counter += 1
        sid    = entry["short_id"]
        player = entry["player"]
        lp     = entry["lp"]
        mr     = entry["mr"]
        try:
            stats    = _fetch_battle_stats(sid, cookie, timeout)
            rank_cnt = stats.get("rank_match_play_count", "?")
            data_rows[idx] = _apply_stats_to_row(data_rows[idx], stats)
            _rewrite_csv(output_path, header_rows + data_rows, encoding)
            updated_count += 1
            print(f"  [{counter}/{total}] 更新  {player}  LP={lp}  MR={mr}  ランクマッチ={rank_cnt}")
        except (PermissionError, ValueError, HTTPError, URLError, RuntimeError) as exc:
            print(f"  [{counter}/{total}] ERR   {player} ({sid}): {exc}")
            error_count += 1
        if counter < total:
            time.sleep(delay)

    # ── 3b. 新規追加 → append_csv_row で1件ごとに即保存
    next_no = _playlog.get_next_no(output_path)
    for entry in to_insert:
        counter += 1
        sid    = entry["short_id"]
        player = entry["player"]
        lp     = entry["lp"]
        mr     = entry["mr"]
        rank   = _resolve_rank(lp, mr, lp_master, mr_master)
        try:
            stats    = _fetch_battle_stats(sid, cookie, timeout)
            rank_cnt = stats.get("rank_match_play_count", "?")
            new_row  = _build_new_row(next_no, date, player, lp, mr, rank, stats)
            _playlog.append_csv_row(output_path, new_row)
            next_no += 1
            inserted_count += 1
            print(f"  [{counter}/{total}] 追加  {player}  LP={lp}  MR={mr}  ランクマッチ={rank_cnt}")
        except (PermissionError, ValueError, HTTPError, URLError, RuntimeError) as exc:
            print(f"  [{counter}/{total}] ERR   {player} ({sid}): {exc}")
            error_count += 1
        if counter < total:
            time.sleep(delay)

    removed_missing, removed_low_lp, removed_duplicate = _cleanup_output_csv(output_path)

    print()
    print(f"完了: 更新={updated_count}  新規追加={inserted_count}  エラー={error_count}  スキップ={skip_count}")
    print(
        f"整形: 同日重複削除={removed_duplicate}  欠損行削除={removed_missing}  LP{MIN_ALLOWED_LP}未満削除={removed_low_lp}"
    )
    print(f"出力: {output_path}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="ランキング取得＋プロフィール実績を一括取得して sf6-playlog-out.csv に書き込む"
    )
    parser.add_argument(
        "--ranking-type", choices=("master", "league"), default=DEFAULT_RANKING_TYPE,
        help=f"ランキング種別（既定: {DEFAULT_RANKING_TYPE}）",
    )
    parser.add_argument(
        "--start-page", type=int, default=DEFAULT_START_PAGE,
        help=f"取得開始ページ（既定: {DEFAULT_START_PAGE}）",
    )
    parser.add_argument(
        "--end-page", type=int, default=DEFAULT_END_PAGE,
        help=f"取得終了ページ（既定: {DEFAULT_END_PAGE}）",
    )
    parser.add_argument(
        "--page-step", type=int, default=None, metavar="N",
        help=(
            "ページを N 件ずつ飛ばして取得する（例: 1, 1+N, 1+2N ...）。"
            f"安全のため N は {MIN_PAGE_STEP} 以上のみ許可"
        ),
    )
    parser.add_argument(
        "--random-start-offset", action="store_true",
        help="--page-step 指定時に、開始ページを [start, start+N-1] の範囲でランダム化する",
    )
    parser.add_argument(
        "--random-seed", type=int, default=None,
        help="--random-start-offset の乱数シード（同じ値で同じ開始オフセットを再現）",
    )
    parser.add_argument("--locale", default=_ranking.DEFAULT_LOCALE, help="ロケール（既定: en）")
    parser.add_argument(
        "--delay", type=float, default=DEFAULT_DELAY,
        help=f"リクエスト間の待機秒数（既定: {DEFAULT_DELAY}）",
    )
    parser.add_argument("--timeout", type=int, default=_ranking.DEFAULT_TIMEOUT)
    parser.add_argument(
        "--output", default=DEFAULT_OUTPUT,
        help=f"書き込み先 playlog CSV（既定: {DEFAULT_OUTPUT}）",
    )
    parser.add_argument("--cookie", default="", help="Cookie文字列（任意）")
    parser.add_argument(
        "--cookie-file", default=_ranking.DEFAULT_COOKIE_FILE,
        help="Cookie文字列を保存したファイル（既定: .buckler_cookie.txt）",
    )
    parser.add_argument(
        "--limit", type=int, default=None, metavar="N",
        help="処理件数の上限（省略時: 全件）",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="取得・書き込みを行わず対象件数のみ表示",
    )
    args = parser.parse_args()

    if args.start_page < 1 or args.end_page < args.start_page:
        raise ValueError("ページ範囲が不正です (start-page <= end-page かつ 1 以上)")
    if args.page_step is not None and args.page_step < MIN_PAGE_STEP:
        raise ValueError(f"--page-step は {MIN_PAGE_STEP} 以上で指定してください")
    if args.random_start_offset and args.page_step is None:
        raise ValueError("--random-start-offset を使う場合は --page-step も指定してください")

    cookie = _ranking.load_cookie_text(args.cookie, args.cookie_file)
    _ranking.validate_cookie_text(cookie)

    collect(
        ranking_type=args.ranking_type,
        start_page=args.start_page,
        end_page=args.end_page,
        page_step=args.page_step,
        random_start_offset=args.random_start_offset,
        random_seed=args.random_seed,
        locale=(args.locale.strip() or _ranking.DEFAULT_LOCALE),
        delay=args.delay,
        timeout=args.timeout,
        cookie=cookie,
        output_path=Path(args.output),
        limit=args.limit,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()

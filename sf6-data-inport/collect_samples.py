"""SF6 Buckler ランク帯別サンプル収集スクリプト (Phase 0 タスク2)

各ランク帯（Bronze/Silver/Gold/Platinum/Diamond/Master）から
プレイヤーをサンプリングし、battle_stats を CSV に保存する。

使い方:
    python collect_samples.py
    python collect_samples.py --rank platinum
    python collect_samples.py --rank platinum --count 30
    python collect_samples.py --rank gold --dry-run
    python collect_samples.py --count 20  # 全ランク帯を各20件収集

動作確認方法:
    1. Cookie ファイルを用意する（.buckler_cookie.txt にブラウザのCookieを貼り付ける）
    2. dry-run で対象件数を確認する:
       python collect_samples.py --rank platinum --count 5 --dry-run
    3. 実際に収集する:
       python collect_samples.py --rank platinum --count 5
    4. 出力先を確認する: data/samples/YYYY-MM-DD_rank_platinum.csv

注意:
    - 既存の scrape_rankings.py・analyze_step1.py は変更しない
    - プレイヤーIDはコード内にハードコードしない
    - Bucklerへの負荷軽減のため: リクエスト間隔3秒以上・50リクエスト上限/セッション
"""
from __future__ import annotations

import argparse
import csv
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError

# Windows環境でUTF-8出力を強制
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

import scrape_rankings as _ranking

# ---------------------------------------------------------------------------
# 定数
# ---------------------------------------------------------------------------

# ランキング種別とランク帯のマッピング
# league: Bronze / Silver / Gold / Platinum / Diamond
# master: Master
RANK_CONFIG: dict[str, dict] = {
    "bronze":   {"ranking_type": "league", "lp_min": 0,     "lp_max": 4999,  "label": "bronze"},
    "silver":   {"ranking_type": "league", "lp_min": 5000,  "lp_max": 9999,  "label": "silver"},
    "gold":     {"ranking_type": "league", "lp_min": 10000, "lp_max": 14999, "label": "gold"},
    "platinum": {"ranking_type": "league", "lp_min": 15000, "lp_max": 19999, "label": "platinum"},
    "diamond":  {"ranking_type": "league", "lp_min": 20000, "lp_max": 24999, "label": "diamond"},
    "master":   {"ranking_type": "master", "lp_min": 25000, "lp_max": None,  "label": "master"},
}

ALL_RANKS = ["bronze", "silver", "gold", "platinum", "diamond", "master"]

# レート制限
DEFAULT_DELAY = 3.0          # リクエスト間隔（秒）
MAX_REQUESTS_PER_SESSION = 50  # 1セッションあたり最大リクエスト数
MAX_RETRIES = 2              # エラー時の最大リトライ回数

DEFAULT_COUNT = 50           # 各ランク帯の収集上限（デフォルト）
SAMPLE_EXPIRE_DAYS = 30      # サンプルファイルの有効期間（日）

DEFAULT_COOKIE_FILE = ".buckler_cookie.txt"
SAMPLES_DIR = Path("data") / "samples"

# battle_stats から保存するフィールド（dump_battle_stats.py の MAPPED_FIELDS に相当）
# パーセント表示フィールド（APIの値は0.0〜1.0）
PCT_FIELDS: list[str] = [
    "gauge_rate_drive_guard",            # ドライブパリィ%
    "gauge_rate_drive_impact",           # ドライブインパクト%
    "gauge_rate_drive_arts",             # オーバードライブアーツ%
    "gauge_rate_drive_rush_from_parry",  # パリィドライブラッシュ%
    "gauge_rate_drive_rush_from_cancel", # キャンセルドライブラッシュ%
    "gauge_rate_drive_reversal",         # ドライブリバーサル%
    "gauge_rate_drive_other",            # ダメージ/その他%
    "gauge_rate_sa_lv1",                 # SA Lv1%
    "gauge_rate_sa_lv2",                 # SA Lv2%
    "gauge_rate_sa_lv3",                 # SA Lv3%
    "gauge_rate_ca",                     # CA%
]

# 数値フィールド（平均回数・秒数など）
NUM_FIELDS: list[str] = [
    "drive_reversal",
    "drive_parry",
    "throw_drive_parry",
    "received_throw_drive_parry",
    "just_parry",
    "drive_impact",
    "punish_counter",
    "drive_impact_to_drive_impact",
    "received_drive_impact",
    "received_punish_counter",
    "received_drive_impact_to_drive_impact",
    "stun",
    "received_stun",
    "throw_count",
    "received_throw_count",
    "throw_tech",
    "corner_time",
    "cornered_time",
    "rank_match_play_count",
    "casual_match_play_count",
    "custom_room_match_play_count",
    "battle_hub_match_play_count",
    "total_all_character_play_point",
]

# CSVの全列（player_id と rank を先頭に追加）
CSV_COLUMNS: list[str] = (
    ["player_id", "rank"]
    + [f"{f}_pct" for f in PCT_FIELDS]
    + NUM_FIELDS
)

# ランキングのflatten後の列名（collect_playlog.py と同じ定義）
_COL_SHORT_ID = "fighter_banner_info.personal_info.short_id"
_COL_LP       = "fighter_banner_info.favorite_character_league_info.league_point"


# ---------------------------------------------------------------------------
# フォーマット変換
# ---------------------------------------------------------------------------

def _fmt_pct(val: Any) -> str:
    """APIの0.0〜1.0値を '9.01%' 形式に変換する。"""
    if val is None or val == "":
        return ""
    try:
        return f"{float(val) * 100:.2f}%"
    except (ValueError, TypeError):
        return str(val)


def _fmt_num(val: Any) -> str:
    return "" if val is None else str(val)


# ---------------------------------------------------------------------------
# プレイヤーID収集
# ---------------------------------------------------------------------------

def _fetch_short_ids_for_rank(
    rank_key: str,
    count: int,
    cookie: str,
    timeout: int,
    delay: float,
    request_counter: list[int],
) -> list[str]:
    """指定ランク帯のランキングからshort_idを最大count件収集する。"""
    config = RANK_CONFIG[rank_key]
    ranking_type = config["ranking_type"]
    lp_min = config["lp_min"]
    lp_max = config["lp_max"]

    short_ids: list[str] = []

    # ランキング1ページあたり大体20件。必要ページ数を概算
    pages_needed = max(1, (count // 20) + 2)

    # 最初のページでbuild_idを取得
    first_url = _ranking.build_ranking_page_url(ranking_type, 1, _ranking.DEFAULT_LOCALE)
    headers = _ranking.make_headers(cookie=cookie, referer=first_url)

    request_counter[0] += 1
    if request_counter[0] > MAX_REQUESTS_PER_SESSION:
        print(f"  ※ セッション上限({MAX_REQUESTS_PER_SESSION}リクエスト)に達したため停止します")
        return short_ids

    try:
        html = _ranking.fetch_text(first_url, headers, timeout)
    except HTTPError as exc:
        raise PermissionError(
            f"ランキングHTML取得失敗 (HTTP {exc.code}): Cookieを確認してください"
        ) from exc

    build_id = _ranking.get_build_id(html)

    for page in range(1, pages_needed + 1):
        if len(short_ids) >= count:
            break

        if request_counter[0] >= MAX_REQUESTS_PER_SESSION:
            print(f"  ※ セッション上限({MAX_REQUESTS_PER_SESSION}リクエスト)に達したため停止します")
            break

        page_url = _ranking.build_ranking_page_url(ranking_type, page, _ranking.DEFAULT_LOCALE)
        api_url  = _ranking.build_next_data_url(build_id, ranking_type, page, _ranking.DEFAULT_LOCALE)
        api_headers = _ranking.make_headers(cookie=cookie, referer=page_url)

        request_counter[0] += 1
        try:
            data = _ranking.fetch_json(api_url, api_headers, timeout)
        except (HTTPError, URLError) as exc:
            print(f"  ランキングページ{page}取得失敗: {exc}")
            break

        page_props = data.get("pageProps", {})
        try:
            payload = _ranking.get_ranking_payload(page_props, ranking_type)
        except (ValueError, PermissionError) as exc:
            print(f"  ランキングデータ取得失敗: {exc}")
            break

        ranking_items = payload.get("ranking_fighter_list", [])
        if not isinstance(ranking_items, list):
            break

        for item in ranking_items:
            if len(short_ids) >= count:
                break
            if not isinstance(item, dict):
                continue

            flat = _ranking.flatten_item(item)
            short_id = flat.get(_COL_SHORT_ID, "").strip()
            lp_str   = flat.get(_COL_LP, "").strip().replace(",", "")

            if not short_id:
                continue

            # LP範囲フィルタリング（league ランクの場合は LP でランク帯を絞る）
            if ranking_type == "league" and lp_str.isdigit():
                lp_val = int(lp_str)
                if lp_val < lp_min:
                    continue
                if lp_max is not None and lp_val > lp_max:
                    continue

            short_ids.append(short_id)

        if page < pages_needed:
            time.sleep(delay)

    return short_ids[:count]


# ---------------------------------------------------------------------------
# battle_stats 取得
# ---------------------------------------------------------------------------

def _fetch_battle_stats_with_retry(
    short_id: str,
    cookie: str,
    timeout: int,
    max_retries: int,
    request_counter: list[int],
    delay: float,
) -> dict[str, Any] | None:
    """battle_statsを取得する。エラー時は最大max_retries回リトライ。"""
    for attempt in range(max_retries + 1):
        if request_counter[0] >= MAX_REQUESTS_PER_SESSION:
            return None

        request_counter[0] += 1
        try:
            url     = f"{_ranking.BASE_URL}/profile/{short_id}"
            headers = _ranking.make_headers(cookie=cookie, referer=url)
            html    = _ranking.fetch_text(url, headers, timeout)
            next_data  = _ranking.extract_next_data(html)
            page_props = next_data.get("props", {}).get("pageProps", {})
            if page_props.get("common", {}).get("statusCode") == 403:
                raise PermissionError(f"pageProps.statusCode=403")
            stats = page_props.get("play", {}).get("battle_stats", {})
            return stats

        except (PermissionError, ValueError, HTTPError, URLError, RuntimeError) as exc:
            if attempt < max_retries:
                print(f"    リトライ {attempt + 1}/{max_retries}: {exc}")
                time.sleep(delay)
            else:
                print(f"    スキップ (エラー): {exc}")
                return None

    return None


# ---------------------------------------------------------------------------
# CSV 操作
# ---------------------------------------------------------------------------

def _load_existing_ids(csv_path: Path) -> set[str]:
    """既存CSVからplayer_idのセットを読み込む。"""
    if not csv_path.exists():
        return set()

    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            return {row["player_id"] for row in reader if row.get("player_id")}
    except Exception:
        return set()


def _build_sample_row(short_id: str, rank_key: str, stats: dict[str, Any]) -> dict[str, str]:
    """battle_stats から CSV の1行分の辞書を作る。"""
    row: dict[str, str] = {
        "player_id": short_id,
        "rank": rank_key,
    }
    for field in PCT_FIELDS:
        row[f"{field}_pct"] = _fmt_pct(stats.get(field, ""))
    for field in NUM_FIELDS:
        row[field] = _fmt_num(stats.get(field, ""))
    return row


def _append_rows_to_csv(csv_path: Path, rows: list[dict[str, str]]) -> None:
    """CSVにデータ行を追記する（ファイルがなければ新規作成）。"""
    file_exists = csv_path.exists()
    with csv_path.open("a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)


# ---------------------------------------------------------------------------
# 古いサンプルファイルの削除
# ---------------------------------------------------------------------------

def _cleanup_old_samples(samples_dir: Path, expire_days: int) -> int:
    """expire_days日以上古いサンプルCSVを削除する。削除件数を返す。"""
    if not samples_dir.exists():
        return 0

    cutoff = datetime.now() - timedelta(days=expire_days)
    deleted = 0

    for csv_file in samples_dir.glob("*_rank_*.csv"):
        # ファイル名の先頭の日付でチェック（YYYY-MM-DD_rank_XXX.csv）
        try:
            date_str = csv_file.stem.split("_rank_")[0]
            file_date = datetime.strptime(date_str, "%Y-%m-%d")
            if file_date < cutoff:
                csv_file.unlink()
                print(f"  削除（期限切れ）: {csv_file.name}")
                deleted += 1
        except (ValueError, IndexError):
            # 日付の解析に失敗したファイルはスキップ
            continue

    return deleted


# ---------------------------------------------------------------------------
# サンプル収集メイン処理
# ---------------------------------------------------------------------------

def collect_samples_for_rank(
    rank_key: str,
    count: int,
    cookie: str,
    timeout: int,
    delay: float,
    dry_run: bool,
    request_counter: list[int],
) -> tuple[int, int]:
    """1ランク帯のサンプルを収集する。戻り値: (保存件数, スキップ件数)"""
    today = date.today().strftime("%Y-%m-%d")
    csv_path = SAMPLES_DIR / f"{today}_rank_{rank_key}.csv"

    print(f"\n収集開始: {rank_key}帯 上限{count}件")

    # 既存データのplayer_idを読み込む（重複チェック用）
    existing_ids = _load_existing_ids(csv_path)
    if existing_ids:
        print(f"  既存データ: {len(existing_ids)}件（重複はスキップ）")

    # プレイヤーID収集
    print("  ランキングページ取得中...")
    short_ids = _fetch_short_ids_for_rank(
        rank_key=rank_key,
        count=count,
        cookie=cookie,
        timeout=timeout,
        delay=delay,
        request_counter=request_counter,
    )
    print(f"  プレイヤーID収集: {len(short_ids)}件")

    if not short_ids:
        print("  収集できるIDがありませんでした")
        return 0, 0

    # battle_stats 取得とCSV保存
    rows_to_save: list[dict[str, str]] = []
    saved_count  = 0
    skipped_count = 0

    for idx, short_id in enumerate(short_ids, start=1):
        # セッション上限チェック
        if request_counter[0] >= MAX_REQUESTS_PER_SESSION:
            print(f"  ※ セッション上限({MAX_REQUESTS_PER_SESSION}リクエスト)に達したため停止します")
            break

        # 重複チェック
        if short_id in existing_ids:
            print(f"  [{idx}/{len(short_ids)}] player_id={short_id} スキップ（既存）")
            skipped_count += 1
            continue

        # データ取得
        print(f"  [{idx}/{len(short_ids)}] player_id={short_id} データ取得...", end=" ", flush=True)

        if dry_run:
            print("(dry-run: スキップ)")
            saved_count += 1  # dry-run では保存予定件数としてカウント
            continue

        stats = _fetch_battle_stats_with_retry(
            short_id=short_id,
            cookie=cookie,
            timeout=timeout,
            max_retries=MAX_RETRIES,
            request_counter=request_counter,
            delay=delay,
        )

        if stats is None:
            # _fetch_battle_stats_with_retry の中でエラーログ出力済み
            continue

        row = _build_sample_row(short_id, rank_key, stats)
        rows_to_save.append(row)
        existing_ids.add(short_id)  # 同セッション内の重複防止
        saved_count += 1
        print("OK")

        # リクエスト間隔
        if idx < len(short_ids):
            time.sleep(delay)

    # CSVに保存（dry-run 以外）
    if not dry_run and rows_to_save:
        SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
        _append_rows_to_csv(csv_path, rows_to_save)
        print(f"\n完了: {saved_count}件保存 / {skipped_count}件スキップ")
        print(f"保存先: {csv_path}")
    elif dry_run:
        print(f"\n[dry-run] 完了: {saved_count}件保存予定 / {skipped_count}件スキップ予定")
    else:
        print(f"\n完了: 0件保存 / {skipped_count}件スキップ")

    return saved_count, skipped_count


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="SF6 Bucklerからランク帯別のbattle_statsサンプルを収集してCSVに保存する"
    )
    parser.add_argument(
        "--rank",
        choices=ALL_RANKS,
        default=None,
        help=(
            "対象ランク帯（bronze/silver/gold/platinum/diamond/master）。"
            "省略時は全ランク帯を順番に収集"
        ),
    )
    parser.add_argument(
        "--count",
        type=int,
        default=DEFAULT_COUNT,
        metavar="N",
        help=f"各ランク帯の収集件数上限（デフォルト: {DEFAULT_COUNT}）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="実際には保存せず動作確認だけ行う",
    )
    parser.add_argument(
        "--cookie",
        default="",
        help="Cookie文字列（任意。省略時は .buckler_cookie.txt を使用）",
    )
    parser.add_argument(
        "--cookie-file",
        default=DEFAULT_COOKIE_FILE,
        help=f"Cookie文字列を保存したファイル（デフォルト: {DEFAULT_COOKIE_FILE}）",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY,
        help=f"リクエスト間の待機秒数（デフォルト: {DEFAULT_DELAY}秒）",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=_ranking.DEFAULT_TIMEOUT,
        help="HTTPタイムアウト秒数",
    )
    args = parser.parse_args()

    if args.count < 1:
        print("エラー: --count は1以上を指定してください")
        sys.exit(1)
    if args.delay < 1.0:
        print("警告: --delay が1秒未満です。Bucklerへの負荷に注意してください")

    # Cookie 読み込み
    cookie = _ranking.load_cookie_text(args.cookie, args.cookie_file)
    try:
        _ranking.validate_cookie_text(cookie)
    except ValueError as exc:
        print(f"エラー: {exc}")
        sys.exit(1)

    if dry_run_label := ("（dry-run モード）" if args.dry_run else ""):
        print(f"=== SF6 サンプル収集スクリプト {dry_run_label}===")
    else:
        print("=== SF6 サンプル収集スクリプト ===")

    # 起動時: 期限切れサンプルファイルを削除
    deleted = _cleanup_old_samples(SAMPLES_DIR, SAMPLE_EXPIRE_DAYS)
    if deleted > 0:
        print(f"期限切れファイルを{deleted}件削除しました（{SAMPLE_EXPIRE_DAYS}日以上経過）")

    # 収集対象ランク帯
    target_ranks = [args.rank] if args.rank else ALL_RANKS

    # セッション全体のリクエストカウンター（共有参照）
    request_counter = [0]

    total_saved   = 0
    total_skipped = 0

    for rank_key in target_ranks:
        if request_counter[0] >= MAX_REQUESTS_PER_SESSION:
            print(f"\n※ セッション上限({MAX_REQUESTS_PER_SESSION}リクエスト)に達したため、残りのランク帯はスキップします")
            break

        try:
            saved, skipped = collect_samples_for_rank(
                rank_key=rank_key,
                count=args.count,
                cookie=cookie,
                timeout=args.timeout,
                delay=args.delay,
                dry_run=args.dry_run,
                request_counter=request_counter,
            )
            total_saved   += saved
            total_skipped += skipped

        except PermissionError as exc:
            print(f"\n権限エラー: {exc}")
            print("Cookie が必要な場合は .buckler_cookie.txt を確認してください")
            break
        except ConnectionError as exc:
            print(f"\n接続エラー: {exc}")
            break
        except KeyboardInterrupt:
            print("\n\n中断されました（Ctrl+C）")
            break

    print(f"\n=== 全体集計 ===")
    print(f"合計保存: {total_saved}件 / 合計スキップ: {total_skipped}件")
    print(f"リクエスト総数: {request_counter[0]}件")


if __name__ == "__main__":
    main()

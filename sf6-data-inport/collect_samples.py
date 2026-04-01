"""SF6 Buckler ランク帯別サンプル収集スクリプト (Phase 0 タスク2)

各ランク帯（Bronze/Silver/Gold/Platinum/Diamond/Master）から
プレイヤーをサンプリングし、play 配下の全データを JSON で保存する。

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
    4. 出力先を確認する: data/samples/YYYY-MM-DD_<short_id>.json

注意:
    - 既存の scrape_rankings.py・analyze_step1.py は変更しない
    - プレイヤーIDはコード内にハードコードしない
    - Bucklerへの負荷軽減のため: リクエスト間隔3秒以上・50リクエスト上限/セッション
"""
from __future__ import annotations

import argparse
import json
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
# league_rank の値は Buckler の pageProps.league_rank リストに基づく（2025年現在）
#   0=All, 1-5=Rookie, 6-10=Iron, 11-15=Bronze, 16-20=Silver,
#   21-25=Gold, 26-30=Platinum, 31-35=Diamond, 36=Master
# league: Bronze / Silver / Gold / Platinum / Diamond
# master: Master（league_rank=36, ranking_type="master" を使用）
RANK_CONFIG: dict[str, dict] = {
    "bronze":   {"ranking_type": "league", "league_rank_min": 11, "league_rank_max": 15, "label": "bronze"},
    "silver":   {"ranking_type": "league", "league_rank_min": 16, "league_rank_max": 20, "label": "silver"},
    "gold":     {"ranking_type": "league", "league_rank_min": 21, "league_rank_max": 25, "label": "gold"},
    "platinum": {"ranking_type": "league", "league_rank_min": 26, "league_rank_max": 30, "label": "platinum"},
    "diamond":  {"ranking_type": "league", "league_rank_min": 31, "league_rank_max": 35, "label": "diamond"},
    "master":   {"ranking_type": "master", "league_rank_min": 36, "league_rank_max": 36, "label": "master"},
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

# ランキングのflatten後の列名（collect_playlog.py と同じ定義）
_COL_SHORT_ID = "fighter_banner_info.personal_info.short_id"
_COL_LP       = "fighter_banner_info.favorite_character_league_info.league_point"


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
    exclude_ids: set[str] | None = None,
) -> list[str]:
    """指定ランク帯のランキングからshort_idを最大count件収集する。
    exclude_ids に含まれるIDはスキップし、count件の新規IDを集める。"""
    config = RANK_CONFIG[rank_key]
    ranking_type = config["ranking_type"]
    # league_rank の範囲（各ランク帯のサブランク最小〜最大）
    league_rank_min: int = config["league_rank_min"]
    league_rank_max: int = config["league_rank_max"]

    _exclude = exclude_ids or set()
    short_ids: list[str] = []
    seen: set[str] = set()  # セッション内重複防止

    # 既存IDを除いてcount件集めるため、ページ数を多めに確保
    pages_per_sub_rank = max(3, (count // 10) + 2)

    # 最初のページでbuild_idを取得（league_rank パラメータなしで取得）
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

    # league_rank_min〜league_rank_max のサブランクをループして収集
    for sub_rank in range(league_rank_min, league_rank_max + 1):
        if len(short_ids) >= count:
            break

        for page in range(1, pages_per_sub_rank + 1):
            if len(short_ids) >= count:
                break

            if request_counter[0] >= MAX_REQUESTS_PER_SESSION:
                print(f"  ※ セッション上限({MAX_REQUESTS_PER_SESSION}リクエスト)に達したため停止します")
                return short_ids[:count]

            page_url = _ranking.build_ranking_page_url(
                ranking_type, page, _ranking.DEFAULT_LOCALE, league_rank=sub_rank
            )
            api_url = _ranking.build_next_data_url(
                build_id, ranking_type, page, _ranking.DEFAULT_LOCALE, league_rank=sub_rank
            )
            api_headers = _ranking.make_headers(cookie=cookie, referer=page_url)

            request_counter[0] += 1
            try:
                data = _ranking.fetch_json(api_url, api_headers, timeout)
            except (HTTPError, URLError) as exc:
                print(f"  ランキングページ取得失敗 (league_rank={sub_rank}, page={page}): {exc}")
                break

            page_props = data.get("pageProps", {})
            try:
                payload = _ranking.get_ranking_payload(page_props, ranking_type)
            except PermissionError:
                raise
            except ValueError as exc:
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

                if not short_id:
                    continue
                if short_id in _exclude or short_id in seen:
                    continue

                seen.add(short_id)
                short_ids.append(short_id)

            if not ranking_items:
                break  # ページが空なら次のサブランクへ

            if page < pages_per_sub_rank:
                time.sleep(delay)

    return short_ids[:count]


# ---------------------------------------------------------------------------
# play データ取得
# ---------------------------------------------------------------------------

def _fetch_play_data_with_retry(
    short_id: str,
    cookie: str,
    timeout: int,
    max_retries: int,
    request_counter: list[int],
    delay: float,
) -> dict[str, Any] | None:
    """play 配下の全データを取得する。エラー時は最大max_retries回リトライ。"""
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
            play_data = page_props.get("play", {})
            return play_data

        except (PermissionError, ValueError, HTTPError, URLError, RuntimeError) as exc:
            if attempt < max_retries:
                print(f"    リトライ {attempt + 1}/{max_retries}: {exc}")
                time.sleep(delay)
            else:
                print(f"    スキップ (エラー): {exc}")
                return None

    return None


# ---------------------------------------------------------------------------
# JSON 操作
# ---------------------------------------------------------------------------

def _load_existing_ids(samples_dir: Path, today: str = "") -> set[str]:
    """保存済みのplayer_idセットを返す（全期間対象）。
    today引数は後方互換のために残すが使用しない。"""
    if not samples_dir.exists():
        return set()
    ids: set[str] = set()
    for p in samples_dir.glob("????-??-??_*.json"):
        parts = p.stem.split("_", 1)
        if len(parts) == 2:
            ids.add(parts[1])
    return ids


def _save_sample_json(json_path: Path, payload: dict[str, Any]) -> None:
    """JSONファイルとして保存する（上書き）。"""
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 古いサンプルファイルの削除
# ---------------------------------------------------------------------------

def _cleanup_old_samples(samples_dir: Path, expire_days: int) -> int:
    """expire_days日以上古いサンプルJSONを削除する。削除件数を返す。"""
    if not samples_dir.exists():
        return 0

    cutoff = datetime.now() - timedelta(days=expire_days)
    deleted = 0

    for json_file in samples_dir.glob("????-??-??_*.json"):
        # ファイル名の先頭の日付でチェック（YYYY-MM-DD_<short_id>.json）
        try:
            date_str = json_file.stem[:10]
            file_date = datetime.strptime(date_str, "%Y-%m-%d")
            if file_date < cutoff:
                json_file.unlink()
                print(f"  削除（期限切れ）: {json_file.name}")
                deleted += 1
        except (ValueError, IndexError):
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

    print(f"\n収集開始: {rank_key}帯 上限{count}件")

    # 既存データのplayer_idを読み込む（重複チェック用）
    existing_ids = _load_existing_ids(SAMPLES_DIR, today)
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
        exclude_ids=existing_ids,
    )
    print(f"  プレイヤーID収集: {len(short_ids)}件")

    if not short_ids:
        print("  収集できるIDがありませんでした")
        return 0, 0

    # play データ取得と JSON 保存
    saved_count   = 0
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

        play_data = _fetch_play_data_with_retry(
            short_id=short_id,
            cookie=cookie,
            timeout=timeout,
            max_retries=MAX_RETRIES,
            request_counter=request_counter,
            delay=delay,
        )

        if play_data is None:
            continue

        # JSON保存: fetch_my_data.py と同一構造（rank フィールドはサンプル側追加）
        payload: dict[str, Any] = {
            "fetch_date": today,
            "player_id": short_id,
            "rank": rank_key,
            "league_info": {},
            "play": play_data,
        }

        SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
        json_path = SAMPLES_DIR / f"{today}_{short_id}.json"
        _save_sample_json(json_path, payload)
        existing_ids.add(short_id)  # 同セッション内の重複防止
        saved_count += 1
        print("OK")

        # リクエスト間隔
        if idx < len(short_ids):
            time.sleep(delay)

    if dry_run:
        print(f"\n[dry-run] 完了: {saved_count}件保存予定 / {skipped_count}件スキップ予定")
    else:
        print(f"\n完了: {saved_count}件保存 / {skipped_count}件スキップ")
        print(f"保存先: {SAMPLES_DIR}/{today}_<short_id>.json")

    return saved_count, skipped_count


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="SF6 Bucklerからランク帯別のplay配下データをサンプル収集してJSONに保存する"
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
            print("Buckler のランキングデータ取得には Cookie（ログイン状態）が必要です。")
            print("ブラウザで Buckler にログイン後、Cookie を .buckler_cookie.txt に保存してください。")
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

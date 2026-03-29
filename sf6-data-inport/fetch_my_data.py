"""SF6 個人データ取得スクリプト (Phase 0 タスク3)

ひゅーさん自身のプレイヤーID（Buckler short_id）を引数で受け取り、
play 配下の全データを取得して data/my/ フォルダに JSON で保存する。

使い方:
    python fetch_my_data.py --short-id <あなたのshort_id>
    python fetch_my_data.py --short-id <あなたのshort_id> --dry-run
    python fetch_my_data.py --short-id <あなたのshort_id> --cookie-file .buckler_cookie.txt

短縮形:
    python fetch_my_data.py -i <あなたのshort_id>

注意:
    - short_id はコードにハードコードしていません。必ず引数で渡してください
    - プロフィールページ (https://www.streetfighter.com/6/buckler/profile/<short_id>) の
      URLに含まれる英数字の文字列が short_id です
    - 既存の analyze_step1.py / analyze_playlog.py は変更しません
    - 保存先: data/my/YYYY-MM-DD_<short_id>.json
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date
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

DEFAULT_COOKIE_FILE = ".buckler_cookie.txt"
MY_DATA_DIR = Path("data") / "my"


# ---------------------------------------------------------------------------
# データ取得
# ---------------------------------------------------------------------------

def fetch_play_data(
    short_id: str,
    cookie: str,
    timeout: int,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """プロフィールページから league_info と play 配下の全データを返す。"""
    url = f"{_ranking.BASE_URL}/profile/{short_id}"
    headers = _ranking.make_headers(cookie=cookie, referer=url)

    try:
        html = _ranking.fetch_text(url, headers, timeout)
    except HTTPError as exc:
        if exc.code == 403:
            raise PermissionError(
                f"short_id={short_id}: プロフィールが403で拒否されました。"
                "Cookieが必要な場合は --cookie-file を使ってください。"
            ) from exc
        raise

    next_data = _ranking.extract_next_data(html)
    page_props = next_data.get("props", {}).get("pageProps", {})

    if page_props.get("common", {}).get("statusCode") == 403:
        raise PermissionError(
            f"short_id={short_id}: pageProps.statusCode=403。Cookieが無効かもしれません。"
        )

    league_info = (
        page_props
        .get("fighter_banner_info", {})
        .get("favorite_character_league_info", {})
    )
    play_data: dict[str, Any] = page_props.get("play", {})

    return league_info, play_data


# ---------------------------------------------------------------------------
# JSON 保存
# ---------------------------------------------------------------------------

def save_my_json(json_path: Path, payload: dict[str, Any]) -> None:
    """データをJSONファイルとして保存する（上書き）。"""
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 表示
# ---------------------------------------------------------------------------

def print_summary(short_id: str, league_info: dict[str, Any], battle_stats: dict[str, Any]) -> None:
    """取得したデータの主要指標をCUIで表示する。"""
    today = date.today().strftime("%Y-%m-%d")
    print()
    print("=" * 60)
    print("  個人データ取得結果")
    print("=" * 60)
    print(f"  player_id   : {short_id}")
    print(f"  取得日      : {today}")
    print(f"  LP          : {league_info.get('league_point', '')}")
    print(f"  MR          : {league_info.get('master_rating', '')}")
    print(f"  ランク      : {league_info.get('league_rank', '')}")
    print()
    print("  --- 主要指標 ---")

    def _pct(key: str, label: str) -> None:
        val = battle_stats.get(key, "")
        try:
            formatted = f"{float(val) * 100:.2f}%" if val != "" and val is not None else ""
        except (ValueError, TypeError):
            formatted = str(val)
        print(f"  {label:<22}: {formatted}")

    def _num(key: str, label: str) -> None:
        val = battle_stats.get(key, "")
        print(f"  {label:<22}: {'' if val is None else val}")

    _pct("gauge_rate_drive_guard",   "ドライブパリィ%")
    _num("drive_parry",              "ドライブパリィ回数")
    _num("just_parry",               "ジャストパリィ回数")
    _num("throw_tech",               "投げ抜け回数")
    _num("received_drive_impact",    "DI被弾回数")
    _num("received_punish_counter",  "パニカン被弾回数")
    _num("received_stun",            "スタン被弾回数")
    _num("corner_time",              "端追い詰め時間")
    _num("cornered_time",            "端追い詰められ時間")
    _num("rank_match_play_count",    "ランクマッチ試合数")

    print("=" * 60)
    print()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="自分のBuckler short_idを指定してplay配下の全データを取得・保存する"
    )
    parser.add_argument(
        "--short-id", "-i",
        required=True,
        metavar="SHORT_ID",
        help=(
            "BucklerのプレイヤーID（short_id）。"
            "プロフィールURL の末尾の英数字文字列。"
            "例: https://www.streetfighter.com/6/buckler/profile/XXXXXXXX の XXXXXXXX 部分"
        ),
    )
    parser.add_argument(
        "--cookie",
        default="",
        help="Cookie文字列（任意。省略時は --cookie-file を使用）",
    )
    parser.add_argument(
        "--cookie-file",
        default=DEFAULT_COOKIE_FILE,
        help=f"Cookie文字列を保存したファイル（デフォルト: {DEFAULT_COOKIE_FILE}）",
    )
    parser.add_argument(
        "--output-dir",
        default=str(MY_DATA_DIR),
        help=f"保存先フォルダ（デフォルト: {MY_DATA_DIR}）",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=_ranking.DEFAULT_TIMEOUT,
        help="HTTPタイムアウト秒数",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="取得結果を表示するだけで保存しない",
    )
    args = parser.parse_args()

    # short_id の簡易バリデーション
    short_id = args.short_id.strip()
    if not short_id:
        print("エラー: --short-id が空です")
        sys.exit(1)

    # Cookie 読み込み（任意）
    cookie = _ranking.load_cookie_text(args.cookie, args.cookie_file)
    if cookie:
        try:
            _ranking.validate_cookie_text(cookie)
        except ValueError as exc:
            print(f"エラー: Cookie形式が不正です: {exc}")
            sys.exit(1)

    output_dir = Path(args.output_dir)
    today = date.today().strftime("%Y-%m-%d")
    json_path = output_dir / f"{today}_{short_id}.json"

    print(f"=== SF6 個人データ取得 ===")
    print(f"short_id : {short_id}")
    print(f"保存先   : {json_path}")
    if args.dry_run:
        print("（dry-run モード: 保存しません）")
    print()
    print("データ取得中...")

    try:
        league_info, play_data = fetch_play_data(
            short_id=short_id,
            cookie=cookie,
            timeout=args.timeout,
        )
    except PermissionError as exc:
        print(f"権限エラー: {exc}")
        sys.exit(1)
    except (ValueError, HTTPError, URLError, RuntimeError) as exc:
        print(f"取得エラー: {exc}")
        sys.exit(1)

    battle_stats = play_data.get("battle_stats", {})
    print_summary(short_id, league_info, battle_stats)

    if args.dry_run:
        print("（dry-run モード: 保存をスキップしました）")
        print(f"  取得できたplay配下のキー: {list(play_data.keys())}")
        return

    # JSON保存: collect_samples.py と同一構造（rank は個人データでは空文字）
    payload: dict[str, Any] = {
        "fetch_date": today,
        "player_id": short_id,
        "rank": "",
        "league_info": league_info,
        "play": play_data,
    }

    output_dir.mkdir(parents=True, exist_ok=True)
    save_my_json(json_path, payload)
    print(f"保存完了: {json_path}")
    print(f"  play配下のキー: {list(play_data.keys())}")


if __name__ == "__main__":
    main()

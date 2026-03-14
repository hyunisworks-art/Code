"""SF6 Buckler プロフィールページから実績データを収集するスクリプト。

ranking-output/*.csv から short_id を読み込み、各プレイヤーの
プロフィールページ（battle_stats）を取得して CSV に保存する。
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import time
from datetime import date
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


BASE_URL = "https://www.streetfighter.com/6/buckler"
PROFILE_URL_TEMPLATE = BASE_URL + "/profile/{short_id}"
DEFAULT_DELAY = 1.5
DEFAULT_TIMEOUT = 30
DEFAULT_INPUT_CSV = "ranking-output/master_p1-p3.csv"
DEFAULT_OUTPUT_DIR = "profile-output"
DEFAULT_COOKIE_FILE = ".buckler_cookie.txt"
SHORT_ID_COLUMN = "fighter_banner_info.personal_info.short_id"
PLAYER_NAME_COLUMN = "fighter_banner_info.personal_info.fighter_id"

NEXT_DATA_PATTERN = re.compile(
    r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>'
)
COOKIE_HEADER_PATTERN = re.compile(r"-H\s+['\"]cookie:\s*([^'\"]+)['\"]", re.IGNORECASE)

# battle_stats の全キー（取得順を固定するため明示）
BATTLE_STATS_KEYS = [
    "battle_hub_match_play_count",
    "casual_match_play_count",
    "corner_time",
    "cornered_time",
    "custom_room_match_play_count",
    "drive_impact",
    "drive_impact_to_drive_impact",
    "drive_parry",
    "drive_reversal",
    "gauge_rate_ca",
    "gauge_rate_drive_arts",
    "gauge_rate_drive_guard",
    "gauge_rate_drive_impact",
    "gauge_rate_drive_other",
    "gauge_rate_drive_reversal",
    "gauge_rate_drive_rush_from_cancel",
    "gauge_rate_drive_rush_from_parry",
    "gauge_rate_sa_lv1",
    "gauge_rate_sa_lv2",
    "gauge_rate_sa_lv3",
    "just_parry",
    "punish_counter",
    "rank_match_play_count",
    "received_drive_impact",
    "received_drive_impact_to_drive_impact",
    "received_punish_counter",
    "received_stun",
    "received_throw_count",
    "received_throw_drive_parry",
    "rival_ai_achieved_challenge_count",
    "rival_ai_highest_league_rank",
    "rival_ai_highest_league_rank_txt",
    "stun",
    "target_clear_count",
    "throw_count",
    "throw_drive_parry",
    "throw_tech",
    "total_all_character_play_point",
]

OUTPUT_COLUMNS = (
    ["short_id", "fetch_date", "player_name", "league_point", "master_rating", "league_rank"]
    + BATTLE_STATS_KEYS
)


# ---------------------------------------------------------------------------
# HTTP ユーティリティ（scrape_rankings.py と同じパターン）
# ---------------------------------------------------------------------------

def make_headers(cookie: str = "", referer: str = "") -> dict[str, str]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/137.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        "Cache-Control": "max-age=0",
    }
    if referer:
        headers["Referer"] = referer
    if cookie:
        headers["Cookie"] = cookie
    return headers


def get_curl_command() -> str | None:
    return shutil.which("curl.exe") or shutil.which("curl")


def fetch_bytes_with_curl(url: str, headers: dict[str, str], timeout: int) -> bytes:
    curl_command = get_curl_command()
    if not curl_command:
        raise RuntimeError("curl コマンドが見つかりません")
    command = [curl_command, "-sS", "-L", "--compressed", "--max-time", str(timeout)]
    for key, value in headers.items():
        command.extend(["-H", f"{key}: {value}"])
    command.append(url)
    completed = subprocess.run(command, capture_output=True, check=False)
    if completed.returncode != 0:
        stderr = completed.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(f"curl による取得に失敗しました: {stderr or completed.returncode}")
    return completed.stdout


def fetch_text(url: str, headers: dict[str, str], timeout: int) -> str:
    request = Request(url, headers=headers)
    try:
        with urlopen(request, timeout=timeout) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            raw = response.read()
        return raw.decode(charset, errors="replace")
    except HTTPError as exc:
        if exc.code != 403:
            raise
        raw = fetch_bytes_with_curl(url, headers, timeout)
        return raw.decode("utf-8", errors="replace")


def extract_next_data(html: str) -> dict[str, Any]:
    match = NEXT_DATA_PATTERN.search(html)
    if not match:
        raise ValueError("__NEXT_DATA__ が見つかりませんでした")
    return json.loads(match.group(1))


# ---------------------------------------------------------------------------
# Cookie ユーティリティ（scrape_rankings.py と同じパターン）
# ---------------------------------------------------------------------------

def load_cookie_text(explicit_cookie: str, cookie_file: str) -> str:
    if explicit_cookie.strip():
        return normalize_cookie_text(explicit_cookie.strip())
    env_cookie = os.environ.get("BUCKLER_COOKIE", "").strip()
    if env_cookie:
        return normalize_cookie_text(env_cookie)
    file_path = Path(cookie_file)
    if file_path.exists():
        return normalize_cookie_text(file_path.read_text(encoding="utf-8", errors="replace").strip())
    return ""


def normalize_cookie_text(raw_text: str) -> str:
    text = raw_text.strip()
    if not text:
        return ""
    if "curl" in text.lower() and "-h" in text.lower():
        match = COOKIE_HEADER_PATTERN.search(text)
        if match:
            return match.group(1).strip()
    for line in text.splitlines():
        if line.lower().startswith("cookie:"):
            return line.split(":", 1)[1].strip()
    if text.lower().startswith("cookie:"):
        return text.split(":", 1)[1].strip()
    if text.startswith("{stamp:") and ";" in text:
        _, remainder = text.split(";", 1)
        return remainder.strip()
    return text


def validate_cookie_text(cookie: str) -> None:
    if not cookie:
        return
    stripped = cookie.strip()
    if "=" not in stripped:
        raise ValueError(
            "Cookie文字列の形式が不正です。name=value 形式のCookieが含まれていません。"
        )


# ---------------------------------------------------------------------------
# short_id 収集
# ---------------------------------------------------------------------------

def load_short_ids_from_csv(csv_path: Path) -> list[tuple[str, str]]:
    """ランキングCSVから (short_id, player_name) の一覧を返す。重複は除去する。"""
    if not csv_path.exists():
        raise FileNotFoundError(f"ランキングCSVが見つかりません: {csv_path}")

    seen: dict[str, str] = {}
    with csv_path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            sid = row.get(SHORT_ID_COLUMN, "").strip()
            name = row.get(PLAYER_NAME_COLUMN, "").strip()
            if sid and sid not in seen:
                seen[sid] = name
    return list(seen.items())


# ---------------------------------------------------------------------------
# プロフィール取得
# ---------------------------------------------------------------------------

def fetch_profile(short_id: str, cookie: str, timeout: int) -> dict[str, Any]:
    """プロフィールページを取得して __NEXT_DATA__ の pageProps を返す。"""
    url = PROFILE_URL_TEMPLATE.format(short_id=short_id)
    headers = make_headers(cookie=cookie, referer=url)
    try:
        html = fetch_text(url, headers, timeout)
    except HTTPError as exc:
        if exc.code == 403:
            raise PermissionError(
                f"short_id={short_id} のプロフィールが 403 で拒否されました。"
                "Cookie を --cookie / BUCKLER_COOKIE / --cookie-file で渡してください。"
            ) from exc
        raise
    next_data = extract_next_data(html)
    page_props = next_data.get("props", {}).get("pageProps", {})
    status = page_props.get("common", {}).get("statusCode")
    if status == 403:
        raise PermissionError(
            f"short_id={short_id}: pageProps.common.statusCode=403。Cookie が無効な可能性があります。"
        )
    return page_props


def extract_profile_row(short_id: str, player_name: str, page_props: dict[str, Any]) -> dict[str, str]:
    """pageProps から出力列に対応する dict を作る。"""
    league_info = (
        page_props.get("fighter_banner_info", {})
        .get("favorite_character_league_info", {})
    )
    battle_stats: dict[str, Any] = page_props.get("play", {}).get("battle_stats", {})

    row: dict[str, str] = {}
    today = date.today()
    row["short_id"] = short_id
    row["fetch_date"] = f"{today.year}/{today.month}/{today.day}"
    row["player_name"] = str(league_info.get("personal_info", {}).get("fighter_id", player_name) or player_name)
    row["league_point"] = str(league_info.get("league_point", ""))
    row["master_rating"] = str(league_info.get("master_rating", ""))
    row["league_rank"] = str(league_info.get("league_rank", ""))

    for key in BATTLE_STATS_KEYS:
        val = battle_stats.get(key, "")
        row[key] = "" if val is None else str(val)

    return row


# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------

def scrape_profiles(
    input_csv: Path,
    output_dir: Path,
    cookie: str,
    delay: float,
    timeout: int,
    limit: int | None,
    dry_run: bool,
) -> None:
    pairs = load_short_ids_from_csv(input_csv)
    if limit is not None:
        pairs = pairs[:limit]

    print(f"対象: {len(pairs)} 件 (入力: {input_csv})")

    if dry_run:
        print("--dry-run 指定のため取得は行いません。")
        for sid, name in pairs:
            print(f"  {sid}  {name}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = date.today().strftime("%Y%m%d")
    out_path = output_dir / f"profiles_{timestamp}.csv"

    # 既存ファイルがあれば追記モードにして重複をスキップ
    existing_ids: set[str] = set()
    write_mode = "w"
    if out_path.exists():
        with out_path.open(encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for r in reader:
                sid = r.get("short_id", "").strip()
                if sid:
                    existing_ids.add(sid)
        write_mode = "a"
        print(f"既存ファイル {out_path} に追記します（既存 {len(existing_ids)} 件をスキップ）")

    success = 0
    skip = 0
    error = 0

    with out_path.open(write_mode, encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS, extrasaction="ignore")
        if write_mode == "w":
            writer.writeheader()

        for i, (sid, name) in enumerate(pairs):
            if sid in existing_ids:
                skip += 1
                continue

            try:
                page_props = fetch_profile(sid, cookie, timeout)
                row = extract_profile_row(sid, name, page_props)
                writer.writerow(row)
                f.flush()
                success += 1
                lp = row.get("league_point", "")
                mr = row.get("master_rating", "")
                rank_cnt = row.get("rank_match_play_count", "")
                print(f"[{i+1}/{len(pairs)}] {sid} {name}  LP={lp} MR={mr} ランクマッチ={rank_cnt}")
            except PermissionError as exc:
                print(f"[{i+1}/{len(pairs)}] ERROR {sid}: {exc}")
                error += 1
            except (ValueError, HTTPError, URLError, RuntimeError) as exc:
                print(f"[{i+1}/{len(pairs)}] ERROR {sid}: {exc}")
                error += 1

            if i + 1 < len(pairs):
                time.sleep(delay)

    print()
    print(f"完了: 取得={success}  スキップ={skip}  エラー={error}")
    print(f"csv={out_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="ランキングCSVの short_id からプロフィール実績を収集して保存する"
    )
    parser.add_argument(
        "--input-csv",
        default=DEFAULT_INPUT_CSV,
        help=f"short_id の入力元ランキングCSV（既定: {DEFAULT_INPUT_CSV}）",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"出力先フォルダ（既定: {DEFAULT_OUTPUT_DIR}）",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY,
        help=f"リクエスト間の待機秒数（既定: {DEFAULT_DELAY}）",
    )
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT)
    parser.add_argument(
        "--cookie",
        default="",
        help="ブラウザから取得したCookie文字列（任意）",
    )
    parser.add_argument(
        "--cookie-file",
        default=DEFAULT_COOKIE_FILE,
        help=f"Cookie文字列を保存したファイル（既定: {DEFAULT_COOKIE_FILE}）",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        metavar="N",
        help="取得件数の上限（省略時: 全件）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="件数と short_id を確認するだけで取得しない",
    )
    args = parser.parse_args()

    cookie = load_cookie_text(args.cookie, args.cookie_file)
    validate_cookie_text(cookie)

    scrape_profiles(
        input_csv=Path(args.input_csv),
        output_dir=Path(args.output_dir),
        cookie=cookie,
        delay=args.delay,
        timeout=args.timeout,
        limit=args.limit,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()

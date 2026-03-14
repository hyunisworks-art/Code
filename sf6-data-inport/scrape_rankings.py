from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import time
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


BASE_URL = "https://www.streetfighter.com/6/buckler"
DEFAULT_LOCALE = "en"
DEFAULT_DELAY = 1.2
DEFAULT_TIMEOUT = 30
DEFAULT_OUTPUT_DIR = "ranking-output"
DEFAULT_COOKIE_FILE = ".buckler_cookie.txt"
NEXT_DATA_PATTERN = re.compile(
    r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>'
)
RANKING_PAGE_KEYS = {
    "master": "master_rating_ranking",
    "league": "league_point_ranking",
}
COOKIE_HEADER_PATTERN = re.compile(r"-H\s+['\"]cookie:\s*([^'\"]+)['\"]", re.IGNORECASE)


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
            content_type = response.headers.get_content_charset() or "utf-8"
            raw = response.read()
        return raw.decode(content_type, errors="replace")
    except HTTPError as exc:
        if exc.code != 403:
            raise
        raw = fetch_bytes_with_curl(url, headers, timeout)
        return raw.decode("utf-8", errors="replace")


def fetch_json(url: str, headers: dict[str, str], timeout: int) -> dict[str, Any]:
    json_headers = dict(headers)
    json_headers["Accept"] = "application/json,text/plain,*/*"
    request = Request(url, headers=json_headers)
    try:
        with urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8", errors="replace"))
    except HTTPError as exc:
        if exc.code != 403:
            raise
        raw = fetch_bytes_with_curl(url, json_headers, timeout)
        return json.loads(raw.decode("utf-8", errors="replace"))


def extract_next_data(html: str) -> dict[str, Any]:
    match = NEXT_DATA_PATTERN.search(html)
    if not match:
        raise ValueError("__NEXT_DATA__ を公開HTMLから取得できませんでした")
    return json.loads(match.group(1))


def get_build_id(html: str) -> str:
    next_data = extract_next_data(html)
    build_id = str(next_data.get("buildId", "")).strip()
    if not build_id:
        raise ValueError("buildId を取得できませんでした")
    return build_id


def build_ranking_page_url(ranking_type: str, page: int, locale: str) -> str:
    if locale == DEFAULT_LOCALE:
        return f"{BASE_URL}/ranking/{ranking_type}?page={page}"
    return f"{BASE_URL}/{locale}/ranking/{ranking_type}?page={page}"


def build_next_data_url(build_id: str, ranking_type: str, page: int, locale: str) -> str:
    params = urlencode({"page": page})
    return f"{BASE_URL}/_next/data/{build_id}/{locale}/ranking/{ranking_type}.json?{params}"


def get_ranking_payload(page_props: dict[str, Any], ranking_type: str) -> dict[str, Any]:
    expected_key = RANKING_PAGE_KEYS[ranking_type]
    payload = page_props.get(expected_key)
    if isinstance(payload, dict):
        return payload

    for key, value in page_props.items():
        if key.endswith("_ranking") and isinstance(value, dict):
            return value

    status = page_props.get("common", {}).get("statusCode")
    if status == 403:
        raise PermissionError(
            "ランキングJSONの取得が 403 で拒否されました。"
            "この状態は Buckler 側で未ログイン扱いのときに発生します。"
            "ブラウザで Buckler にログイン後、Cookie を --cookie / BUCKLER_COOKIE / --cookie-file のいずれかで渡してください。"
        )

    raise ValueError("ランキングデータ本体を pageProps から見つけられませんでした")


def flatten_item(item: Any, prefix: str = "") -> dict[str, str]:
    flattened: dict[str, str] = {}

    if isinstance(item, dict):
        for key, value in item.items():
            child_prefix = f"{prefix}.{key}" if prefix else str(key)
            flattened.update(flatten_item(value, child_prefix))
        return flattened

    if isinstance(item, list):
        if not prefix:
            return {"value": json.dumps(item, ensure_ascii=False)}
        return {prefix: json.dumps(item, ensure_ascii=False)}

    if prefix:
        flattened[prefix] = "" if item is None else str(item)
    return flattened


def choose_columns(rows: list[dict[str, str]]) -> list[str]:
    seen: dict[str, None] = {}
    for row in rows:
        for key in row:
            seen.setdefault(key, None)
    return list(seen.keys())


def write_jsonl(path: Path, rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8") as file:
        for row in rows:
            file.write(json.dumps(row, ensure_ascii=False) + "\n")


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    columns = choose_columns(rows)
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def detect_total_pages(payload: dict[str, Any], collected_pages: int) -> int:
    for key in ("total_page", "total_pages"):
        value = payload.get(key)
        if isinstance(value, int) and value > 0:
            return value
        if isinstance(value, str) and value.isdigit():
            return int(value)
    return collected_pages


def scrape_rankings(
    ranking_type: str,
    start_page: int,
    end_page: int,
    locale: str,
    delay: float,
    timeout: int,
    cookie: str,
) -> tuple[list[dict[str, str]], dict[str, Any]]:
    first_page_url = build_ranking_page_url(ranking_type, start_page, locale)
    html_headers = make_headers(cookie=cookie, referer=first_page_url)
    try:
        html = fetch_text(first_page_url, html_headers, timeout)
    except HTTPError as exc:
        if exc.code == 403:
            raise PermissionError(
                "公開ランキングHTMLの取得が 403 で拒否されました。"
                "ブラウザCookieを --cookie または BUCKLER_COOKIE で指定して再試行してください。"
            ) from exc
        raise
    build_id = get_build_id(html)

    rows: list[dict[str, str]] = []
    total_pages = end_page

    for page in range(start_page, end_page + 1):
        page_url = build_ranking_page_url(ranking_type, page, locale)
        api_url = build_next_data_url(build_id, ranking_type, page, locale)
        headers = make_headers(cookie=cookie, referer=page_url)

        try:
            data = fetch_json(api_url, headers, timeout)
        except HTTPError as exc:
            if exc.code == 403:
                raise PermissionError(
                    "ランキングJSONの取得が 403 で拒否されました。"
                    "公開ページのみを読みに行く実装ですが、CloudFront 側の制御で Cookie が必要な場合があります。"
                ) from exc
            raise
        except URLError as exc:
            raise ConnectionError(f"ランキングJSONの取得に失敗しました: {exc}") from exc

        page_props = data.get("pageProps", {})
        payload = get_ranking_payload(page_props, ranking_type)
        total_pages = detect_total_pages(payload, total_pages)
        ranking_items = payload.get("ranking_fighter_list", [])
        if not isinstance(ranking_items, list):
            raise ValueError("ranking_fighter_list の形式が不正です")

        for item in ranking_items:
            if not isinstance(item, dict):
                continue
            flattened = flatten_item(item)
            flattened["source.page"] = str(page)
            flattened["source.ranking_type"] = ranking_type
            flattened["source.locale"] = locale
            rows.append(flattened)

        if page < end_page:
            time.sleep(delay)

    metadata = {
        "ranking_type": ranking_type,
        "locale": locale,
        "start_page": start_page,
        "end_page": end_page,
        "total_pages_detected": total_pages,
        "build_id": build_id,
        "row_count": len(rows),
    }
    return rows, metadata


def write_metadata(path: Path, metadata: dict[str, Any]) -> None:
    path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")


def load_cookie_text(explicit_cookie: str, cookie_file: str) -> str:
    source = ""
    if explicit_cookie.strip():
        source = explicit_cookie.strip()
        return normalize_cookie_text(source)

    env_cookie = os.environ.get("BUCKLER_COOKIE", "").strip()
    if env_cookie:
        source = env_cookie
        return normalize_cookie_text(source)

    file_path = Path(cookie_file)
    if file_path.exists():
        source = file_path.read_text(encoding="utf-8", errors="replace").strip()
        return normalize_cookie_text(source)

    return ""


def normalize_cookie_text(raw_text: str) -> str:
    text = raw_text.strip()
    if not text:
        return ""

    # Edge/Chrome の "Copy as cURL" をそのまま貼り付けた場合に対応
    if "curl" in text.lower() and "-h" in text.lower():
        match = COOKIE_HEADER_PATTERN.search(text)
        if match:
            return match.group(1).strip()

    # Request Headers を貼り付けた場合に対応
    for line in text.splitlines():
        if line.lower().startswith("cookie:"):
            return line.split(":", 1)[1].strip()

    # cookie: プレフィックス単体にも対応
    if text.lower().startswith("cookie:"):
        return text.split(":", 1)[1].strip()

    # 先頭に同意管理の値だけが混ざるケースを除去
    # 例: {stamp:...necessary:true...}; name1=value1; name2=value2
    if text.startswith("{stamp:") and ";" in text:
        _, remainder = text.split(";", 1)
        return remainder.strip()

    return text


def validate_cookie_text(cookie: str) -> None:
    if not cookie:
        return

    stripped = cookie.strip()
    lowered = stripped.lower()

    if "=" not in stripped:
        if (
            "necessary:true" in lowered
            or "statistics:true" in lowered
            or "marketing:true" in lowered
            or stripped.startswith("{stamp:")
        ):
            raise ValueError(
                "現在のCookie文字列は同意管理用の値で、認証Cookieではありません。"
                "Network の Request Headers から cookie: の値全体をコピーしてください。"
            )
        raise ValueError(
            "Cookie文字列の形式が不正です。name=value 形式のCookieが含まれていません。"
            "Edge の Network から 'Copy as cURL' した全文を貼り付けても利用できます。"
        )

    parts = [part.strip() for part in stripped.split(";") if part.strip()]
    invalid_parts = [
        part
        for part in parts
        if "=" not in part and not part.startswith("{stamp:")
    ]
    if invalid_parts:
        raise ValueError(
            "Cookie文字列の形式が不正です。"
            "';' 区切りの各要素は name=value 形式である必要があります。"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="SF6 Buckler の公開ランキングを取得して保存する")
    parser.add_argument("--ranking-type", choices=("master", "league"), default="master")
    parser.add_argument("--start-page", type=int, default=1)
    parser.add_argument("--end-page", type=int, default=3)
    parser.add_argument("--locale", default=DEFAULT_LOCALE, help="既定: en")
    parser.add_argument("--delay", type=float, default=DEFAULT_DELAY, help="ページ間の待機秒数")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT)
    parser.add_argument("--cookie", default="", help="ブラウザから取得したCookie文字列（任意）")
    parser.add_argument(
        "--cookie-file",
        default=DEFAULT_COOKIE_FILE,
        help="Cookie文字列を保存したファイル (既定: .buckler_cookie.txt)",
    )
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR)
    args = parser.parse_args()

    if args.start_page < 1 or args.end_page < args.start_page:
        raise ValueError("ページ範囲が不正です")

    cookie = load_cookie_text(args.cookie, args.cookie_file)
    validate_cookie_text(cookie)
    rows, metadata = scrape_rankings(
        ranking_type=args.ranking_type,
        start_page=args.start_page,
        end_page=args.end_page,
        locale=args.locale.strip() or DEFAULT_LOCALE,
        delay=args.delay,
        timeout=args.timeout,
        cookie=cookie,
    )

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    stem = f"{args.ranking_type}_p{args.start_page}-p{args.end_page}"
    jsonl_path = output_dir / f"{stem}.jsonl"
    csv_path = output_dir / f"{stem}.csv"
    meta_path = output_dir / f"{stem}.meta.json"

    write_jsonl(jsonl_path, rows)
    write_csv(csv_path, rows)
    write_metadata(meta_path, metadata)

    print(f"rows={len(rows)}")
    print(f"csv={csv_path}")
    print(f"jsonl={jsonl_path}")
    print(f"meta={meta_path}")


if __name__ == "__main__":
    main()
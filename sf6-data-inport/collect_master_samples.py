"""マスター帯のサブ段階別サンプル収集スクリプト。

ランキングの後ページを取得し、MR値で4段階に分類して保存する。
既存のcollect_samples.pyとは別に、マスター帯専用で動かす。

使い方:
    python collect_master_samples.py
    python collect_master_samples.py --pages 50-300
    python collect_master_samples.py --target 15 --dry-run

MR段階:
    MASTER:   MR < 1600
    HIGH:     MR 1600-1699
    GRAND:    MR 1700-1799
    ULTIMATE: MR 1800+
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import date
from pathlib import Path
from typing import Any

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

import scrape_rankings as sr

SAMPLES_DIR = Path("data") / "samples"
DEFAULT_COOKIE_FILE = ".buckler_cookie.txt"
DELAY = 3.0
MAX_REQUESTS = 500  # マスター帯専用なので大きめに

MR_BANDS = {
    "master":   (0, 1599),
    "high":     (1600, 1699),
    "grand":    (1700, 1799),
    "ultimate": (1800, 99999),
}

TARGET_PER_BAND = 15  # 各段階の目標件数（サブランク均等化のため）


def classify_mr(mr: int) -> str:
    if mr >= 1800:
        return "ultimate"
    elif mr >= 1700:
        return "grand"
    elif mr >= 1600:
        return "high"
    else:
        return "master"


def load_existing_master_ids() -> set[str]:
    """既存のmasterサンプルのplayer_idセットを返す。"""
    if not SAMPLES_DIR.exists():
        return set()
    ids = set()
    for p in SAMPLES_DIR.glob("*.json"):
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            if data.get("rank", "").startswith("master"):
                ids.add(str(data.get("player_id", "")))
        except Exception:
            continue
    return ids


def count_existing_by_band() -> dict[str, int]:
    """既存マスターサンプルのMR段階別件数を返す。"""
    counts = {k: 0 for k in MR_BANDS}
    if not SAMPLES_DIR.exists():
        return counts
    for p in SAMPLES_DIR.glob("*.json"):
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            if not data.get("rank", "").startswith("master"):
                continue
            cli = (data.get("play") or {}).get("character_league_infos") or []
            max_mr = max(
                ((c.get("league_info") or {}).get("master_rating") or 0) for c in cli
            ) if cli else 0
            if max_mr > 0:
                band = classify_mr(max_mr)
                counts[band] += 1
        except Exception:
            continue
    return counts


def fetch_play_data(short_id: str, cookie: str, timeout: int) -> dict[str, Any] | None:
    """プレイヤーのplayデータを取得する。"""
    try:
        url = f"{sr.BASE_URL}/profile/{short_id}"
        headers = sr.make_headers(cookie=cookie, referer=url)
        html = sr.fetch_text(url, headers, timeout)
        next_data = sr.extract_next_data(html)
        page_props = next_data.get("props", {}).get("pageProps", {})
        if page_props.get("common", {}).get("statusCode") == 403:
            return None
        return page_props.get("play", {})
    except Exception as exc:
        print(f"    取得失敗: {exc}")
        return None


def main() -> None:
    parser = argparse.ArgumentParser(description="マスター帯サブ段階別サンプル収集")
    parser.add_argument("--pages", default="10-300", help="ランキングページ範囲 (例: 50-300)")
    parser.add_argument("--target", type=int, default=TARGET_PER_BAND, help="各段階の目標件数")
    parser.add_argument("--dry-run", action="store_true", help="IDリスト収集のみ")
    args = parser.parse_args()

    page_start, page_end = map(int, args.pages.split("-"))
    target = args.target

    # Cookie読み込み
    try:
        cookie = open(DEFAULT_COOKIE_FILE).read().strip()
    except FileNotFoundError:
        print(f"エラー: {DEFAULT_COOKIE_FILE} が見つかりません")
        return

    # 既存データ確認
    existing_ids = load_existing_master_ids()
    band_counts = count_existing_by_band()
    print("現在のマスター帯サンプル:")
    for band, count in band_counts.items():
        need = max(0, target - count)
        print(f"  {band:10s}: {count}件 (あと{need}件必要)")

    # 各段階であと何件必要か
    needs = {band: max(0, target - count) for band, count in band_counts.items()}
    if all(n == 0 for n in needs.values()):
        print("\n全段階で目標件数を達成済みです。")
        return

    # build_id取得
    url1 = sr.build_ranking_page_url("master", 1, sr.DEFAULT_LOCALE)
    headers = sr.make_headers(cookie=cookie, referer=url1)
    html = sr.fetch_text(url1, headers, 30)
    build_id = sr.get_build_id(html)

    # ランキングページをスキャンしてID + MR段階を収集
    candidates: dict[str, list[tuple[str, int]]] = {k: [] for k in MR_BANDS}  # band -> [(short_id, mr)]
    request_count = 1
    today = date.today().strftime("%Y-%m-%d")

    print(f"\nランキングページ {page_start}-{page_end} をスキャン中...")
    for page in range(page_start, page_end + 1):
        if all(len(candidates[b]) >= needs[b] for b in needs if needs[b] > 0):
            print("全段階で候補が十分集まりました。")
            break

        if request_count >= MAX_REQUESTS:
            print(f"リクエスト上限({MAX_REQUESTS})に達しました。")
            break

        page_url = sr.build_ranking_page_url("master", page, sr.DEFAULT_LOCALE, league_rank=36)
        api_url = sr.build_next_data_url(build_id, "master", page, sr.DEFAULT_LOCALE, league_rank=36)
        api_headers = sr.make_headers(cookie=cookie, referer=page_url)

        request_count += 1
        try:
            data = sr.fetch_json(api_url, api_headers, 30)
        except Exception as exc:
            print(f"  Page {page}: エラー {exc}")
            time.sleep(DELAY)
            continue

        pp = data.get("pageProps", {})
        try:
            payload = sr.get_ranking_payload(pp, "master")
        except Exception:
            time.sleep(DELAY)
            continue

        items = payload.get("ranking_fighter_list", [])
        if not items:
            print(f"  Page {page}: 空ページ。終了。")
            break

        for item in items:
            flat = sr.flatten_item(item)
            short_id = flat.get("fighter_banner_info.personal_info.short_id", "").strip()
            mr = int(flat.get("fighter_banner_info.favorite_character_league_info.master_rating", 0) or 0)

            if not short_id or short_id in existing_ids:
                continue
            if mr <= 0:
                continue

            band = classify_mr(mr)
            if needs[band] > 0 and len(candidates[band]) < needs[band]:
                candidates[band].append((short_id, mr))

        if page % 20 == 0:
            status = ", ".join(f"{b}:{len(candidates[b])}/{needs[b]}" for b in MR_BANDS if needs[b] > 0)
            print(f"  Page {page}: {status}")

        time.sleep(DELAY)

    # 候補サマリー
    print("\n候補収集結果:")
    for band in MR_BANDS:
        print(f"  {band:10s}: {len(candidates[band])}件の候補")

    if args.dry_run:
        print("\n(dry-run: データ取得はスキップ)")
        return

    # playデータ取得と保存
    print("\nplayデータ取得・保存中...")
    for band in MR_BANDS:
        if not candidates[band]:
            continue
        print(f"\n  --- {band} ---")
        saved = 0
        for short_id, mr in candidates[band]:
            if request_count >= MAX_REQUESTS:
                print(f"  リクエスト上限に達しました。")
                break

            print(f"  [{saved+1}/{len(candidates[band])}] {short_id} (MR={mr})...", end=" ", flush=True)
            request_count += 1
            play_data = fetch_play_data(short_id, cookie, 30)

            if play_data is None:
                print("スキップ")
                time.sleep(DELAY)
                continue

            # master_XXX としてランク名にサブ段階を含める
            payload = {
                "fetch_date": today,
                "player_id": short_id,
                "rank": f"master_{band}",
                "league_info": {},
                "play": play_data,
            }

            SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
            json_path = SAMPLES_DIR / f"{today}_{short_id}.json"
            with json_path.open("w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)

            existing_ids.add(short_id)
            saved += 1
            print("OK")
            time.sleep(DELAY)

        print(f"  {band}: {saved}件保存")

    # 最終結果
    final_counts = count_existing_by_band()
    print("\n最終結果:")
    for band, count in final_counts.items():
        print(f"  {band:10s}: {count}件")


if __name__ == "__main__":
    main()

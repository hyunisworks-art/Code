"""MR 1500-1800 帯のサンプルを追加収集する（記事#4 補充用）。
1. page範囲を指定して master ランキングから short_id リスト取得
2. 表示MR 1500-1800 でフィルタ
3. 既存Supabase DB（data_type=sample）と重複チェック
4. play データ取得（Buckler profile API）
5. max(MR) 1500-1800 のサンプルのみ Supabaseに投入
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import date
from pathlib import Path
from urllib.error import HTTPError, URLError

from dotenv import load_dotenv

import scrape_rankings as r

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

MR_MIN = 1500
MR_MAX = 1800
MAX_REQUESTS = 50
DELAY = 2.5
RANK_LABEL = "master_mrband"  # DB上の rank 値（既存と区別）

OUT_DIR = Path("data/samples")


def extract_short_id_and_mr(item):
    fb = item.get("fighter_banner_info") or {}
    pi = fb.get("personal_info") or {}
    short_id = str(pi.get("short_id", "")).strip()
    ci = fb.get("favorite_character_league_info") or {}
    mr = ci.get("master_rating") or 0
    return short_id, mr


def fetch_short_ids_in_range(cookie, pages, request_counter):
    """ページ範囲からMR 1500-1800の表示MRを持つshort_idリストを返す"""
    first_url = r.build_ranking_page_url("master", pages[0], "en")
    headers = r.make_headers(cookie=cookie, referer=first_url)
    request_counter[0] += 1
    html = r.fetch_text(first_url, headers, 30)
    build_id = r.get_build_id(html)

    short_ids = []
    seen = set()
    for page in pages:
        if request_counter[0] >= MAX_REQUESTS:
            print(f"  セッション上限到達・ランキング取得停止")
            break
        api_url = r.build_next_data_url(build_id, "master", page, "en")
        api_headers = r.make_headers(cookie=cookie, referer=r.build_ranking_page_url("master", page, "en"))
        request_counter[0] += 1
        try:
            data = r.fetch_json(api_url, api_headers, 30)
        except Exception as e:
            print(f"  page={page} ランキング取得失敗: {e}")
            continue
        payload = r.get_ranking_payload(data.get("pageProps", {}), "master")
        items = payload.get("ranking_fighter_list", []) or []

        kept_in_page = 0
        for it in items:
            if not isinstance(it, dict):
                continue
            sid, mr = extract_short_id_and_mr(it)
            if not sid or sid in seen:
                continue
            if not (MR_MIN <= mr <= MR_MAX):
                continue
            seen.add(sid)
            short_ids.append((sid, mr))
            kept_in_page += 1
        print(f"  page={page}: {len(items)}件中 {kept_in_page}件が表示MR 1500-1800")
        time.sleep(DELAY)

    return short_ids


def fetch_existing_player_ids():
    from supabase import create_client
    c = create_client(os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_KEY"))
    ids = set()
    start = 0
    while True:
        r = c.table("player_data").select("player_id").eq("data_type", "sample").range(start, start+999).execute()
        if not r.data:
            break
        for row in r.data:
            ids.add(str(row.get("player_id") or ""))
        if len(r.data) < 1000:
            break
        start += 1000
    return ids


def fetch_play(short_id, cookie, request_counter):
    import collect_samples as cs
    return cs._fetch_play_data_with_retry(
        short_id=short_id,
        cookie=cookie,
        timeout=30,
        max_retries=2,
        request_counter=request_counter,
        delay=DELAY,
    )


def max_mr_from_play(play):
    cli = (play or {}).get("character_league_infos") or []
    mrs = []
    for c in cli:
        if not c.get("is_played"):
            continue
        li = c.get("league_info") or {}
        mr = li.get("master_rating", 0)
        if isinstance(mr, (int, float)) and mr > 0:
            mrs.append(mr)
    return max(mrs) if mrs else None


def upsert_supabase(payloads):
    if not payloads:
        return 0
    from supabase import create_client
    c = create_client(os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_KEY"))
    rows = []
    for p in payloads:
        rows.append({
            "player_id": p["player_id"],
            "fetch_date": p["fetch_date"],
            "rank": RANK_LABEL,
            "data_type": "sample",
            "league_info": {},
            "play": p["play"],
        })
    result = c.table("player_data").upsert(
        rows,
        on_conflict="player_id,fetch_date,data_type",
    ).execute()
    return len(result.data)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--page-start", type=int, default=2995)
    parser.add_argument("--page-end", type=int, default=3005)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    cookie = r.load_cookie_text("", ".buckler_cookie.txt")
    r.validate_cookie_text(cookie)
    today = date.today().strftime("%Y-%m-%d")

    print(f"=== MR帯 1500-1800 サンプル収集 ({args.page_start}〜{args.page_end}) ===")
    print(f"1. 既存DB player_id 取得中...")
    existing = fetch_existing_player_ids()
    print(f"  既存: {len(existing)}件")

    print(f"2. ランキングから short_id 抽出...")
    request_counter = [0]
    candidates = fetch_short_ids_in_range(
        cookie, list(range(args.page_start, args.page_end + 1)), request_counter
    )
    print(f"  候補: {len(candidates)}件（表示MR 1500-1800）")

    # 既存除外
    new_candidates = [(sid, mr) for sid, mr in candidates if sid not in existing]
    print(f"  既存除外後: {len(new_candidates)}件")

    if args.dry_run:
        print("(dry-run: play取得スキップ)")
        return

    print(f"3. play データ取得（残りリクエスト枠: {MAX_REQUESTS - request_counter[0]}）...")
    saved_payloads = []
    rejected_out_of_range = 0
    rejected_no_mr = 0

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for idx, (sid, display_mr) in enumerate(new_candidates, 1):
        if request_counter[0] >= MAX_REQUESTS:
            print(f"  セッション上限到達・play取得停止")
            break
        print(f"  [{idx}/{len(new_candidates)}] sid={sid} 表示MR={display_mr} 取得中...", end=" ", flush=True)
        play = fetch_play(sid, cookie, request_counter)
        if play is None:
            print("失敗")
            continue
        max_mr = max_mr_from_play(play)
        if max_mr is None:
            print(f"NG(MR取得不可)")
            rejected_no_mr += 1
            continue
        if not (MR_MIN <= max_mr <= MR_MAX):
            print(f"NG(max_mr={max_mr} 範囲外)")
            rejected_out_of_range += 1
            continue
        print(f"OK(max_mr={max_mr})")
        saved_payloads.append({
            "player_id": sid,
            "fetch_date": today,
            "play": play,
        })
        # ローカルにもバックアップ保存
        (OUT_DIR / f"{today}_{sid}.json").write_text(
            json.dumps({
                "fetch_date": today,
                "player_id": sid,
                "rank": RANK_LABEL,
                "league_info": {},
                "play": play,
            }, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        time.sleep(DELAY)

    print(f"\n4. Supabase投入...")
    uploaded = upsert_supabase(saved_payloads)
    print(f"  投入: {uploaded}件")

    print(f"\n=== サマリー ===")
    print(f"  候補（表示MR範囲内）: {len(candidates)}")
    print(f"  既存除外後: {len(new_candidates)}")
    print(f"  Supabase投入: {uploaded}")
    print(f"  範囲外で破棄: {rejected_out_of_range}")
    print(f"  MR取得不可で破棄: {rejected_no_mr}")
    print(f"  総リクエスト: {request_counter[0]}")


if __name__ == "__main__":
    main()

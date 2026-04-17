"""マスターランキングの飛び飛びページでMR分布を見る。
MR 1500-1800 帯がどのページにあるか特定する。
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import scrape_rankings as r

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

PROBE_PAGES = [1, 100, 500, 1000, 3000, 8000, 15000, 25000, 35000]


def main():
    cookie = r.load_cookie_text("", ".buckler_cookie.txt")
    r.validate_cookie_text(cookie)

    first_url = r.build_ranking_page_url("master", 1, "en")
    headers = r.make_headers(cookie=cookie, referer=first_url)
    html = r.fetch_text(first_url, headers, 30)
    build_id = r.get_build_id(html)

    print(f"build_id: {build_id}")
    print(f"--- MR分布探索 ---")
    print(f"{'page':>6} | {'min_MR':>7} | {'max_MR':>7} | {'min_rank':>9} | {'max_rank':>9}")

    for page in PROBE_PAGES:
        api_url = r.build_next_data_url(build_id, "master", page, "en")
        api_headers = r.make_headers(cookie=cookie, referer=r.build_ranking_page_url("master", page, "en"))
        try:
            data = r.fetch_json(api_url, api_headers, 30)
        except Exception as e:
            print(f"{page:>6} | ERROR: {e}")
            time.sleep(2)
            continue

        page_props = data.get("pageProps", {})
        payload = r.get_ranking_payload(page_props, "master")
        items = payload.get("ranking_fighter_list", [])

        mrs, ranks = [], []
        for it in items:
            if not isinstance(it, dict):
                continue
            fb = (it.get("fighter_banner_info") or {}).get("favorite_character_league_info") or {}
            mr = fb.get("master_rating")
            rk = fb.get("master_rating_ranking")
            if mr:
                mrs.append(mr)
            if rk:
                ranks.append(rk)

        if mrs:
            print(f"{page:>6} | {min(mrs):>7} | {max(mrs):>7} | {min(ranks):>9} | {max(ranks):>9}")
        else:
            print(f"{page:>6} | no data")

        time.sleep(2)  # Buckler負荷軽減


if __name__ == "__main__":
    main()

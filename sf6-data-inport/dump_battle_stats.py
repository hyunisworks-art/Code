"""Buckler battle_stats の全フィールドを確認するスクリプト

使い方:
    python dump_battle_stats.py --short-id YOUR_SHORT_ID
    python dump_battle_stats.py --short-id YOUR_SHORT_ID --cookie-file .buckler_cookie.txt
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

import scrape_rankings as _ranking

# collect_playlog.py で現在マッピング済みのフィールド
MAPPED_FIELDS = {
    "gauge_rate_drive_guard",
    "gauge_rate_drive_impact",
    "gauge_rate_drive_arts",
    "gauge_rate_drive_rush_from_parry",
    "gauge_rate_drive_rush_from_cancel",
    "gauge_rate_drive_reversal",
    "gauge_rate_drive_other",
    "gauge_rate_sa_lv1",
    "gauge_rate_sa_lv2",
    "gauge_rate_sa_lv3",
    "gauge_rate_ca",
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
}


def flatten_dict(d: dict, prefix: str = "") -> dict:
    """ネストされた辞書をフラットに展開する"""
    result = {}
    for k, v in d.items():
        key = f"{prefix}.{k}" if prefix else k
        if isinstance(v, dict):
            result.update(flatten_dict(v, key))
        else:
            result[key] = v
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description="Buckler battle_stats の全フィールドを確認")
    parser.add_argument("--short-id", required=True, help="プレイヤーのshort_id")
    parser.add_argument("--cookie", default="", help="Cookie文字列")
    parser.add_argument("--cookie-file", default=_ranking.DEFAULT_COOKIE_FILE)
    parser.add_argument("--dump-json", action="store_true", help="JSON全体をそのまま出力")
    args = parser.parse_args()

    cookie = _ranking.load_cookie_text(args.cookie, args.cookie_file)

    url = f"{_ranking.BASE_URL}/profile/{args.short_id}"
    headers = _ranking.make_headers(cookie=cookie, referer=url)

    print(f"取得中: {url}")
    html = _ranking.fetch_text(url, headers, _ranking.DEFAULT_TIMEOUT)
    next_data = _ranking.extract_next_data(html)
    page_props = next_data.get("props", {}).get("pageProps", {})
    play = page_props.get("play", {})
    battle_stats = play.get("battle_stats", {})

    if not battle_stats:
        print("battle_statsが取得できませんでした。ログインが必要な可能性があります。")
        print("取得できたキー:", list(play.keys()))
        return

    if args.dump_json:
        print(json.dumps(battle_stats, ensure_ascii=False, indent=2))
        return

    # フラットに展開して全フィールドを表示
    flat = flatten_dict(battle_stats)
    all_keys = set(flat.keys())
    unmapped = all_keys - MAPPED_FIELDS

    print(f"\n{'='*60}")
    print(f"  battle_stats フィールド確認")
    print(f"  総フィールド数: {len(all_keys)}")
    print(f"  マッピング済み: {len(MAPPED_FIELDS & all_keys)}")
    print(f"  未マッピング:   {len(unmapped)}")
    print(f"{'='*60}")

    if unmapped:
        print("\n【未マッピングのフィールド】← 取得可能だが現在CSVに入っていない")
        for k in sorted(unmapped):
            print(f"  {k}: {flat[k]}")

    print("\n【マッピング済みフィールド（現在取得中）】")
    for k in sorted(MAPPED_FIELDS & all_keys):
        print(f"  {k}: {flat[k]}")

    # play以下の他のキーも確認
    other_play_keys = set(play.keys()) - {"battle_stats"}
    if other_play_keys:
        print(f"\n【play配下のbattle_stats以外のデータ】← 追加取得の可能性あり")
        for k in other_play_keys:
            v = play[k]
            preview = str(v)[:80] + "..." if len(str(v)) > 80 else str(v)
            print(f"  {k}: {preview}")


if __name__ == "__main__":
    main()

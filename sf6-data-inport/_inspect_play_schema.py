"""play JSONBの構造を把握する。ランク分布と、play配下のキー一覧を出す。"""
from __future__ import annotations

import json
import os
import sys
from collections import Counter
from pathlib import Path

from dotenv import load_dotenv

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

OUT = Path("play_schema_report.txt")
lines: list[str] = []


def log(msg: str) -> None:
    lines.append(msg)


def walk_keys(obj, prefix: str = "", depth: int = 0, out: list[str] | None = None) -> list[str]:
    if out is None:
        out = []
    if depth > 4:
        return out
    if isinstance(obj, dict):
        for k, v in obj.items():
            path = f"{prefix}.{k}" if prefix else k
            if isinstance(v, (dict, list)):
                out.append(f"{path} ({type(v).__name__})")
                walk_keys(v, path, depth + 1, out)
            else:
                out.append(f"{path} = {type(v).__name__}")
    elif isinstance(obj, list) and obj:
        walk_keys(obj[0], f"{prefix}[0]", depth + 1, out)
    return out


def main() -> None:
    from supabase import create_client

    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    client = create_client(url, key)

    # ランク分布（全件）
    r = client.table("player_data").select("rank,data_type").limit(1000).execute()
    by_rank = Counter(row["rank"] for row in r.data)
    by_type = Counter(row["data_type"] for row in r.data)
    log("=== ランク分布 ===")
    for rank, cnt in sorted(by_rank.items(), key=lambda x: -x[1]):
        log(f"  {rank}: {cnt}")
    log(f"\n=== data_type分布 ===")
    for dt, cnt in by_type.items():
        log(f"  {dt}: {cnt}")

    # play構造の確認（3件抜粋）
    r2 = (
        client.table("player_data")
        .select("player_id,rank,fetch_date,play")
        .limit(3)
        .execute()
    )
    for i, row in enumerate(r2.data):
        log(f"\n=== sample {i+1}: {row['player_id']} / {row['rank']} / {row['fetch_date']} ===")
        play = row.get("play") or {}
        keys = walk_keys(play)
        for k in keys[:80]:
            log(f"  {k}")
        if len(keys) > 80:
            log(f"  ... (total {len(keys)} keys)")

    # 生JSON 1件だけ保存（構造把握用）
    if r2.data:
        Path("play_sample_raw.json").write_text(
            json.dumps(r2.data[0], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        log("\nraw sample: play_sample_raw.json")


if __name__ == "__main__":
    try:
        main()
    finally:
        OUT.write_text("\n".join(lines), encoding="utf-8")

"""Supabase接続の生存確認（最小）。結果はsupabase_health_report.txtに書き出す。"""
from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path

from dotenv import load_dotenv

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

OUT = Path("supabase_health_report.txt")
lines: list[str] = []


def log(msg: str) -> None:
    lines.append(msg)
    print(msg)


def main() -> None:
    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    log(f"SUPABASE_URL set: {bool(url)}")
    log(f"SUPABASE_KEY set: {bool(key)}")
    if not url or not key:
        log("NG: .envが読めないか未設定")
        return

    try:
        from supabase import create_client
    except Exception as e:
        log(f"NG: supabase-py未インストール ({e})")
        return

    try:
        client = create_client(url, key)
        log("OK: client作成成功")
    except Exception as e:
        log(f"NG: client作成失敗 ({e})")
        return

    try:
        r = client.table("player_data").select("id", count="exact").limit(1).execute()
        log(f"OK: select成功 / total_count={r.count} / sample_rows={len(r.data)}")
        if r.data:
            log(f"  先頭ID: {r.data[0].get('id')}")
    except Exception as e:
        log(f"NG: select失敗 ({type(e).__name__}: {e})")
        log(traceback.format_exc())
        return

    try:
        r2 = (
            client.table("player_data")
            .select("player_id,fetch_date,data_type")
            .limit(3)
            .execute()
        )
        log(f"OK: 軽量取得成功 / 取得件数={len(r2.data)}")
        for row in r2.data:
            log(f"  {row}")
    except Exception as e:
        log(f"NG: 軽量取得失敗 ({type(e).__name__}: {e})")


if __name__ == "__main__":
    try:
        main()
    finally:
        OUT.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n結果書き出し: {OUT}")

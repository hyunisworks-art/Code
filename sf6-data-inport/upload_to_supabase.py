"""既存JSONデータをSupabaseにアップロードするスクリプト。

使い方:
    python upload_to_supabase.py
    python upload_to_supabase.py --dry-run
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from dotenv import load_dotenv
import os

load_dotenv()

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

SAMPLES_DIR = Path("data/samples")
MY_DIR = Path("data/my")


def load_json_file(path: Path) -> dict | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"  スキップ: {path.name} ({e})")
        return None


def build_row(data: dict, data_type: str) -> dict:
    """JSONデータを player_data テーブルの行に変換する。"""
    league_info = data.get("league_info") or {}
    play = data.get("play") or {}

    return {
        "player_id": str(data.get("player_id", "")),
        "fetch_date": data.get("fetch_date", ""),
        "rank": (data.get("rank") or "").lower(),
        "data_type": data_type,
        "league_info": league_info,
        "play": play,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="JSONデータをSupabaseにアップロード")
    parser.add_argument("--dry-run", action="store_true", help="件数確認のみ（実際にはアップロードしない）")
    args = parser.parse_args()

    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    if not url or not key:
        print("エラー: .env に SUPABASE_URL / SUPABASE_KEY を設定してください")
        return

    # データ収集
    rows: list[dict] = []

    # サンプルデータ
    sample_files = sorted(SAMPLES_DIR.glob("*.json")) if SAMPLES_DIR.exists() else []
    print(f"サンプルデータ: {len(sample_files)} ファイル検出")
    for path in sample_files:
        data = load_json_file(path)
        if data:
            rows.append(build_row(data, "sample"))

    # 個人データ
    my_files = sorted(MY_DIR.glob("*.json")) if MY_DIR.exists() else []
    print(f"個人データ: {len(my_files)} ファイル検出")
    for path in my_files:
        data = load_json_file(path)
        if data:
            rows.append(build_row(data, "personal"))

    print(f"\n合計: {len(rows)} 件")

    if args.dry_run:
        print("(dry-run: アップロードはスキップ)")
        return

    # Supabase接続・アップロード
    from supabase import create_client
    client = create_client(url, key)

    # バッチアップロード（50件ずつ・JSONBが大きいため小バッチ）
    batch_size = 50
    uploaded = 0
    for i in range(0, len(rows), batch_size):
        batch = rows[i:i + batch_size]
        try:
            result = client.table("player_data").upsert(
                batch,
                on_conflict="player_id,fetch_date,data_type",
            ).execute()
            uploaded += len(result.data)
            print(f"  アップロード: {uploaded}/{len(rows)} 件完了")
        except Exception as e:
            print(f"  エラー (batch {i}): {e}")

    # 確認
    count = client.table("player_data").select("id", count="exact").execute()
    print(f"\nSupabase player_data テーブル: {count.count} 件")
    print("完了!")


if __name__ == "__main__":
    main()

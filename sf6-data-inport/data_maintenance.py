"""SF6 サンプルデータ自動メンテナンスモジュール

実行フロー:
  1. 期限切れファイル削除  （STALE_DAYS 日以上前のファイルを全削除）
  2. 過剰サンプル削除      （TARGET を超えるバンドから古い順に削除）
  3. 不足サンプル収集      （MIN を下回るバンドに対して収集）

LP バンド（gold / platinum / diamond）と MR バンド（master）を個別に管理する。
MASTER は LP ではなく MR サンプルとして管理する（同じ rank_key="master"）。

【LP/MR 分布バランスについて】
現状のサンプル JSON には LP/MR 値が保存されていないため、削除は日付ベース（古い順）で行う。
収集時は collect_samples.py が各サブランク（例: Diamond1〜Diamond5）を順に収集するので、
1 バンドあたりの収集件数を sub_rank 数で割り、サブランクごとに均等に取るよう拡張している。

使い方:
    # スタンドアロン実行
    python data_maintenance.py
    python data_maintenance.py --dry-run
    python data_maintenance.py --no-collect  # 削除のみ（収集しない）

    # モジュールとして使う（dashboard.py など）
    from data_maintenance import run_maintenance, MaintenanceResult
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

import collect_samples as cs
import scrape_rankings as _ranking

# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------

STALE_DAYS = 30          # この日数より古いファイルは全バンド問わず削除

# LP バンド（MASTER 未満）の目標・最低件数
LP_BAND_CONFIG: dict[str, dict[str, int]] = {
    "gold":     {"min": 60, "target": 80},
    "platinum": {"min": 60, "target": 80},
    "diamond":  {"min": 60, "target": 80},
}

# MR バンド（MASTER 以上）の目標・最低件数
MR_BAND_CONFIG: dict[str, dict[str, int]] = {
    "master": {"min": 30, "target": 60},
}

MAX_COLLECT_PER_BAND = 50   # 1 バンドあたり 1 回の収集上限（レート制限配慮）
DEFAULT_DELAY   = cs.DEFAULT_DELAY
DEFAULT_TIMEOUT = _ranking.DEFAULT_TIMEOUT

SAMPLES_DIR = cs.SAMPLES_DIR


# ---------------------------------------------------------------------------
# データクラス
# ---------------------------------------------------------------------------

@dataclass
class BandStatus:
    rank_key: str
    count: int
    target: int
    min_count: int
    excess: int      # 0 以下なら過剰なし
    shortage: int    # 0 以下なら不足なし
    deleted: int = 0
    collected: int = 0
    errors: list[str] = field(default_factory=list)

    @property
    def is_excess(self) -> bool:
        return self.excess > 0

    @property
    def is_shortage(self) -> bool:
        return self.shortage > 0


@dataclass
class MaintenanceResult:
    stale_deleted: int = 0
    bands: list[BandStatus] = field(default_factory=list)
    total_deleted: int = 0
    total_collected: int = 0
    errors: list[str] = field(default_factory=list)

    def summary_lines(self) -> list[str]:
        lines = [f"期限切れ削除: {self.stale_deleted} 件"]
        for b in self.bands:
            parts = []
            if b.deleted:
                parts.append(f"削除 {b.deleted}")
            if b.collected:
                parts.append(f"収集 {b.collected}")
            after = b.count - b.deleted + b.collected
            status = f"{b.rank_key}: {b.count}件 → {after}件"
            if parts:
                status += f"（{', '.join(parts)}）"
            lines.append(status)
        lines.append(f"合計: 削除 {self.total_deleted} 件 / 収集 {self.total_collected} 件")
        return lines


# ---------------------------------------------------------------------------
# ファイル操作
# ---------------------------------------------------------------------------

def _list_samples_by_rank(samples_dir: Path) -> dict[str, list[Path]]:
    """ランク別にサンプルファイルを返す。ファイル名の日付順（古い順）でソート済み。"""
    by_rank: dict[str, list[Path]] = {}
    if not samples_dir.exists():
        return by_rank

    for json_file in sorted(samples_dir.glob("????-??-??_*.json")):
        try:
            data = json.loads(json_file.read_text(encoding="utf-8"))
            rank = data.get("rank", "").strip().lower()
            if rank:
                by_rank.setdefault(rank, []).append(json_file)
        except (json.JSONDecodeError, OSError):
            continue

    return by_rank


def _delete_stale(samples_dir: Path, stale_days: int, dry_run: bool) -> int:
    """期限切れファイルを削除する。削除件数を返す。"""
    if not samples_dir.exists():
        return 0
    cutoff = datetime.now() - timedelta(days=stale_days)
    deleted = 0
    for json_file in samples_dir.glob("????-??-??_*.json"):
        try:
            date_str = json_file.stem[:10]
            file_date = datetime.strptime(date_str, "%Y-%m-%d")
            if file_date < cutoff:
                if not dry_run:
                    json_file.unlink()
                print(f"  [stale] 削除: {json_file.name}")
                deleted += 1
        except (ValueError, OSError):
            continue
    return deleted


def _trim_excess(
    files: list[Path],
    current: int,
    target: int,
    dry_run: bool,
) -> int:
    """過剰なサンプルを古い順に削除する。削除件数を返す。

    ファイルリストはすでに古い順（昇順）でソートされている前提。
    """
    need_delete = current - target
    if need_delete <= 0:
        return 0

    deleted = 0
    for f in files:
        if deleted >= need_delete:
            break
        try:
            if not dry_run:
                f.unlink()
            print(f"  [trim] 削除: {f.name}")
            deleted += 1
        except OSError as exc:
            print(f"  [trim] 削除失敗 {f.name}: {exc}")
    return deleted


# ---------------------------------------------------------------------------
# 収集（LP/MR 分布を考慮した均等収集）
# ---------------------------------------------------------------------------

def _collect_balanced(
    rank_key: str,
    total_needed: int,
    cookie: str,
    timeout: int,
    delay: float,
    request_counter: list[int],
    dry_run: bool,
) -> int:
    """サブランク（sub_rank）ごとに均等に収集して分布バランスを保つ。

    collect_samples.py の RANK_CONFIG にある league_rank_min/max を使い、
    各サブランクに均等な件数を割り当てて収集する。
    """
    if rank_key not in cs.RANK_CONFIG:
        print(f"  [collect] 未知のランクキー: {rank_key}")
        return 0

    config = cs.RANK_CONFIG[rank_key]
    sub_min: int = config["league_rank_min"]
    sub_max: int = config["league_rank_max"]
    sub_count = sub_max - sub_min + 1

    # サブランクごとの収集件数（均等割り。端数は最後のサブランクに）
    per_sub = max(1, total_needed // sub_count)
    remainder = total_needed - per_sub * sub_count

    total_saved = 0

    for i, sub_rank in enumerate(range(sub_min, sub_max + 1)):
        if request_counter[0] >= cs.MAX_REQUESTS_PER_SESSION:
            print(f"  [collect] セッション上限に達したため停止")
            break

        count_for_sub = per_sub + (1 if i == sub_count - 1 else 0)
        if remainder > 0 and i < remainder:
            count_for_sub += 1

        if count_for_sub <= 0:
            continue

        print(f"  [collect] {rank_key} サブランク{sub_rank}: {count_for_sub}件収集")

        # collect_samples.py の内部関数を直接使う（サブランク指定は league_rank に相当）
        # collect_samples_for_rank はサブランク指定ができないため、短縮版を実行
        saved = _collect_for_subrank(
            rank_key=rank_key,
            sub_rank=sub_rank,
            count=count_for_sub,
            cookie=cookie,
            timeout=timeout,
            delay=delay,
            request_counter=request_counter,
            dry_run=dry_run,
        )
        total_saved += saved

    return total_saved


def _collect_for_subrank(
    rank_key: str,
    sub_rank: int,
    count: int,
    cookie: str,
    timeout: int,
    delay: float,
    request_counter: list[int],
    dry_run: bool,
) -> int:
    """指定サブランクのみを対象に収集する。"""
    from datetime import date as _date
    import time as _time

    today = _date.today().strftime("%Y-%m-%d")
    SAMPLES_DIR.mkdir(parents=True, exist_ok=True)

    # 既存 ID をロード（重複防止）
    existing_ids: set[str] = set()
    for f in SAMPLES_DIR.glob("*.json"):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            pid = data.get("player_id", "")
            if pid:
                existing_ids.add(pid)
        except (json.JSONDecodeError, OSError):
            continue

    # ランキングから short_id を取得
    ranking_type = cs.RANK_CONFIG[rank_key]["ranking_type"]
    first_url = _ranking.build_ranking_page_url(ranking_type, 1, _ranking.DEFAULT_LOCALE)
    headers = _ranking.make_headers(cookie=cookie, referer=first_url)
    request_counter[0] += 1

    try:
        html = _ranking.fetch_text(first_url, headers, timeout)
    except Exception as exc:
        print(f"    ランキング取得失敗: {exc}")
        return 0

    build_id = _ranking.get_build_id(html)
    short_ids: list[str] = []

    for page in range(1, 5):
        if len(short_ids) >= count * 2:
            break
        if request_counter[0] >= cs.MAX_REQUESTS_PER_SESSION:
            break

        page_url = _ranking.build_ranking_page_url(
            ranking_type, page, _ranking.DEFAULT_LOCALE, league_rank=sub_rank
        )
        api_url = _ranking.build_next_data_url(
            build_id, ranking_type, page, _ranking.DEFAULT_LOCALE, league_rank=sub_rank
        )
        api_headers = _ranking.make_headers(cookie=cookie, referer=page_url)
        request_counter[0] += 1

        try:
            data = _ranking.fetch_json(api_url, api_headers, timeout)
        except Exception:
            break

        page_props = data.get("pageProps", {})
        try:
            payload = _ranking.get_ranking_payload(page_props, ranking_type)
        except Exception:
            break

        for item in payload.get("ranking_fighter_list", []):
            if not isinstance(item, dict):
                continue
            flat = _ranking.flatten_item(item)
            sid = flat.get(cs._COL_SHORT_ID, "").strip()
            if sid and sid not in existing_ids:
                short_ids.append(sid)

        _time.sleep(delay)

    # play データ取得 & 保存
    saved = 0
    for sid in short_ids[:count]:
        if request_counter[0] >= cs.MAX_REQUESTS_PER_SESSION:
            break
        if dry_run:
            print(f"    [dry-run] {sid}")
            saved += 1
            continue

        request_counter[0] += 1
        try:
            url = f"{_ranking.BASE_URL}/profile/{sid}"
            hdrs = _ranking.make_headers(cookie=cookie, referer=url)
            html = _ranking.fetch_text(url, hdrs, timeout)
            next_data = _ranking.extract_next_data(html)
            page_props = next_data.get("props", {}).get("pageProps", {})
            play_data = page_props.get("play", {})

            payload_out: dict[str, Any] = {
                "fetch_date": today,
                "player_id": sid,
                "rank": rank_key,
                "sub_rank": sub_rank,
                "league_info": {},
                "play": play_data,
            }
            json_path = SAMPLES_DIR / f"{today}_{sid}.json"
            json_path.write_text(
                json.dumps(payload_out, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            existing_ids.add(sid)
            saved += 1
            print(f"    保存: {json_path.name}")
        except Exception as exc:
            print(f"    スキップ ({sid}): {exc}")

        _time.sleep(delay)

    return saved


# ---------------------------------------------------------------------------
# メンテナンス実行
# ---------------------------------------------------------------------------

def run_maintenance(
    cookie: str = "",
    timeout: int = DEFAULT_TIMEOUT,
    delay: float = DEFAULT_DELAY,
    dry_run: bool = False,
    no_collect: bool = False,
) -> MaintenanceResult:
    """メンテナンスのメインエントリポイント。

    Args:
        cookie:     Buckler Cookie（収集が必要な場合に必須）
        timeout:    HTTP タイムアウト秒
        delay:      リクエスト間隔秒
        dry_run:    True なら削除・収集を実行しない（確認のみ）
        no_collect: True なら削除のみ実行（収集しない）

    Returns:
        MaintenanceResult
    """
    result = MaintenanceResult()
    request_counter = [0]

    print("=== SF6 データメンテナンス開始 ===")

    # Step 1: 期限切れデータ削除
    print(f"\n[Step 1] 期限切れデータ削除（{STALE_DAYS}日以上前）")
    result.stale_deleted = _delete_stale(SAMPLES_DIR, STALE_DAYS, dry_run)
    if result.stale_deleted == 0:
        print("  期限切れファイルなし")
    else:
        print(f"  削除: {result.stale_deleted} 件")

    # Step 2: 現在のバンド別件数を集計
    print("\n[Step 2] バンド別件数確認")
    by_rank = _list_samples_by_rank(SAMPLES_DIR)

    all_band_configs = {**LP_BAND_CONFIG, **MR_BAND_CONFIG}
    band_statuses: list[BandStatus] = []

    for rank_key, cfg in all_band_configs.items():
        files = by_rank.get(rank_key, [])
        count = len(files)
        target = cfg["target"]
        min_count = cfg["min"]
        excess = count - target
        shortage = min_count - count

        status = BandStatus(
            rank_key=rank_key,
            count=count,
            target=target,
            min_count=min_count,
            excess=max(0, excess),
            shortage=max(0, shortage),
        )
        band_statuses.append(status)
        flag = "⚠ 過剰" if status.is_excess else ("⚠ 不足" if status.is_shortage else "OK")
        print(f"  {rank_key:<10}: {count:>4}件  目標={target}  最低={min_count}  [{flag}]")

    # Step 3: 過剰削除
    print("\n[Step 3] 過剰サンプル削除")
    for status in band_statuses:
        if not status.is_excess:
            continue
        files = by_rank.get(status.rank_key, [])
        deleted = _trim_excess(files, status.count, status.target, dry_run)
        status.deleted = deleted
        result.total_deleted += deleted
        print(f"  {status.rank_key}: {deleted} 件削除")

    no_excess = all(not s.is_excess for s in band_statuses)
    if no_excess:
        print("  過剰バンドなし")

    # Step 4: 不足収集
    print("\n[Step 4] 不足サンプル収集")
    if no_collect:
        print("  --no-collect 指定のためスキップ")
    elif not cookie:
        print("  Cookie 未設定のためスキップ（Cookie を設定すると自動収集が有効になります）")
    else:
        shortage_bands = [s for s in band_statuses if s.is_shortage]
        if not shortage_bands:
            print("  不足バンドなし")
        else:
            for status in shortage_bands:
                if request_counter[0] >= cs.MAX_REQUESTS_PER_SESSION:
                    print("  セッション上限に達したため収集停止")
                    break
                needed = min(status.shortage, MAX_COLLECT_PER_BAND)
                print(f"  {status.rank_key}: {needed} 件収集開始")
                collected = _collect_balanced(
                    rank_key=status.rank_key,
                    total_needed=needed,
                    cookie=cookie,
                    timeout=timeout,
                    delay=delay,
                    request_counter=request_counter,
                    dry_run=dry_run,
                )
                status.collected = collected
                result.total_collected += collected

    result.bands = band_statuses
    print("\n=== メンテナンス完了 ===")
    for line in result.summary_lines():
        print(f"  {line}")

    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="SF6 サンプルデータ自動メンテナンス")
    parser.add_argument("--dry-run", action="store_true", help="削除・収集を行わず状況確認のみ")
    parser.add_argument("--no-collect", action="store_true", help="削除のみ実行（収集しない）")
    parser.add_argument("--cookie", default="", help="Buckler Cookie 文字列")
    parser.add_argument("--cookie-file", default=cs.DEFAULT_COOKIE_FILE)
    parser.add_argument("--delay", type=float, default=DEFAULT_DELAY)
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT)
    args = parser.parse_args()

    cookie = _ranking.load_cookie_text(args.cookie, args.cookie_file)

    run_maintenance(
        cookie=cookie,
        timeout=args.timeout,
        delay=args.delay,
        dry_run=args.dry_run,
        no_collect=args.no_collect,
    )


if __name__ == "__main__":
    main()

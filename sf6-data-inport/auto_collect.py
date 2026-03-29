"""SF6 サンプル自動補充スクリプト (Phase 0 タスク3)

data/samples/ にある既存JSONを集計し、目標件数（60件）に満たない
ランク帯を自動検出して collect_samples.py の収集ロジックで補充する。

使い方:
    python auto_collect.py              # 全ランク帯をチェックして不足分を補充
    python auto_collect.py --dry-run    # 収集は行わず、不足状況だけ表示
    python auto_collect.py --rank platinum  # 特定ランク帯のみ対象

注意:
    - collect_samples.py を import して使う（別プロセス起動はしない）
    - analyze_step1.py / analyze_playlog.py / sf6-playlog-out.csv は変更しない
    - collect_samples.py 自体は変更しない
    - 1ランク帯あたりのリクエスト間隔は DEFAULT_DELAY（3秒）以上
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

# Windows環境でUTF-8出力を強制
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

import collect_samples as cs
import scrape_rankings as _ranking

# ---------------------------------------------------------------------------
# 定数
# ---------------------------------------------------------------------------

TARGET_COUNT = 60       # 目標件数（これを下回るランク帯を収集対象にする）
MAX_COLLECT_PER_RUN = 50  # 1回の実行で各ランク帯最大50件まで（レート制限配慮）

# ランク表示名（表示用）
RANK_DISPLAY = {
    "bronze":   "Bronze",
    "silver":   "Silver",
    "gold":     "Gold",
    "platinum": "Platinum",
    "diamond":  "Diamond",
    "master":   "Master",
}


# ---------------------------------------------------------------------------
# 集計
# ---------------------------------------------------------------------------

def count_samples_by_rank(samples_dir: Path) -> Counter:
    """data/samples/ のJSONを読み込みランク帯別の件数を返す。"""
    counter: Counter = Counter()
    if not samples_dir.exists():
        return counter

    for json_file in samples_dir.glob("*.json"):
        try:
            data = json.loads(json_file.read_text(encoding="utf-8"))
            rank = data.get("rank", "").strip().lower()
            if rank in cs.ALL_RANKS:
                counter[rank] += 1
        except (json.JSONDecodeError, OSError):
            continue

    return counter


def detect_shortage(counter: Counter, target_ranks: list[str]) -> list[tuple[str, int, int]]:
    """不足ランク帯を検出する。

    Returns:
        [(rank_key, current_count, shortage)] のリスト（shortage > 0 のみ）
    """
    result = []
    for rank_key in target_ranks:
        current = counter.get(rank_key, 0)
        shortage = TARGET_COUNT - current
        if shortage > 0:
            result.append((rank_key, current, shortage))
    return result


# ---------------------------------------------------------------------------
# 表示
# ---------------------------------------------------------------------------

def print_status_table(counter: Counter, target_ranks: list[str]) -> None:
    """ランク帯別の状況テーブルを表示する。"""
    print(f"{'ランク帯':<12} {'現在件数':>8}  {'目標':>6}  {'不足数':>8}")
    print("-" * 44)

    for rank_key in target_ranks:
        current = counter.get(rank_key, 0)
        shortage = TARGET_COUNT - current
        label = RANK_DISPLAY.get(rank_key, rank_key)

        if shortage <= 0:
            status = "達成 OK"
            print(f"{label:<12} {current:>6}件   {TARGET_COUNT:>4}件   {status:>10}")
        else:
            status = f"-{shortage}件  <- 収集対象"
            print(f"{label:<12} {current:>6}件   {TARGET_COUNT:>4}件   {status}")


def print_collection_plan(shortage_list: list[tuple[str, int, int]]) -> None:
    """収集計画のサマリーを表示する。"""
    if not shortage_list:
        print("\n全ランク帯が目標件数（60件）を達成しています。収集は不要です。")
        return

    items = []
    total = 0
    for rank_key, _current, shortage in shortage_list:
        collect_count = min(shortage, MAX_COLLECT_PER_RUN)
        label = RANK_DISPLAY.get(rank_key, rank_key)
        items.append(f"{label}({collect_count}件)")
        total += collect_count

    print(f"\n収集対象: {', '.join(items)}")
    print(f"合計収集予定: {total}件")
    if any(shortage > MAX_COLLECT_PER_RUN for _, _, shortage in shortage_list):
        print(f"※ 1ランク帯あたり最大{MAX_COLLECT_PER_RUN}件のため、1回では全不足を補えないランク帯があります")


# ---------------------------------------------------------------------------
# 収集実行
# ---------------------------------------------------------------------------

def run_collection(
    shortage_list: list[tuple[str, int, int]],
    cookie: str,
    timeout: int,
    delay: float,
    request_counter: list[int],
) -> dict[str, tuple[int, int]]:
    """不足ランク帯の収集を実行する。

    Args:
        request_counter: セッション全体のリクエストカウンター（共有参照）

    Returns:
        {rank_key: (saved_count, skipped_count)} の辞書
    """
    results: dict[str, tuple[int, int]] = {}

    for rank_key, _current, shortage in shortage_list:
        if request_counter[0] >= cs.MAX_REQUESTS_PER_SESSION:
            print(f"\n※ セッション上限({cs.MAX_REQUESTS_PER_SESSION}リクエスト)に達したため残りをスキップします")
            break

        collect_count = min(shortage, MAX_COLLECT_PER_RUN)

        try:
            saved, skipped = cs.collect_samples_for_rank(
                rank_key=rank_key,
                count=collect_count,
                cookie=cookie,
                timeout=timeout,
                delay=delay,
                dry_run=False,
                request_counter=request_counter,
            )
            results[rank_key] = (saved, skipped)

        except PermissionError as exc:
            print(f"\n権限エラー: {exc}")
            print("Buckler のランキングデータ取得には Cookie（ログイン状態）が必要です。")
            print("ブラウザで Buckler にログイン後、Cookie を .buckler_cookie.txt に保存してください。")
            break
        except ConnectionError as exc:
            print(f"\n接続エラー: {exc}")
            break
        except KeyboardInterrupt:
            print("\n\n中断されました（Ctrl+C）")
            break

    return results


# ---------------------------------------------------------------------------
# サマリー表示
# ---------------------------------------------------------------------------

def print_summary(
    before: Counter,
    after: Counter,
    target_ranks: list[str],
    results: dict[str, tuple[int, int]],
    request_counter: list[int],
) -> None:
    """完了後のサマリーを表示する。"""
    print("\n=== 収集完了サマリー ===")
    print(f"{'ランク帯':<12} {'収集前':>8}  {'収集後':>8}  {'保存':>6}  {'スキップ':>8}")
    print("-" * 52)

    total_saved = 0
    total_skipped = 0

    for rank_key in target_ranks:
        label = RANK_DISPLAY.get(rank_key, rank_key)
        before_count = before.get(rank_key, 0)
        after_count = after.get(rank_key, 0)
        saved, skipped = results.get(rank_key, (0, 0))
        total_saved += saved
        total_skipped += skipped
        print(f"{label:<12} {before_count:>6}件   {after_count:>6}件  {saved:>4}件  {skipped:>6}件")

    print("-" * 52)
    print(f"{'合計':<12} {'':>8}  {'':>8}  {total_saved:>4}件  {total_skipped:>6}件")
    print(f"\nリクエスト総数: {request_counter[0]}件")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="SF6 サンプル不足チェック＆自動補充スクリプト"
    )
    parser.add_argument(
        "--rank",
        choices=cs.ALL_RANKS,
        default=None,
        help=(
            "対象ランク帯（bronze/silver/gold/platinum/diamond/master）。"
            "省略時は全ランク帯をチェック"
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="収集は行わず、不足状況だけ表示する",
    )
    parser.add_argument(
        "--cookie",
        default="",
        help="Cookie文字列（任意。省略時は .buckler_cookie.txt を使用）",
    )
    parser.add_argument(
        "--cookie-file",
        default=cs.DEFAULT_COOKIE_FILE,
        help=f"Cookie文字列を保存したファイル（デフォルト: {cs.DEFAULT_COOKIE_FILE}）",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=cs.DEFAULT_DELAY,
        help=f"リクエスト間の待機秒数（デフォルト: {cs.DEFAULT_DELAY}秒）",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=_ranking.DEFAULT_TIMEOUT,
        help="HTTPタイムアウト秒数",
    )
    args = parser.parse_args()

    if args.delay < cs.DEFAULT_DELAY:
        print(f"警告: --delay が {cs.DEFAULT_DELAY}秒未満です。Bucklerへの負荷に注意してください")

    # 対象ランク帯
    target_ranks = [args.rank] if args.rank else cs.ALL_RANKS

    # 現在の件数を集計
    print("=== サンプル不足チェック ===")
    counter = count_samples_by_rank(cs.SAMPLES_DIR)

    print_status_table(counter, target_ranks)

    # 不足ランク帯を検出
    shortage_list = detect_shortage(counter, target_ranks)
    print_collection_plan(shortage_list)

    if args.dry_run or not shortage_list:
        return

    # Cookie 読み込み
    cookie = _ranking.load_cookie_text(args.cookie, args.cookie_file)
    try:
        _ranking.validate_cookie_text(cookie)
    except ValueError as exc:
        print(f"\nエラー: {exc}")
        print("収集を行うには Cookie が必要です。--dry-run で不足状況のみ確認できます。")
        sys.exit(1)

    # 収集実行
    print("\n=== 収集開始 ===")
    request_counter = [0]

    results = run_collection(
        shortage_list=shortage_list,
        cookie=cookie,
        timeout=args.timeout,
        delay=args.delay,
        request_counter=request_counter,
    )

    # 収集後の件数を再集計
    after_counter = count_samples_by_rank(cs.SAMPLES_DIR)

    total_saved = sum(s for s, _ in results.values())
    total_skipped = sum(sk for _, sk in results.values())

    print_summary(counter, after_counter, target_ranks, results, request_counter)
    print(f"\n保存合計: {total_saved}件 / スキップ合計: {total_skipped}件")


if __name__ == "__main__":
    main()

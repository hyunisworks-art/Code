"""SF6 Step1分析: 同ランク帯との比較スクリプト

使い方:
    python analyze_step1.py --player "プレイヤー名"
    python analyze_step1.py --player "プレイヤー名" --rank PLATINUM
    python analyze_step1.py --player "プレイヤー名" --input sf6-playlog-out.csv
"""
from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path
from typing import Any

# Windows環境でUTF-8出力を強制
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

from analyze_playlog import load_playlog_rows, get_feature_names, parse_numeric

# 評価方向: Trueなら「低い方が良い」指標
LOWER_IS_BETTER = {
    "相手に追い詰められている時間",
    "スタンさせられた回数",
    "ドライブインパクト_受けた回数",
    "投げ_受けた回数",
    "パニッシュカウンターを受けた回数",
    "相手のドライブインパクトに決めた回数",  # 相手のDIが通った回数
    "自分のドライブパリィを投げられた",
}

RANK_ORDER = [
    "ROOKIE", "IRON", "BRONZE", "SILVER", "GOLD",
    "PLATINUM", "DIAMOND", "MASTER",
]


def get_rank_group(rank: str) -> str:
    for r in RANK_ORDER:
        if rank.upper().startswith(r):
            return r
    return rank.upper()


def filter_rows_by_rank(rows: list[dict[str, Any]], rank_group: str) -> list[dict[str, Any]]:
    return [row for row in rows if get_rank_group(str(row.get("ランク", ""))) == rank_group]


def find_player_rows(rows: list[dict[str, Any]], player_name: str) -> list[dict[str, Any]]:
    return [row for row in rows if player_name in str(row.get("プレイヤー名", ""))]


def compute_mean(values: list[float]) -> float | None:
    if not values:
        return None
    return sum(values) / len(values)


def compute_std(values: list[float]) -> float:
    if len(values) < 2:
        return 0.0
    avg = sum(values) / len(values)
    variance = sum((v - avg) ** 2 for v in values) / (len(values) - 1)
    return math.sqrt(variance)


def welch_t_test(val: float, group: list[float]) -> tuple[float, str]:
    """1サンプル vs グループ平均のt検定（Welch's t-test近似）。
    戻り値: (t値, 有意性マーク)
    """
    n = len(group)
    if n < 3:
        return 0.0, "   "
    mean_g = sum(group) / n
    std_g = compute_std(group)
    if std_g == 0:
        return 0.0, "   "
    t = (val - mean_g) / (std_g / math.sqrt(n))
    # 自由度n-1の両側t検定の簡易判定
    # df=99でt>1.984→p<0.05、t>2.626→p<0.01
    abs_t = abs(t)
    if abs_t > 2.626:
        return t, "**"   # p<0.01
    elif abs_t > 1.984:
        return t, "*"    # p<0.05
    else:
        return t, "   "  # 有意差なし


def get_numeric_values(rows: list[dict[str, Any]], feature: str) -> list[float]:
    result = []
    for row in rows:
        val = parse_numeric(str(row.get(feature, "")))
        if val is not None:
            result.append(val)
    return result


def is_percent_feature(row: dict[str, Any], feature: str) -> bool:
    return "%" in str(row.get(feature, ""))


def print_comparison(
    player_name: str,
    player_row: dict[str, Any],
    rank_group: str,
    rank_rows: list[dict[str, Any]],
    feature_names: list[str],
) -> None:
    print(f"\n{'='*72}")
    print(f"  Step1分析: {player_name} vs {rank_group}帯 ({len(rank_rows)}名)")
    print("  * p<0.05（有意差あり）  ** p<0.01（強い有意差）")
    print(f"{'='*72}")
    print(f"  {'指標':<30} {'あなた':>9} {'同帯平均':>9} {'差分':>9} {'sig':>4}  {'評価'}")
    print(f"  {'-'*68}")

    for feature in feature_names:
        player_val = parse_numeric(str(player_row.get(feature, "")))
        rank_vals = get_numeric_values(rank_rows, feature)
        rank_mean = compute_mean(rank_vals)

        if player_val is None or rank_mean is None or len(rank_vals) < 3:
            continue

        diff = player_val - rank_mean
        _, sig = welch_t_test(player_val, rank_vals)

        lower_better = feature in LOWER_IS_BETTER
        if lower_better:
            mark = "[良]  " if diff < -0.01 else ("[--]  " if abs(diff) < 0.01 else "[課題]")
        else:
            mark = "[良]  " if diff > 0.01 else ("[--]  " if abs(diff) < 0.01 else "[課題]")

        pct = is_percent_feature(player_row, feature)
        if pct:
            print(f"  {feature:<30} {player_val:>8.1f}% {rank_mean:>8.1f}% {diff:>+8.1f}%  {sig:>2}  {mark}")
        else:
            print(f"  {feature:<30} {player_val:>9.2f} {rank_mean:>9.2f} {diff:>+9.2f}  {sig:>2}  {mark}")

    print(f"{'='*72}")
    print(f"  {player_name}  ランク: {player_row.get('ランク','')}  LP: {player_row.get('リーグポイント','')}")
    print(f"  サンプル数: {len(rank_rows)}名  （統計的有意性: n>=30で信頼性向上、n>=100で十分）")
    print(f"{'='*72}\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="同ランク帯との指標比較（Step1分析）")
    parser.add_argument("--player", required=True, help="比較対象のプレイヤー名")
    parser.add_argument("--input", default="sf6-playlog-out.csv", help="分析対象CSV")
    parser.add_argument("--rank", default=None, help="比較ランク帯（省略時: プレイヤーのランクを使用）")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"エラー: ファイルが見つかりません: {input_path}")
        return

    columns, rows = load_playlog_rows(input_path)
    feature_names = get_feature_names(columns)
    print(f"データ読み込み完了: {len(rows)}件 / 分析項目: {len(feature_names)}項目")

    player_rows = find_player_rows(rows, args.player)
    if not player_rows:
        print(f"エラー: プレイヤー '{args.player}' が見つかりません")
        print("データ内のプレイヤー名（先頭10件）:")
        for row in rows[:10]:
            print(f"  - {row.get('プレイヤー名', '')}")
        return

    player_row = player_rows[-1]
    player_rank = str(player_row.get("ランク", ""))
    rank_group = args.rank.upper() if args.rank else get_rank_group(player_rank)

    print(f"プレイヤー: {args.player}  ランク: {player_rank}  比較帯: {rank_group}")

    rank_rows = [
        row for row in filter_rows_by_rank(rows, rank_group)
        if args.player not in str(row.get("プレイヤー名", ""))
    ]

    if len(rank_rows) < 30:
        print(f"警告: {rank_group}帯のデータが少ないです（{len(rank_rows)}件）。精度向上には30件以上推奨。")

    print_comparison(args.player, player_row, rank_group, rank_rows, feature_names)


if __name__ == "__main__":
    main()

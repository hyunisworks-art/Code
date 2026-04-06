"""SF6 Step2分析: 1つ上のランク帯との比較スクリプト

【設計方針】 coaching-design.md / spec.md より
  Step 2 = 「次のランクに上がるために何が足りないか？」
  比較対象: プレイヤーの現在ランクより1つ上のランク帯のサンプル
  出力: 差分が大きく統計的根拠がある項目 = 練習課題候補

【Step 1 との違い】
  Step 1: 同ランク帯内での自分の位置把握（今の自分は帯内でどこ？）
  Step 2: 次のランク帯との差分把握（次のランクに上がるには何が要る？）
  → この2つを混在させない。必ず別セクションで提示する。

【目標ランク】
  ひゅーさん目標ランク = MASTER（2026-04-06確定）
  省略時は next_rank() で自動判定（DIAMOND → MASTER）
  --target-rank 引数で任意ランクを指定可能

使い方（JSONモード）:
    python analyze_step2.py --player 2202760091
    python analyze_step2.py --player 2202760091 --target-rank MASTER

使い方（CSVモード・後方互換）:
    python analyze_step2.py --player "プレイヤー名" --input sf6-playlog-out.csv
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

from analyze_step1 import (
    RANK_ORDER,
    JSON_FEATURE_NAMES,
    LOWER_IS_BETTER,
    load_json_rows,
    filter_rows_by_rank,
    find_player_rows,
    get_rank_group,
    compute_mean,
    compute_std,
    welch_t_test,
    get_numeric_values,
    is_percent_feature,
)

from analyze_playlog import load_playlog_rows, get_feature_names, parse_numeric


def next_rank(current_rank: str) -> str | None:
    """現在ランクの1つ上のランク名を返す。MASTERの場合はNone。"""
    rank = get_rank_group(current_rank)
    try:
        idx = RANK_ORDER.index(rank)
    except ValueError:
        return None
    if idx + 1 >= len(RANK_ORDER):
        return None
    return RANK_ORDER[idx + 1]


def print_step2_comparison(
    player_name: str,
    player_row: dict,
    current_rank: str,
    target_rank: str,
    target_rows: list[dict],
    feature_names: list[str],
) -> None:
    """Step 2分析結果を出力する。"""
    print(f"\n{'='*72}")
    print(f"  Step2分析: {player_name} （{current_rank}帯）→ {target_rank}帯への課題")
    print(f"  比較対象: {target_rank}帯 {len(target_rows)}名")
    print("  * p<0.05（有意差あり）  ** p<0.01（強い有意差）")
    print(f"{'='*72}")
    print(f"  {'指標':<30} {'あなた':>9} {target_rank+'帯平均':>12} {'差分':>9} {'sig':>4}  {'評価'}")
    print(f"  {'-'*70}")

    shortage_list = []

    for feature in feature_names:
        player_val = parse_numeric(str(player_row.get(feature, "")))
        target_vals = get_numeric_values(target_rows, feature)
        target_mean = compute_mean(target_vals)

        if player_val is None or target_mean is None or len(target_vals) < 3:
            continue

        diff = player_val - target_mean
        t, sig = welch_t_test(player_val, target_vals)

        lower_better = feature in LOWER_IS_BETTER
        if lower_better:
            mark = "[良]  " if diff < -0.01 else ("[--]  " if abs(diff) < 0.01 else "[課題]")
            is_shortage = diff > 0.01
        else:
            mark = "[良]  " if diff > 0.01 else ("[--]  " if abs(diff) < 0.01 else "[課題]")
            is_shortage = diff < -0.01

        pct = is_percent_feature(player_row, feature)
        if pct:
            print(f"  {feature:<30} {player_val:>8.1f}% {target_mean:>10.1f}%  {diff:>+8.1f}%  {sig:>2}  {mark}")
        else:
            print(f"  {feature:<30} {player_val:>9.2f} {target_mean:>12.2f} {diff:>+9.2f}  {sig:>2}  {mark}")

        if is_shortage and sig.strip():
            shortage_list.append((feature, abs(diff), sig))

    print(f"{'='*72}")

    # 優先課題サマリー（有意差ありの課題のみ）
    if shortage_list:
        shortage_list.sort(key=lambda x: x[1], reverse=True)
        print(f"\n  【Step 2 優先課題（有意差あり）上位5件】")
        for i, (feat, diff_abs, sig) in enumerate(shortage_list[:5], 1):
            print(f"  {i}. {feat}  （差分: {diff_abs:.3f}, {sig.strip()}）")
    else:
        print(f"\n  有意差のある課題が検出されませんでした。サンプル数を増やすか、上位ランク帯との比較を検討してください。")

    print(f"\n  【注意】これは統計的傾向の提示です。相関≠因果。課題の優先順位は練習の文脈で判断してください。")
    print(f"{'='*72}\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Step2分析: 1つ上のランク帯との差分比較")
    parser.add_argument("--player", required=True, help="short_id（JSONモード）またはプレイヤー名（CSVモード）")
    parser.add_argument("--input", default=None, help="CSVファイル（省略時はJSONモード）")
    parser.add_argument("--target-rank", default=None, help="比較対象ランク（省略時: 現在ランクの1つ上）")
    args = parser.parse_args()

    _SAMPLES_DIR = Path("data/samples")
    use_json = args.input is None and _SAMPLES_DIR.exists() and any(_SAMPLES_DIR.glob("*.json"))

    if use_json:
        feature_names, sample_rows, player_row = load_json_rows(args.player)
        print(f"JSONモード: サンプル {len(sample_rows)} 件 / 分析項目: {len(feature_names)} 項目")

        if player_row is None:
            print(f"エラー: short_id '{args.player}' の個人データが data/my/ に見つかりません")
            return

        player_rank = str(player_row.get("ランク", ""))
        current_rank_group = get_rank_group(player_rank)

        if args.target_rank:
            target_rank_group = args.target_rank.upper()
        else:
            target_rank_group = next_rank(player_rank)
            if target_rank_group is None:
                print(f"プレイヤーはすでに最上位ランク（{current_rank_group}）です。Step 2の比較対象がありません。")
                return

        print(f"プレイヤー: {args.player}  現在: {current_rank_group}  → 目標: {target_rank_group}")

        target_rows = [
            row for row in filter_rows_by_rank(sample_rows, target_rank_group)
            if args.player not in str(row.get("プレイヤー名", ""))
        ]

        if len(target_rows) < 10:
            print(f"警告: {target_rank_group}帯のサンプルが不足しています（{len(target_rows)}件）。精度向上には30件以上推奨。")

        print_step2_comparison(args.player, player_row, current_rank_group, target_rank_group, target_rows, feature_names)

    else:
        csv_path = Path(args.input) if args.input else Path("sf6-playlog-out.csv")
        if not csv_path.exists():
            print(f"エラー: ファイルが見つかりません: {csv_path}")
            return

        columns, rows = load_playlog_rows(csv_path)
        feature_names = get_feature_names(columns)

        player_rows = find_player_rows(rows, args.player)
        if not player_rows:
            print(f"エラー: プレイヤー '{args.player}' が見つかりません")
            return

        player_row = player_rows[-1]
        player_rank = str(player_row.get("ランク", ""))
        current_rank_group = get_rank_group(player_rank)

        target_rank_group = args.target_rank.upper() if args.target_rank else next_rank(player_rank)
        if target_rank_group is None:
            print(f"プレイヤーはすでに最上位ランク（{current_rank_group}）です。")
            return

        target_rows = [
            row for row in filter_rows_by_rank(rows, target_rank_group)
            if args.player not in str(row.get("プレイヤー名", ""))
        ]

        if len(target_rows) < 10:
            print(f"警告: {target_rank_group}帯のサンプルが不足しています（{len(target_rows)}件）。")

        print_step2_comparison(args.player, player_row, current_rank_group, target_rank_group, target_rows, feature_names)


if __name__ == "__main__":
    main()

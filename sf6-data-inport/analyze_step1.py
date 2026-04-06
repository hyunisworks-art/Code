"""SF6 Step1分析: 同ランク帯との比較スクリプト

使い方（JSONモード・推奨）:
    python analyze_step1.py --player 2202760091
    python analyze_step1.py --player 2202760091 --rank DIAMOND

使い方（CSVモード・後方互換）:
    python analyze_step1.py --player "プレイヤー名" --input sf6-playlog-out.csv

JSONモードは data/samples/*.json と data/my/*_<short_id>.json を自動検出する。
--input を明示した場合は常にCSVモードになる。
"""
from __future__ import annotations

import argparse
import json
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

# JSONモードのデフォルトディレクトリ
_SAMPLES_DIR = Path("data/samples")
_MY_DIR = Path("data/my")

# JSONのbattle_statsキー → CSVカラム名のマッピング
# パーセント系フィールド（JSONは0-1小数、出力は"XX.XX%"文字列）
_PCT_FIELDS: list[tuple[str, str]] = [
    ("gauge_rate_drive_guard",           "ドライブパリィ"),
    ("gauge_rate_drive_impact",          "ドライブインパクト"),
    ("gauge_rate_drive_arts",            "オーバードライブアーツ"),
    ("gauge_rate_drive_rush_from_parry", "パリィドライブラッシュ"),
    ("gauge_rate_drive_rush_from_cancel","キャンセルドライブラッシュ"),
    ("gauge_rate_drive_reversal",        "ドライブリバーサル"),
    ("gauge_rate_drive_other",           "ダメージ"),
    ("gauge_rate_sa_lv1",                "Lv1"),
    ("gauge_rate_sa_lv2",                "Lv2"),
    ("gauge_rate_sa_lv3",                "Lv3"),
    ("gauge_rate_ca",                    "CA"),
]
# 数値系フィールド（回数・時間・ポイント）
_NUM_FIELDS: list[tuple[str, str]] = [
    ("drive_reversal",                       "使用回数"),
    ("drive_parry",                          "成功回数"),
    ("throw_drive_parry",                    "相手のドライブパリィを投げた"),
    ("received_throw_drive_parry",           "自分のドライブパリィを投げられた"),
    ("just_parry",                           "ジャストパリィ回数"),
    ("drive_impact",                         "ドライブインパクト_決めた回数"),
    ("punish_counter",                       "パニッシュカウンターを決めた回数"),
    ("drive_impact_to_drive_impact",         "相手のドライブインパクトに決めた回数"),
    ("received_drive_impact",                "ドライブインパクト_受けた回数"),
    ("received_punish_counter",              "パニッシュカウンターを受けた回数"),
    ("received_drive_impact_to_drive_impact","相手にドライブインパクトで返された回数"),
    ("stun",                                 "スタンさせた回数"),
    ("received_stun",                        "スタンさせられた回数"),
    ("throw_count",                          "投げ_決めた回数"),
    ("received_throw_count",                 "投げ_受けた回数"),
    ("throw_tech",                           "投げ抜け回数"),
    ("corner_time",                          "相手を追い詰めている時間"),
    ("cornered_time",                        "相手に追い詰められている時間"),
    ("rank_match_play_count",                "ランクマッチプレイ回数"),
    ("casual_match_play_count",              "カジュアルマッチプレイ回数"),
    ("custom_room_match_play_count",         "ルームマッチプレイ回数"),
    ("battle_hub_match_play_count",          "バトルハブマッチプレイ回数"),
    ("total_all_character_play_point",       "累計プレイポイント"),
]

JSON_FEATURE_NAMES: list[str] = [col for _, col in _PCT_FIELDS + _NUM_FIELDS]


def _json_to_row(data: dict[str, Any]) -> dict[str, Any]:
    """JSON 1件を analyze_step1 が扱える行辞書に変換する。"""
    li = data.get("league_info") or {}
    bs = (data.get("play") or {}).get("battle_stats") or {}

    # ランク: サンプルは top-level "rank"（"diamond" 等）、個人データは league_rank_info
    rank_name: str = (
        (li.get("league_rank_info") or {}).get("league_rank_name")
        or data.get("rank")
        or ""
    )

    row: dict[str, Any] = {
        "No": None,
        "データ取得日": data.get("fetch_date", ""),
        "プレイヤー名": str(data.get("player_id", "")),
        "リーグポイント": li.get("league_point"),
        "ランク": rank_name,
        "MR": li.get("master_rating") or None,
    }

    for json_key, col in _PCT_FIELDS:
        val = bs.get(json_key)
        row[col] = f"{float(val) * 100:.2f}%" if val is not None else ""

    for json_key, col in _NUM_FIELDS:
        val = bs.get(json_key)
        row[col] = str(val) if val is not None else ""

    return row


def load_json_rows(
    short_id: str,
    samples_dir: Path = _SAMPLES_DIR,
    my_dir: Path = _MY_DIR,
) -> tuple[list[str], list[dict[str, Any]], dict[str, Any] | None]:
    """サンプルJSONと個人JSONを読み込む。

    Returns:
        (feature_names, sample_rows, player_row_or_None)
    """
    sample_rows: list[dict[str, Any]] = []
    for path in sorted(samples_dir.glob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            sample_rows.append(_json_to_row(data))
        except Exception:
            continue

    # 個人データ: *_<short_id>.json の最新ファイル
    my_files = sorted(my_dir.glob(f"*_{short_id}.json"))
    player_row: dict[str, Any] | None = None
    if my_files:
        try:
            data = json.loads(my_files[-1].read_text(encoding="utf-8"))
            player_row = _json_to_row(data)
            player_row["プレイヤー名"] = short_id  # find_player_rows の検索キーと一致させる
        except Exception:
            pass

    return JSON_FEATURE_NAMES, sample_rows, player_row


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
    parser.add_argument("--player", required=True, help="short_id（JSONモード）またはプレイヤー名（CSVモード）")
    parser.add_argument("--input", default=None, help="CSVファイルを明示（省略時はJSONモードを自動検出）")
    parser.add_argument("--rank", default=None, help="比較ランク帯（省略時: プレイヤーのランクを使用）")
    args = parser.parse_args()

    # モード判定: --input 未指定かつ data/samples/ に JSON があれば JSON モード
    use_json = args.input is None and _SAMPLES_DIR.exists() and any(_SAMPLES_DIR.glob("*.json"))

    if use_json:
        feature_names, sample_rows, player_row = load_json_rows(args.player)
        print(f"JSONモード: サンプル {len(sample_rows)} 件 / 分析項目: {len(feature_names)} 項目")

        if player_row is None:
            print(f"エラー: short_id '{args.player}' の個人データが {_MY_DIR}/ に見つかりません")
            my_files = sorted(_MY_DIR.glob("*.json")) if _MY_DIR.exists() else []
            if my_files:
                print("利用可能なファイル（最新5件）:")
                for f in my_files[-5:]:
                    print(f"  {f.name}")
            return

        player_rank = str(player_row.get("ランク", ""))
        rank_group = args.rank.upper() if args.rank else get_rank_group(player_rank)

        print(f"プレイヤーID: {args.player}  ランク: {player_rank}  比較帯: {rank_group}")

        rank_rows = [
            row for row in filter_rows_by_rank(sample_rows, rank_group)
            if args.player not in str(row.get("プレイヤー名", ""))
        ]

        if len(rank_rows) < 30:
            print(f"警告: {rank_group}帯のデータが少ないです（{len(rank_rows)}件）。精度向上には30件以上推奨。")

        print_comparison(args.player, player_row, rank_group, rank_rows, feature_names)

    else:
        csv_path = Path(args.input) if args.input else Path("sf6-playlog-out.csv")
        if not csv_path.exists():
            print(f"エラー: ファイルが見つかりません: {csv_path}")
            return

        columns, rows = load_playlog_rows(csv_path)
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

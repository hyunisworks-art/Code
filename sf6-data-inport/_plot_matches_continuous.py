"""sf#05用: プレイ量系指標をLP/MR連続軸で可視化する。
- matches_rank_all.png: ランクマッチプレイ回数のみ（メイン）
- plays_breakdown_all.png: 5種のプレイ量指標（ランクマッチ/累計PP/ルーム/バトルハブ/カジュアル）を重ね

note掲載用に文字・記号サイズを大きめに設定。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from dotenv import load_dotenv
from matplotlib import rcParams

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

rcParams["font.family"] = "Yu Gothic"
rcParams["axes.unicode_minus"] = False
rcParams["font.size"] = 13

OUT_DIR = Path("analysis-output")
OUT_DIR.mkdir(exist_ok=True)

LP_RANKS = {"bronze", "silver", "gold", "platinum", "diamond"}
MR_RANKS = {"master", "master_high", "master_master", "master_grand"}

LP_MIN, LP_MAX = 3000, 24999
MR_MIN, MR_MAX = 1500, 1800
MAX_PER_MR = 12

PLAY_KEYS_MAIN = [
    ("rank_match_play_count", "ランクマッチ"),
]

PLAY_KEYS_ALL = [
    ("rank_match_play_count", "ランクマッチ"),
    ("custom_room_match_play_count", "ルームマッチ"),
    ("battle_hub_match_play_count", "バトルハブ"),
    ("casual_match_play_count", "カジュアルマッチ"),
]


def fetch_all():
    """ローカルJSONサンプル（data/samples/*.json）から読み込む。"""
    import json
    samples_dir = Path("data/samples")
    rows = []
    for p in sorted(samples_dir.glob("*.json")):
        try:
            d = json.loads(p.read_text(encoding="utf-8"))
        except Exception as e:
            continue
        rows.append({
            "player_id": d.get("player_id"),
            "rank": d.get("rank"),
            "play": d.get("play"),
        })
    return rows


def extract_rating(row):
    p = row.get("play") or {}
    cli = p.get("character_league_infos") or []
    lps, mrs = [], []
    for c in cli:
        if not c.get("is_played"):
            continue
        li = c.get("league_info") or {}
        lp = li.get("league_point", -1)
        mr = li.get("master_rating", 0)
        if isinstance(lp, (int, float)) and LP_MIN <= lp <= LP_MAX:
            lps.append(lp)
        if isinstance(mr, (int, float)) and mr >= MR_MIN:
            mrs.append(mr)
    return (max(lps) if lps else None, max(mrs) if mrs else None)


def to_df(rows):
    records = []
    all_keys = [k for k, _ in PLAY_KEYS_ALL]
    for row in rows:
        rank = (row.get("rank") or "").strip()
        bs = ((row.get("play") or {}).get("battle_stats")) or {}
        lp, mr = extract_rating(row)
        rec = {"rank": rank, "lp": lp, "mr": mr}
        for key in all_keys:
            v = bs.get(key)
            rec[key] = float(v) if isinstance(v, (int, float)) else None
        records.append(rec)
    df = pd.DataFrame(records)
    for key in all_keys:
        df[key] = pd.to_numeric(df[key], errors="coerce")
    return df


def smooth(x, y, bandwidth_frac=0.12, n_points=200):
    idx = np.argsort(x)
    xs, ys = np.asarray(x)[idx,], np.asarray(y)[idx]
    if len(xs) < 5:
        return xs, ys
    x_min, x_max = xs.min(), xs.max()
    sigma = (x_max - x_min) * bandwidth_frac
    xi = np.linspace(x_min, x_max, n_points)
    yi = np.empty_like(xi)
    for i, xv in enumerate(xi):
        w = np.exp(-0.5 * ((xs - xv) / sigma) ** 2)
        yi[i] = np.sum(w * ys) / np.sum(w)
    return xi, yi


def _plot_side_by_side(lp_df, mr_df, keys_labels, title, out_path, y_label, y_max=None, is_pp=False):
    fig, (ax_l, ax_r) = plt.subplots(1, 2, figsize=(15, 7.2), gridspec_kw={"width_ratios": [3, 1], "wspace": 0.05}, sharey=True)
    cmap = plt.get_cmap("tab10")

    if y_max is None:
        vals = []
        for key, _ in keys_labels:
            if key in lp_df.columns:
                vals += lp_df[key].dropna().tolist()
            if key in mr_df.columns:
                vals += mr_df[key].dropna().tolist()
        if vals:
            y_max = float(np.percentile(vals, 98)) * 1.12

    # LP panel
    for i, (key, label) in enumerate(keys_labels):
        sub = lp_df[["lp", key]].dropna()
        if len(sub) < 10:
            continue
        color = cmap(i % 10)
        ax_l.scatter(sub["lp"], sub[key], s=22, alpha=0.22, color=color)
        xs, ys = smooth(sub["lp"].to_numpy(), sub[key].to_numpy())
        ax_l.plot(xs, ys, linewidth=3.2, color=color, label=label)
    ax_l.set_xlim(LP_MIN, LP_MAX)
    ax_l.set_xlabel("LP（リーグポイント）", fontsize=15)
    ax_l.set_ylabel(y_label, fontsize=15)
    ax_l.set_title(f"LP帯（ブロンズ〜ダイヤ・n={len(lp_df)}）", fontsize=15)
    ax_l.grid(True, alpha=0.3)
    ax_l.tick_params(axis="both", labelsize=13)

    # MR panel
    for i, (key, label) in enumerate(keys_labels):
        sub = mr_df[["mr", key]].dropna()
        if len(sub) < 10:
            continue
        color = cmap(i % 10)
        ax_r.scatter(sub["mr"], sub[key], s=22, alpha=0.22, color=color)
        xs, ys = smooth(sub["mr"].to_numpy(), sub[key].to_numpy())
        ax_r.plot(xs, ys, linewidth=3.2, color=color, label=label)
    ax_r.set_xlim(MR_MIN, MR_MAX)
    ax_r.set_xlabel("MR（マスターレート）", fontsize=15)
    ax_r.set_title(f"MR帯（マスター〜ハイマス・n={len(mr_df)}）", fontsize=15)
    ax_r.grid(True, alpha=0.3)
    ax_r.tick_params(axis="both", labelsize=13)

    if y_max:
        ax_l.set_ylim(0, y_max)
        ax_r.set_ylim(0, y_max)

    # y軸のフォーマット（大きい数値用にカンマ）
    ax_l.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{int(x):,}"))

    # 凡例（右パネルの中・上寄り）
    handles, labels = ax_l.get_legend_handles_labels()
    if handles:
        ax_r.legend(handles, labels, loc="upper left", fontsize=13, frameon=True, framealpha=0.9)

    fig.suptitle(title, fontsize=19, y=0.99)
    fig.subplots_adjust(top=0.88, bottom=0.11, left=0.08, right=0.98)
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  saved: {out_path}")


def main():
    print("Supabaseから取得中...")
    rows = fetch_all()
    print(f"  取得 {len(rows)} 件")
    df = to_df(rows)

    mr_df = df[df["mr"].notna() & (df["mr"] <= MR_MAX)]
    mr_df = mr_df.sort_values("mr").groupby("mr", group_keys=False).head(MAX_PER_MR)
    lp_df = df[df["mr"].isna() & df["lp"].notna()]
    print(f"LP帯サンプル: {len(lp_df)} / MR帯サンプル: {len(mr_df)}")

    # 1. メイン: ランクマッチ回数のみ
    _plot_side_by_side(
        lp_df, mr_df, PLAY_KEYS_MAIN,
        "ランクマッチのプレイ回数 — ランク帯別の分布",
        OUT_DIR / "matches_rank_all.png",
        y_label="ランクマッチプレイ回数",
    )

    # 2. 4種のプレイ種別を重ねて比較（カジュアル反転を可視化）
    _plot_side_by_side(
        lp_df, mr_df, PLAY_KEYS_ALL,
        "プレイ種別ごとの試合数 — ランク帯別の分布",
        OUT_DIR / "plays_breakdown_all.png",
        y_label="プレイ回数",
    )


if __name__ == "__main__":
    main()

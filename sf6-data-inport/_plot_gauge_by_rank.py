"""ランク帯別のゲージ使用率を折れ線グラフ化する。
LP帯（bronze〜diamond）・MR帯（master〜master_grand）、Dゲージ・SAゲージで4枚。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from dotenv import load_dotenv
from matplotlib import rcParams

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

rcParams["font.family"] = "Yu Gothic"
rcParams["axes.unicode_minus"] = False

OUT_DIR = Path("analysis-output")
OUT_DIR.mkdir(exist_ok=True)

RANK_LP = ["bronze", "silver", "gold", "platinum", "diamond"]
RANK_MR = ["master", "master_high", "master_master", "master_grand"]
RANK_ALL = RANK_LP + RANK_MR
RANK_LABEL = {
    "bronze": "ブロンズ", "silver": "シルバー", "gold": "ゴールド",
    "platinum": "プラチナ", "diamond": "ダイヤ",
    "master": "マスター", "master_high": "M.High",
    "master_master": "M.Master", "master_grand": "M.Grand",
}

DRIVE_KEYS = [
    ("gauge_rate_drive_impact", "ドライブインパクト"),
    ("gauge_rate_drive_rush_from_cancel", "ラッシュ(キャンセル)"),
    ("gauge_rate_drive_rush_from_parry", "ラッシュ(パリィ)"),
    ("gauge_rate_drive_guard", "ガード"),
    ("gauge_rate_drive_reversal", "ドライブリバーサル"),
    ("gauge_rate_drive_arts", "ODアーツ"),
    ("gauge_rate_drive_other", "その他"),
]
SA_KEYS = [
    ("gauge_rate_sa_lv1", "SA1"),
    ("gauge_rate_sa_lv2", "SA2"),
    ("gauge_rate_sa_lv3", "SA3"),
    ("gauge_rate_ca", "CA"),
]


def fetch_all():
    from supabase import create_client
    client = create_client(os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_KEY"))
    rows = []
    page = 100
    start = 0
    while True:
        r = (
            client.table("player_data")
            .select("rank,data_type,play")
            .eq("data_type", "sample")
            .range(start, start + page - 1)
            .execute()
        )
        if not r.data:
            break
        rows.extend(r.data)
        if len(r.data) < page:
            break
        start += page
    return rows


def to_df(rows):
    records = []
    all_keys = [k for k, _ in DRIVE_KEYS + SA_KEYS]
    for row in rows:
        rank = (row.get("rank") or "").strip()
        bs = ((row.get("play") or {}).get("battle_stats")) or {}
        rec = {"rank": rank}
        for key in all_keys:
            v = bs.get(key)
            rec[key] = float(v) if isinstance(v, (int, float)) else None
        records.append(rec)
    return pd.DataFrame(records)


def rank_means(df, ranks, keys):
    rows = []
    for rank in ranks:
        sub = df[df["rank"] == rank]
        row = {"rank": rank, "n": len(sub)}
        for key in keys:
            vals = sub[key].dropna()
            row[key] = float(vals.mean() * 100) if len(vals) else None
        rows.append(row)
    return pd.DataFrame(rows)


def plot_lines(agg, keys_labels, ranks, title, out_path, y_max=None, split_at=None):
    fig, ax = plt.subplots(figsize=(11, 5.5))
    x = list(range(len(ranks)))
    xlabels = [f"{RANK_LABEL[r]}\n(n={agg[agg['rank']==r]['n'].iloc[0]})" for r in ranks]
    cmap = plt.get_cmap("tab10")
    for i, (key, label) in enumerate(keys_labels):
        y = [agg[agg["rank"] == r][key].iloc[0] for r in ranks]
        ax.plot(x, y, marker="o", linewidth=2, label=label, color=cmap(i % 10))
    ax.set_xticks(x)
    ax.set_xticklabels(xlabels)
    ax.set_ylabel("使用率（%）")
    ax.set_title(title)
    ax.grid(True, alpha=0.3)
    if y_max is not None:
        ax.set_ylim(0, y_max)
    if split_at is not None:
        ax.axvline(x=split_at - 0.5, color="gray", linestyle="--", alpha=0.6)
        ax.text(split_at - 0.5 - 0.05, y_max * 0.97 if y_max else 0, "LP帯",
                ha="right", va="top", fontsize=10, color="gray")
        ax.text(split_at - 0.5 + 0.05, y_max * 0.97 if y_max else 0, "MR帯",
                ha="left", va="top", fontsize=10, color="gray")
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0), fontsize=9)
    fig.tight_layout()
    fig.savefig(out_path, dpi=140, bbox_inches="tight")
    plt.close(fig)
    print(f"  saved: {out_path}")


def main():
    print("Supabaseから取得中...")
    rows = fetch_all()
    print(f"  取得 {len(rows)} 件")
    df = to_df(rows)

    d_keys = [k for k, _ in DRIVE_KEYS]
    sa_keys = [k for k, _ in SA_KEYS]

    agg_all = rank_means(df, RANK_ALL, d_keys + sa_keys)

    # D/SA 共通のY軸上限を計算
    d_max = float(agg_all[d_keys].max().max())
    sa_max = float(agg_all[sa_keys].max().max())
    y_max_common = max(d_max, sa_max) * 1.15

    split = len(RANK_LP)  # LP/MR境界のインデックス（5）

    plot_lines(agg_all, DRIVE_KEYS, RANK_ALL,
               "Dゲージ使用率 — ランク別推移（ブロンズ〜マスターGrand）",
               OUT_DIR / "gauge_drive_all.png",
               y_max=y_max_common, split_at=split)
    plot_lines(agg_all, SA_KEYS, RANK_ALL,
               "SAゲージ使用率 — ランク別推移（ブロンズ〜マスターGrand）",
               OUT_DIR / "gauge_sa_all.png",
               y_max=y_max_common, split_at=split)

    print("\nランク別サンプル数:")
    for rank in RANK_LP + RANK_MR:
        print(f"  {RANK_LABEL[rank]}: {len(df[df['rank']==rank])}")


if __name__ == "__main__":
    main()

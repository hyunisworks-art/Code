"""LP/MR連続軸でゲージ使用率を散布＋平滑化ラインで可視化する。
LP帯: league_point（is_playedキャラのmax値）
MR帯: master_rating（is_playedキャラのmax値）
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

OUT_DIR = Path("analysis-output")
OUT_DIR.mkdir(exist_ok=True)

LP_RANKS = {"bronze", "silver", "gold", "platinum", "diamond"}
MR_RANKS = {"master", "master_high", "master_master", "master_grand"}

DRIVE_KEYS = [
    ("gauge_rate_drive_other", "ドライブパリィ"),
    ("gauge_rate_drive_impact", "ドライブインパクト"),
    ("gauge_rate_drive_arts", "オーバードライブアーツ"),
    ("gauge_rate_drive_rush_from_parry", "パリィドライブラッシュ"),
    ("gauge_rate_drive_rush_from_cancel", "キャンセルドライブラッシュ"),
    ("gauge_rate_drive_reversal", "ドライブリバーサル"),
    ("gauge_rate_drive_guard", "ダメージ"),
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
            .select("player_id,rank,data_type,play")
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


LP_MIN, LP_MAX = 3000, 24999  # ブロンズ〜ダイヤの範囲（マスター以降はLPが累積してしまうため）
MR_MIN = 1500  # マスター到達ライン
MR_MAX = 1800  # ハイマス到達ライン（これより上はサンプル偏りが強いため今回は無視）


def extract_rating(row):
    """play.character_league_infos から、is_played キャラのLP/MRを返す。
    LP は 3000〜24999 の範囲内のみ採用（マスター以降の累積LPを除外）。
    """
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
    all_keys = [k for k, _ in DRIVE_KEYS + SA_KEYS]
    for row in rows:
        rank = (row.get("rank") or "").strip()
        bs = ((row.get("play") or {}).get("battle_stats")) or {}
        lp, mr = extract_rating(row)
        rec = {"rank": rank, "lp": lp, "mr": mr}
        for key in all_keys:
            v = bs.get(key)
            rec[key] = float(v) * 100 if isinstance(v, (int, float)) else None
        records.append(rec)
    df = pd.DataFrame(records)
    for key in all_keys:
        df[key] = pd.to_numeric(df[key], errors="coerce")
    return df


def smooth(x, y, bandwidth_frac=0.12, n_points=200):
    """ガウシアンカーネル平滑化。完全に滑らかな曲線を返す。
    bandwidth_frac: x範囲に対するバンド幅（大きいほど滑らか）
    """
    idx = np.argsort(x)
    xs, ys = np.asarray(x)[idx, ], np.asarray(y)[idx]
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


MAX_PER_MR = 12  # 同じMR値への偏りを抑えるための上限（縦帯回避）


def plot_single(df, x_col, keys_labels, title, xlabel, xlim, out_path, y_max=None):
    """1帯を1枚のグラフとして出力（散布＋ガウシアン平滑化）"""
    if y_max is None:
        y_max = float(df[[k for k, _ in keys_labels]].max().max()) * 1.15

    fig, ax = plt.subplots(figsize=(10, 6))
    cmap = plt.get_cmap("tab10")
    for i, (key, label) in enumerate(keys_labels):
        sub = df[[x_col, key]].dropna()
        if len(sub) < 10:
            continue
        color = cmap(i % 10)
        ax.scatter(sub[x_col], sub[key], s=14, alpha=0.20, color=color)
        xs, ys = smooth(sub[x_col].to_numpy(), sub[key].to_numpy())
        ax.plot(xs, ys, linewidth=2.4, color=color, label=label)
    ax.set_title(title, fontsize=12)
    ax.set_xlabel(xlabel)
    ax.set_ylabel("使用率（%）")
    ax.set_ylim(0, y_max)
    ax.set_xlim(*xlim)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0), fontsize=9)
    fig.tight_layout()
    fig.savefig(out_path, dpi=140, bbox_inches="tight")
    plt.close(fig)
    print(f"  saved: {out_path}")


def plot_gauge(df, keys_labels, gauge_label, out_dir):
    """Dゲージ or SAゲージ について、LP帯・MR帯を別ファイルで描画"""
    mr_df = df[df["mr"].notna() & (df["mr"] <= MR_MAX)]
    mr_df = mr_df.sort_values("mr").groupby("mr", group_keys=False).head(MAX_PER_MR)
    lp_df = df[df["mr"].isna() & df["lp"].notna()]

    # Y軸上限は60%で固定（変化幅が小さく、100%だと薄く見えるため）
    y_max = 60

    slug = "drive" if "D" in gauge_label else "sa"
    plot_single(
        lp_df, "lp", keys_labels,
        f"{gauge_label}使用率 — LP帯（n={len(lp_df)}）",
        "LP（リーグポイント）",
        (LP_MIN, LP_MAX),
        out_dir / f"gauge_{slug}_lp.png",
        y_max=y_max,
    )
    plot_single(
        mr_df, "mr", keys_labels,
        f"{gauge_label}使用率 — MR帯（n={len(mr_df)}）",
        "MR（マスターレート）",
        (MR_MIN, MR_MAX),
        out_dir / f"gauge_{slug}_mr.png",
        y_max=y_max,
    )


def main():
    print("Supabaseから取得中...")
    rows = fetch_all()
    print(f"  取得 {len(rows)} 件")
    df = to_df(rows)

    lp_avail = df[df["rank"].isin(LP_RANKS) & df["lp"].notna()]
    mr_avail = df[df["rank"].isin(MR_RANKS) & df["mr"].notna()]
    print(f"LP取得成功: {len(lp_avail)} / {len(df[df['rank'].isin(LP_RANKS)])}")
    print(f"MR取得成功: {len(mr_avail)} / {len(df[df['rank'].isin(MR_RANKS)])}")
    print(f"LP分布: min={lp_avail['lp'].min()} / max={lp_avail['lp'].max()} / median={lp_avail['lp'].median()}")
    print(f"MR分布: min={mr_avail['mr'].min()} / max={mr_avail['mr'].max()} / median={mr_avail['mr'].median()}")

    plot_gauge(df, DRIVE_KEYS, "Dゲージ", OUT_DIR)
    plot_gauge(df, SA_KEYS, "SAゲージ", OUT_DIR)


if __name__ == "__main__":
    main()

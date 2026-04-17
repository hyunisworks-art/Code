"""ダイヤ帯とマスター帯のSAゲージ生データを突き合わせる。
SA1/SA2/SA3/CA の値・合計・分布を確認する。
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

OUT = Path("analysis-output/sa_boundary_inspect.md")

SA_KEYS = [
    ("gauge_rate_sa_lv1", "SA1"),
    ("gauge_rate_sa_lv2", "SA2"),
    ("gauge_rate_sa_lv3", "SA3"),
    ("gauge_rate_ca", "CA"),
]
D_KEYS = [
    ("gauge_rate_drive_arts", "ODアーツ"),
    ("gauge_rate_drive_impact", "DI"),
    ("gauge_rate_drive_parry", "Dパリィ"),
    ("gauge_rate_drive_guard", "ガード"),
    ("gauge_rate_drive_reversal", "リバーサル"),
    ("gauge_rate_drive_rush_from_parry", "RUSH(パリィ)"),
    ("gauge_rate_drive_rush_from_cancel", "RUSH(キャン)"),
    ("gauge_rate_drive_other", "その他"),
]


def fetch(rank):
    from supabase import create_client
    c = create_client(os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_KEY"))
    rows = []
    page = 100
    start = 0
    while True:
        r = (
            c.table("player_data")
            .select("player_id,rank,data_type,play")
            .eq("data_type", "sample")
            .eq("rank", rank)
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


def row_to_sa(row):
    bs = ((row.get("play") or {}).get("battle_stats")) or {}
    rec = {"player_id": row.get("player_id")}
    for key, label in SA_KEYS:
        v = bs.get(key)
        rec[label] = float(v) * 100 if isinstance(v, (int, float)) else None
    for key, label in D_KEYS:
        v = bs.get(key)
        rec[f"D:{label}"] = float(v) * 100 if isinstance(v, (int, float)) else None
    return rec


def describe(df, col):
    s = df[col].dropna()
    if len(s) == 0:
        return f"  {col}: データなし"
    return (
        f"  {col:14s}: n={len(s):3d} mean={s.mean():6.2f}% median={s.median():6.2f}%"
        f" min={s.min():6.2f}% max={s.max():6.2f}% std={s.std():5.2f}"
    )


def main():
    lines = ["# SAゲージ集計・境界検証（ダイヤ vs マスター）", ""]
    for rank in ["diamond", "master"]:
        print(f"fetching {rank}...")
        rows = fetch(rank)
        df = pd.DataFrame([row_to_sa(r) for r in rows])
        sa_cols = [l for _, l in SA_KEYS]
        d_cols = [f"D:{l}" for _, l in D_KEYS]
        # NoneをNaNに揃える（sum等で型エラー回避）
        for col in sa_cols + d_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        lines.append(f"## {rank}  n={len(df)}")
        lines.append("")
        lines.append("### SA系 記述統計")
        for col in sa_cols:
            lines.append(describe(df, col))
        lines.append("")

        # SA4種の合計（行ごと）
        sa_sum = df[sa_cols].sum(axis=1, skipna=False)
        lines.append("### SA4種 行合計の分布（本来は100%想定）")
        lines.append(
            f"  SA合計: n={sa_sum.dropna().shape[0]} mean={sa_sum.mean():.2f}%"
            f" median={sa_sum.median():.2f}% min={sa_sum.min():.2f}% max={sa_sum.max():.2f}%"
        )
        lines.append("")

        # Dゲージ側も参考に合計
        d_sum = df[d_cols].sum(axis=1, skipna=False)
        lines.append("### Dゲージ8種 行合計の分布（本来は100%想定）")
        lines.append(
            f"  D合計:  n={d_sum.dropna().shape[0]} mean={d_sum.mean():.2f}%"
            f" median={d_sum.median():.2f}% min={d_sum.min():.2f}% max={d_sum.max():.2f}%"
        )
        lines.append("")

        # 欠損カウント
        null_sa = df[sa_cols].isna().sum()
        lines.append("### SA系 欠損件数")
        for col in sa_cols:
            lines.append(f"  {col}: 欠損 {null_sa[col]}件 / 合計 {len(df)}件")
        lines.append("")

        # サンプル5件プレビュー（SA系）
        lines.append("### サンプル先頭5件（SA系のみ・%表示）")
        lines.append("```")
        head = df[["player_id"] + sa_cols].head(5)
        head_fmt = head.copy()
        for col in sa_cols:
            head_fmt[col] = head_fmt[col].apply(
                lambda x: f"{x:6.2f}" if pd.notna(x) else "  NaN "
            )
        lines.append(head_fmt.to_string(index=False))
        lines.append("```")
        lines.append("")

        # 参考: SA2が他ランクと比べて異常に高い人を探す（マスター帯だけ見る目的）
        if rank == "master":
            top_sa2 = df.nlargest(5, "SA2")[["player_id"] + sa_cols]
            top_sa2_fmt = top_sa2.copy()
            for col in sa_cols:
                top_sa2_fmt[col] = top_sa2_fmt[col].apply(
                    lambda x: f"{x:6.2f}" if pd.notna(x) else "  NaN "
                )
            lines.append("### マスター: SA2が高い上位5名")
            lines.append("```")
            lines.append(top_sa2_fmt.to_string(index=False))
            lines.append("```")
            lines.append("")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"written: {OUT}")


if __name__ == "__main__":
    main()

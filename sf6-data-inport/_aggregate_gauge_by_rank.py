"""ランク順序×ゲージ使用率のSpearman相関を見る。
相関が強い指標をピックし、ランク推移をセットで出す。"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

OUT_MD = Path("gauge_by_rank_report.md")

RANK_ORDER = [
    "bronze", "silver", "gold", "platinum", "diamond",
    "master", "master_high", "master_master", "master_grand",
]
RANK_INDEX = {r: i + 1 for i, r in enumerate(RANK_ORDER)}
RANK_LABEL = {
    "bronze": "ブロンズ", "silver": "シルバー", "gold": "ゴールド",
    "platinum": "プラチナ", "diamond": "ダイヤ", "master": "マスター",
    "master_high": "マスターHigh", "master_master": "マスターMaster",
    "master_grand": "マスターGrand",
}

DRIVE_KEYS = [
    ("gauge_rate_drive_arts", "ODアーツ"),
    ("gauge_rate_drive_impact", "ドライブインパクト"),
    ("gauge_rate_drive_parry", "ドライブパリィ"),
    ("gauge_rate_drive_guard", "ガード"),
    ("gauge_rate_drive_reversal", "ドライブリバーサル"),
    ("gauge_rate_drive_rush_from_parry", "ラッシュ(パリィから)"),
    ("gauge_rate_drive_rush_from_cancel", "ラッシュ(キャンセルから)"),
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
        print(f"  取得 {len(rows)}件")
        if len(r.data) < page:
            break
        start += page
    return rows


def main():
    rows = fetch_all()
    print(f"sample行数: {len(rows)}")

    # DataFrame化
    records = []
    for row in rows:
        rank = (row.get("rank") or "").strip()
        if rank not in RANK_INDEX:
            continue
        bs = ((row.get("play") or {}).get("battle_stats")) or {}
        rec = {"rank": rank, "rank_idx": RANK_INDEX[rank]}
        for key, _ in DRIVE_KEYS + SA_KEYS:
            v = bs.get(key)
            rec[key] = float(v) if isinstance(v, (int, float)) else None
        records.append(rec)
    df = pd.DataFrame(records)
    print(f"有効行: {len(df)}")

    # Spearman相関（ランク順序 × 各指標）
    corr_rows = []
    for key, label in DRIVE_KEYS + SA_KEYS:
        sub = df[["rank_idx", key]].dropna()
        n = len(sub)
        if n < 30:
            corr_rows.append((label, key, None, n, "D" if "drive" in key else "SA"))
            continue
        rho = sub["rank_idx"].rank().corr(sub[key].rank(), method="pearson")
        corr_rows.append((label, key, rho, n, "D" if "drive" in key else "SA"))

    # |ρ|でソート
    corr_rows_sorted = sorted(
        corr_rows,
        key=lambda x: abs(x[2]) if x[2] is not None else -1,
        reverse=True,
    )

    lines = []
    lines.append("# ランク別 ゲージ使用率 相関分析（sf6記事#4素材）")
    lines.append("")
    lines.append(f"- データ: sample {len(df)}件（Supabase player_data）")
    lines.append(f"- 指標: Dゲージ8種 / SAゲージ4種")
    lines.append(f"- 手法: ランク順序（1=ブロンズ〜9=マスターGrand）× 指標値のSpearman順位相関")
    lines.append(f"- 読み方: ρ>0 は「ランクが上がると使用率が上がる」／ρ<0 は「下がる」")
    lines.append(f"  |ρ|: ~0.1弱い / 0.2〜0.4中程度 / 0.4〜0.6強い / 0.6+非常に強い")
    lines.append("")
    lines.append("## 相関の強い順（全12指標）")
    lines.append("")
    lines.append("| 系統 | 指標 | Spearman ρ | n |")
    lines.append("|---|---|---:|---:|")
    for label, key, rho, n, sys_ in corr_rows_sorted:
        rho_s = "—" if rho is None else f"{rho:+.3f}"
        lines.append(f"| {sys_} | {label} | {rho_s} | {n} |")
    lines.append("")

    # |ρ|>=0.2 の指標についてランク別平均値を出す
    interesting = [r for r in corr_rows_sorted if r[2] is not None and abs(r[2]) >= 0.2]
    if interesting:
        lines.append("## 傾向が出ている指標のランク別平均（|ρ|≧0.2）")
        lines.append("")
        for label, key, rho, n, sys_ in interesting:
            lines.append(f"### {label}（{sys_}系・ρ={rho:+.3f}）")
            lines.append("")
            lines.append("| ランク | 平均使用率 | 中央値 | n |")
            lines.append("|---|---:|---:|---:|")
            for rank in RANK_ORDER:
                sub = df[df["rank"] == rank][key].dropna()
                if len(sub) == 0:
                    lines.append(f"| {RANK_LABEL[rank]} | — | — | 0 |")
                else:
                    m = sub.mean() * 100
                    md = sub.median() * 100
                    lines.append(f"| {RANK_LABEL[rank]} | {m:.2f}% | {md:.2f}% | {len(sub)} |")
            lines.append("")
    else:
        lines.append("## 補記")
        lines.append("")
        lines.append("|ρ|≧0.2 の指標なし（ランクとゲージ使用率の明確な相関は出ていない）")
        lines.append("")

    # 参考: ランク別サンプル数
    lines.append("## 参考: ランク別サンプル数")
    lines.append("")
    lines.append("| ランク | n |")
    lines.append("|---|---:|")
    for rank in RANK_ORDER:
        n = len(df[df["rank"] == rank])
        lines.append(f"| {RANK_LABEL[rank]} | {n} |")
    lines.append("")

    OUT_MD.write_text("\n".join(lines), encoding="utf-8")
    print(f"出力: {OUT_MD}")


if __name__ == "__main__":
    main()

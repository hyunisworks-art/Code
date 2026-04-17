"""最高到達LP・最高到達MR でサンプル分類カウント（extract_rating 本番ロジックと一致）"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import load_dotenv

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

LP_MIN, LP_MAX = 3000, 24999
MR_MIN, MR_MAX = 1500, 1800

OUT = Path("analysis-output/sample_split_counts.md")


def fetch_all():
    from supabase import create_client
    c = create_client(os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_KEY"))
    rows, start = [], 0
    while True:
        r = c.table("player_data").select("rank,play").eq("data_type","sample").range(start, start+99).execute()
        if not r.data: break
        rows.extend(r.data)
        if len(r.data) < 100: break
        start += 100
    return rows


def player_max(row):
    """そのプレイヤーの is_played 全キャラにおける最高LP・最高MRを返す"""
    p = row.get("play") or {}
    cli = p.get("character_league_infos") or []
    lps, mrs = [], []
    for c in cli:
        if not c.get("is_played"): continue
        li = c.get("league_info") or {}
        lp = li.get("league_point", -1)
        mr = li.get("master_rating", 0)
        if isinstance(lp, (int, float)) and lp > 0:
            lps.append(lp)
        if isinstance(mr, (int, float)) and mr > 0:
            mrs.append(mr)
    return (max(lps) if lps else None, max(mrs) if mrs else None)


def main():
    rows = fetch_all()
    total = len(rows)

    cat = {
        "MR帯対象（最高MR 1500-1800）": 0,
        "LP帯対象（最高LP 3000-24999・最高MR<1500 or なし）": 0,
        "除外: 最高MR > 1800（ハイマス以上）": 0,
        "除外: 最高MR < 1500（マスター未到達扱い）かつ LP > 24999": 0,
        "除外: 最高MR < 1500 かつ LP < 3000（極低ランク）": 0,
        "除外: LP・MRどちらも取れない": 0,
    }

    for row in rows:
        max_lp, max_mr = player_max(row)

        # MR優先判定
        if max_mr is not None:
            if MR_MIN <= max_mr <= MR_MAX:
                cat["MR帯対象（最高MR 1500-1800）"] += 1
                continue
            if max_mr > MR_MAX:
                cat["除外: 最高MR > 1800（ハイマス以上）"] += 1
                continue
            # max_mr < MR_MIN: マスター未到達（ありえない？未到達の人は MR=0 なのでそもそも max_mr=None）
            # → 実質このパスは来ない

        # MRなし（マスター未到達）・LPで判定
        if max_lp is None:
            cat["除外: LP・MRどちらも取れない"] += 1
        elif max_lp < LP_MIN:
            cat["除外: 最高MR < 1500 かつ LP < 3000（極低ランク）"] += 1
        elif max_lp > LP_MAX:
            cat["除外: 最高MR < 1500（マスター未到達扱い）かつ LP > 24999"] += 1
        else:
            cat["LP帯対象（最高LP 3000-24999・最高MR<1500 or なし）"] += 1

    lines = ["# サンプル分類カウント（最高到達LP・最高到達MR基準）", ""]
    lines.append(f"**全サンプル: {total}件**")
    lines.append("")
    lines.append("## 分類ルール")
    lines.append("- LP範囲: 3000 〜 24999（最高到達LP）")
    lines.append("- MR範囲: 1500 〜 1800（最高到達MR・ハイマスまで）")
    lines.append("- 最高到達MRが取れる人は MR帯優先で判定。範囲外ならそのプレイヤーは除外")
    lines.append("- 最高到達MRが無い人（マスター未到達）は LP で判定")
    lines.append("")
    lines.append("## 分類内訳")
    lines.append("")
    lines.append("| 区分 | n |")
    lines.append("|---|---:|")
    for k, v in cat.items():
        lines.append(f"| {k} | {v} |")
    lines.append(f"| **合計** | {sum(cat.values())} |")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"written: {OUT}")


if __name__ == "__main__":
    main()

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st

import analyze_playlog as ap

MASTER_LP_THRESHOLD = 25000
DEFAULT_INPUT = "sf6-playlog-out.csv"

CHAR_DEPENDENT_FEATURES = {
    "Lv1",
    "Lv2",
    "Lv3",
    "CA",
    "オーバードライブアーツ",
    "ドライブインパクト",
    "ドライブインパクト_決めた回数",
    "投げ_決めた回数",
}

PLAY_VOLUME_FEATURES = {
    "ランクマッチプレイ回数",
    "カジュアルマッチプレイ回数",
    "ルームマッチプレイ回数",
    "バトルハブマッチプレイ回数",
    "累計プレイポイント",
}

LP_BANDS: list[tuple[int, int, str]] = [
    (9000, 12999, "9k-12k"),
    (13000, 16999, "13k-16k"),
    (17000, 20999, "17k-20k"),
    (21000, 24999, "21k-24k"),
    (25000, 49999, "25k-49k"),
    (50000, 99999, "50k-99k"),
    (100000, 249999, "100k-249k"),
    (250000, 10**9, "250k+"),
]

LP_BAND_RANK_HINT = {
    "9k-12k": "GOLD1-GOLD5",
    "13k-16k": "PLAT1-PLAT4",
    "17k-20k": "PLAT4-DIAMOND1",
    "21k-24k": "DIAMOND2-DIAMOND5",
    "25k-49k": "MASTER",
    "50k-99k": "MASTER",
    "100k-249k": "MASTER",
    "250k+": "MASTER",
}


def lp_band(lp: float | None) -> str:
    if lp is None:
        return "不明"
    for lo, hi, label in LP_BANDS:
        if lo <= lp <= hi:
            return label
    return "範囲外"


def sort_key_for_band(label: str) -> int:
    order = {name: idx for idx, (_, _, name) in enumerate(LP_BANDS)}
    return order.get(label, 10**9)


@st.cache_data(show_spinner=False)
def load_source(path_text: str) -> tuple[list[str], list[dict[str, Any]]]:
    path = Path(path_text)
    if not path.exists():
        raise FileNotFoundError(f"入力CSVが見つかりません: {path}")
    return ap.load_playlog_rows(path)


def rows_to_dataframe(columns: list[str], rows: list[dict[str, Any]]) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    for row in rows:
        rec: dict[str, Any] = {}
        for col in columns:
            value = row.get(col)
            if isinstance(value, (int, float)):
                rec[col] = float(value)
            else:
                rec[col] = value
        records.append(rec)

    df = pd.DataFrame(records)
    for col in columns:
        if col in {"データ取得日", "プレイヤー名", "ランク"}:
            continue
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["LP帯"] = df["リーグポイント"].apply(lp_band)
    df["モデル区分"] = df["リーグポイント"].apply(
        lambda x: "LPモデル" if pd.notna(x) and x < MASTER_LP_THRESHOLD else "MRモデル"
    )
    return df


def summarize_counts(df: pd.DataFrame) -> tuple[int, int, int]:
    total = int(df["リーグポイント"].notna().sum())
    sub = int((df["リーグポイント"] < MASTER_LP_THRESHOLD).sum())
    master = int((df["リーグポイント"] >= MASTER_LP_THRESHOLD).sum())
    return total, sub, master


def band_count_df(df: pd.DataFrame) -> pd.DataFrame:
    counts = (
        df["LP帯"]
        .value_counts(dropna=False)
        .rename_axis("LP帯")
        .reset_index(name="件数")
    )
    counts = counts[counts["LP帯"].isin([name for _, _, name in LP_BANDS])].copy()
    counts["表示LP帯"] = counts["LP帯"].apply(
        lambda x: f"{x}\n({LP_BAND_RANK_HINT.get(x, '-')})"
    )
    counts["目標80との差"] = 80 - counts["件数"]
    counts["最低60との差"] = 60 - counts["件数"]
    counts["LP帯ソート"] = counts["LP帯"].apply(sort_key_for_band)
    counts = counts.sort_values("LP帯ソート").drop(columns=["LP帯ソート"]) 
    return counts


def filter_features(
    features: list[str],
    exclude_char_dependent: bool,
    exclude_play_volume: bool,
) -> list[str]:
    filtered = features[:]
    if exclude_char_dependent:
        filtered = [f for f in filtered if f not in CHAR_DEPENDENT_FEATURES]
    if exclude_play_volume:
        filtered = [f for f in filtered if f not in PLAY_VOLUME_FEATURES]
    return filtered


def top_positive_negative(results: list[dict[str, Any]], top_n: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not results:
        empty = pd.DataFrame(columns=["feature", "correlation", "n", "r_squared"])
        return empty, empty

    positives = [r for r in results if r["correlation"] > 0][:top_n]
    negatives = sorted([r for r in results if r["correlation"] < 0], key=lambda x: x["correlation"])[:top_n]

    pos_df = pd.DataFrame(positives)[["feature", "correlation", "n", "r_squared"]]
    neg_df = pd.DataFrame(negatives)[["feature", "correlation", "n", "r_squared"]]
    return pos_df, neg_df


def plot_factor_bar(df: pd.DataFrame, title: str) -> None:
    if df.empty:
        st.info("表示できる要因がありません。")
        return

    chart_df = df.copy()
    chart_df["abs_corr"] = chart_df["correlation"].abs()
    fig = px.bar(
        chart_df,
        x="correlation",
        y="feature",
        orientation="h",
        color="correlation",
        color_continuous_scale="RdBu",
        title=title,
        hover_data={"n": True, "r_squared": ":.3f", "correlation": ":.3f", "feature": True},
    )
    fig.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10), coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True)


def make_summary_text(
    total: int,
    sub_n: int,
    master_n: int,
    band_df: pd.DataFrame,
    lp_results: list[dict[str, Any]],
    mr_results: list[dict[str, Any]],
) -> str:
    def top_pos(results: list[dict[str, Any]]) -> dict[str, Any] | None:
        for r in results:
            if r["correlation"] > 0:
                return r
        return None

    def top_neg(results: list[dict[str, Any]]) -> dict[str, Any] | None:
        negatives = [r for r in results if r["correlation"] < 0]
        if not negatives:
            return None
        return sorted(negatives, key=lambda x: x["correlation"])[0]

    lp_pos = top_pos(lp_results)
    lp_neg = top_neg(lp_results)
    mr_pos = top_pos(mr_results)
    mr_neg = top_neg(mr_results)

    shortage_bands = band_df[band_df["最低60との差"] > 0]
    if shortage_bands.empty:
        band_comment = "全LP帯で最低60件を満たしており、帯別比較の最低ラインを確保しています。"
    else:
        labels = ", ".join(shortage_bands["LP帯"].tolist())
        band_comment = f"最低60件を下回る帯があります（{labels}）。この帯は追加取得を推奨します。"

    lines = [
        f"今回の分析対象は総件数{total}件で、LPモデル{sub_n}件・MRモデル{master_n}件です。",
        band_comment,
    ]

    if lp_neg is not None:
        lines.append(
            f"LPモデルでは「{lp_neg['feature']}」が最も強い負相関（r={lp_neg['correlation']:.3f}, n={lp_neg['n']}）を示しました。"
        )
    if lp_pos is not None:
        lines.append(
            f"LPモデルでは「{lp_pos['feature']}」が最も強い正相関（r={lp_pos['correlation']:.3f}, n={lp_pos['n']}）でした。"
        )
    if mr_neg is not None:
        lines.append(
            f"MRモデルでは「{mr_neg['feature']}」が最も強い負相関（r={mr_neg['correlation']:.3f}, n={mr_neg['n']}）でした。"
        )
    if mr_pos is not None:
        lines.append(
            f"MRモデルでは「{mr_pos['feature']}」が最も強い正相関（r={mr_pos['correlation']:.3f}, n={mr_pos['n']}）でした。"
        )

    lines.append("この結果は探索分析であり、次段階では検証用データで符号・効果量の再現確認を行う前提で解釈してください。")
    return "\n".join(lines)


def show_factor_section(title: str, results: list[dict[str, Any]], top_n: int) -> None:
    st.subheader(title)
    pos_df, neg_df = top_positive_negative(results, top_n)

    left, right = st.columns(2)
    with left:
        plot_factor_bar(pos_df, f"{title} 正相関 上位")
        st.dataframe(pos_df, use_container_width=True, hide_index=True)
    with right:
        plot_factor_bar(neg_df, f"{title} 負相関 上位")
        st.dataframe(neg_df, use_container_width=True, hide_index=True)


def main() -> None:
    st.set_page_config(page_title="SF6 分析ダッシュボード", layout="wide")
    st.title("SF6 基礎力分析ダッシュボード")
    st.caption("サンプル数・要因分析の根拠・総括文を1画面で確認するビュー")

    with st.sidebar:
        st.header("設定")
        input_path = st.text_input("入力CSV", value=DEFAULT_INPUT)
        top_n = st.slider("上位表示件数", min_value=3, max_value=12, value=8)
        exclude_char_dependent = st.checkbox("キャラ依存の強い指標を除外", value=True)
        exclude_play_volume = st.checkbox("プレイ量系指標を除外", value=False)
        run = st.button("再計算", type="primary", use_container_width=True)

    if not run:
        st.info("左の「再計算」を押すと最新CSVで集計します。")

    try:
        columns, rows = load_source(input_path)
    except Exception as exc:
        st.error(str(exc))
        return

    if not rows:
        st.warning("データ行がありません。")
        return

    df = rows_to_dataframe(columns, rows)
    total, sub_n, master_n = summarize_counts(df)
    band_df = band_count_df(df)

    features = ap.get_feature_names(columns)
    features = filter_features(features, exclude_char_dependent, exclude_play_volume)

    sub_rows = [r for r in rows if ap.is_submaster_row(r)]
    master_rows = [r for r in rows if ap.is_master_row(r)]
    progress_rows = ap.build_progress_score(rows)

    lp_results = ap.analyze_segment(sub_rows, features, "リーグポイント")
    mr_results = ap.analyze_segment(master_rows, features, "MR")
    progress_results = ap.analyze_segment(progress_rows, features, "進捗スコア")

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("総件数", f"{total} 件")
    kpi2.metric("LPモデル件数", f"{sub_n} 件")
    kpi3.metric("MRモデル件数", f"{master_n} 件")

    st.subheader("LP帯サンプル数")
    fig_band = px.bar(
        band_df,
        x="表示LP帯",
        y="件数",
        text="件数",
        color="件数",
        color_continuous_scale="Teal",
    )
    fig_band.add_hline(y=60, line_dash="dash", line_color="orange", annotation_text="最低ライン 60")
    fig_band.add_hline(y=80, line_dash="dot", line_color="green", annotation_text="目標 80")
    fig_band.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
    st.plotly_chart(fig_band, use_container_width=True)
    st.dataframe(band_df, use_container_width=True, hide_index=True)

    show_factor_section("LP要因分析", lp_results, top_n)
    show_factor_section("MR要因分析", mr_results, top_n)
    show_factor_section("進捗スコア要因分析", progress_results, top_n)

    summary_text = make_summary_text(total, sub_n, master_n, band_df, lp_results, mr_results)
    st.subheader("総括文（ドラフト）")
    st.text_area("分析サマリー", value=summary_text, height=220)


if __name__ == "__main__":
    main()

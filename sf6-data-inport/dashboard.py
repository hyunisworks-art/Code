from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st

import analyze_playlog as ap
import collect_playlog as cp
import playlog as pl
import scrape_profiles as sp

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

MASTER_RANK_SET = {"MASTER", "HIGH", "GRAND", "ULTIMATE"}

ACTION_HINTS = {
    "ドライブインパクト_受けた回数": "被DIを減らす意識が最優先です。相手DI間合いで大きい技の空振りを減らし、DI警戒のガード選択を増やしてください。",
    "パニッシュカウンターを受けた回数": "暴れや置き技のタイミングを見直しましょう。反撃確定を受けやすい場面をリプレイで特定するのが効果的です。",
    "投げ_受けた回数": "近距離防御で投げを通され過ぎています。遅らせグラップとバックジャンプの使い分けを練習してください。",
    "投げ抜け回数": "投げ抜け精度を上げる余地があります。近距離の防御シチュエーションをトレモで反復しましょう。",
    "相手に追い詰められている時間": "画面端で不利時間が長めです。中央維持を意識した後退ルートと切り返しを優先して改善してください。",
    "相手を追い詰めている時間": "攻め継続力を伸ばせます。端到達後の攻め継続パターンを2〜3本に絞って反復するのが有効です。",
    "キャンセルドライブラッシュ": "キャンセルDRの活用余地があります。主力通常技ヒット確認からのDRを重点練習してください。",
    "ジャストパリィ回数": "JP成功率に伸びしろがあります。打撃重ねに対するJPのタイミングをトレモで集中的に合わせましょう。",
}

FEATURE_DISPLAY_LABELS = {
    "ドライブパリィ": "ドライブゲージ / ドライブパリィ",
    "ドライブインパクト": "ドライブゲージ / ドライブインパクト",
    "オーバードライブアーツ": "ドライブゲージ / オーバードライブアーツ",
    "パリィドライブラッシュ": "ドライブゲージ / パリィドライブラッシュ",
    "キャンセルドライブラッシュ": "ドライブゲージ / キャンセルドライブラッシュ",
    "ドライブリバーサル": "ドライブゲージ / ドライブリバーサル",
    "ダメージ": "ドライブゲージ / ダメージ",
    "使用回数": "ドライブリバーサル / 使用回数",
    "成功回数": "ドライブパリィ / 成功回数",
    "相手のドライブパリィを投げた": "ドライブパリィ / 相手のドライブパリィを投げた",
    "自分のドライブパリィを投げられた": "ドライブパリィ / 自分のドライブパリィを投げられた",
    "ジャストパリィ回数": "ドライブパリィ / ジャストパリィ回数",
    "ドライブインパクト_決めた回数": "ドライブインパクト / 決めた回数",
    "パニッシュカウンターを決めた回数": "ドライブインパクト / パニッシュカウンターを決めた回数",
    "相手のドライブインパクトに決めた回数": "ドライブインパクト / 相手のドライブインパクトに決めた回数",
    "ドライブインパクト_受けた回数": "ドライブインパクト / 受けた回数",
    "パニッシュカウンターを受けた回数": "ドライブインパクト / パニッシュカウンターを受けた回数",
    "相手にドライブインパクトで返された回数": "ドライブインパクト / 相手にドライブインパクトで返された回数",
    "Lv1": "SAゲージ使用割合 / Lv1",
    "Lv2": "SAゲージ使用割合 / Lv2",
    "Lv3": "SAゲージ使用割合 / Lv3",
    "CA": "SAゲージ使用割合 / CA",
    "スタンさせた回数": "スタン / スタンさせた回数",
    "スタンさせられた回数": "スタン / スタンさせられた回数",
    "投げ_決めた回数": "投げ / 決めた回数",
    "投げ_受けた回数": "投げ / 受けた回数",
    "投げ抜け回数": "投げ / 投げ抜け回数",
    "相手を追い詰めている時間": "壁際 / 相手を追い詰めている時間",
    "相手に追い詰められている時間": "壁際 / 相手に追い詰められている時間",
    "ランクマッチプレイ回数": "プレイ回数 / ランクマッチプレイ回数",
    "カジュアルマッチプレイ回数": "プレイ回数 / カジュアルマッチプレイ回数",
    "ルームマッチプレイ回数": "プレイ回数 / ルームマッチプレイ回数",
    "バトルハブマッチプレイ回数": "プレイ回数 / バトルハブマッチプレイ回数",
    "累計プレイポイント": "プレイポイント / 累計プレイポイント",
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


def to_num(value: Any) -> float | None:
    return ap.parse_numeric(str(value))


def feature_label(feature: str) -> str:
    return FEATURE_DISPLAY_LABELS.get(feature, feature)


def add_display_feature(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    out["feature_display"] = out["feature"].astype(str).apply(feature_label)
    return out


def build_rank_options(df: pd.DataFrame) -> list[str]:
    rank_df = df[["ランク", "リーグポイント"]].dropna(subset=["ランク"]).copy()
    rank_df["リーグポイント"] = pd.to_numeric(rank_df["リーグポイント"], errors="coerce")
    grouped = (
        rank_df.groupby("ランク", dropna=False)["リーグポイント"]
        .median()
        .reset_index(name="lp_median")
        .sort_values(["lp_median", "ランク"])
    )
    return grouped["ランク"].astype(str).tolist()


def build_player_row_from_short_id(short_id: str, cookie_text: str, timeout: int, columns: list[str]) -> dict[str, Any]:
    page_props = sp.fetch_profile(short_id, cookie_text, timeout)
    profile_row = sp.extract_profile_row(short_id, "", page_props)
    stats = page_props.get("play", {}).get("battle_stats", {})

    lp = str(profile_row.get("league_point", "") or "")
    mr = str(profile_row.get("master_rating", "") or "0")
    player_name = str(profile_row.get("player_name", "") or "")
    fetch_date = str(profile_row.get("fetch_date", "") or pl.get_today_date())

    lp_master = pl.load_lp_master(pl.get_default_lp_master_path())
    mr_master = pl.load_mr_master(pl.get_default_mr_master_path())
    rank = cp._resolve_rank(lp, mr, lp_master, mr_master)
    raw_row = cp._build_new_row(0, fetch_date, player_name, lp, mr, rank, stats)

    parsed_row: dict[str, Any] = {}
    for idx, col in enumerate(columns):
        parsed_row[col] = raw_row[idx] if idx < len(raw_row) else ""

    parsed_row["No"] = ap.parse_numeric(str(parsed_row.get("No", "")))
    parsed_row["リーグポイント"] = ap.parse_numeric(str(parsed_row.get("リーグポイント", "")))
    parsed_row["MR"] = ap.parse_numeric(str(parsed_row.get("MR", "")))
    parsed_row["_short_id"] = short_id
    return parsed_row


def pick_target_population(df: pd.DataFrame, target_rank: str) -> tuple[pd.DataFrame, str]:
    target_rows = df[df["ランク"] == target_rank].copy()
    if len(target_rows) >= 20:
        return target_rows, f"目標ランク {target_rank} の {len(target_rows)} 件を基準に比較"

    if target_rows.empty:
        return target_rows, "目標ランクのサンプルが見つかりません"

    target_lp = pd.to_numeric(target_rows["リーグポイント"], errors="coerce").median()
    band = lp_band(target_lp) if pd.notna(target_lp) else "不明"
    fallback = df[df["LP帯"] == band].copy() if band not in {"不明", "範囲外"} else target_rows
    return fallback, f"目標ランクの件数が少ないため、LP帯 {band}（{len(fallback)} 件）を基準に比較"


def compute_feature_gap_table(
    player_row: dict[str, Any],
    target_df: pd.DataFrame,
    model_results: list[dict[str, Any]],
    candidate_features: list[str],
) -> pd.DataFrame:
    corr_map = {r["feature"]: float(r["correlation"]) for r in model_results}
    rows: list[dict[str, Any]] = []

    for feature in candidate_features:
        corr = corr_map.get(feature)
        if corr is None or abs(corr) < 0.12:
            continue

        player_value = to_num(player_row.get(feature, ""))
        if player_value is None:
            continue

        series = pd.to_numeric(target_df[feature], errors="coerce").dropna()
        if len(series) < 15:
            continue

        target_median = float(series.median())
        target_std = float(series.std(ddof=1)) if len(series) >= 2 else 0.0
        if target_std <= 0:
            target_std = 1.0

        direction = 1.0 if corr > 0 else -1.0
        # gap_z > 0 は「目標基準に対して不足」
        gap_z = direction * (target_median - float(player_value)) / target_std

        rows.append(
            {
                "feature": feature,
                "player": float(player_value),
                "target_median": target_median,
                "correlation": corr,
                "gap_z": gap_z,
                "n_target": int(len(series)),
            }
        )

    if not rows:
        return pd.DataFrame(columns=["feature", "player", "target_median", "correlation", "gap_z", "n_target"])
    return pd.DataFrame(rows).sort_values("gap_z", ascending=False)


def build_action_text(feature: str, is_shortage: bool) -> str:
    if feature in ACTION_HINTS:
        return ACTION_HINTS[feature]
    if is_shortage:
        return "目標帯中央値との差があるため、該当シーンの再現練習を増やしてください。"
    return "現状は目標帯基準を満たしています。再現性維持を優先してください。"


def build_play_volume_advice(player_row: dict[str, Any], target_df: pd.DataFrame) -> list[str]:
    advices: list[str] = []

    rank_match_player = to_num(player_row.get("ランクマッチプレイ回数", ""))
    casual_player = to_num(player_row.get("カジュアルマッチプレイ回数", ""))

    rank_match_target = pd.to_numeric(target_df["ランクマッチプレイ回数"], errors="coerce").dropna()
    if not rank_match_target.empty and rank_match_player is not None:
        target_median = float(rank_match_target.median())
        if rank_match_player < target_median * 0.6:
            advices.append("実戦量が目標帯に対して少なめです。まずはランクマッチの試行回数を増やし、判断の母数を作りましょう。")
        elif rank_match_player > target_median * 1.4:
            advices.append("実戦量は十分です。次はトレーニングモードで不足指標に直結する反復練習の比率を増やすと効率的です。")
        else:
            advices.append("実戦量は目標帯と同程度です。実戦とトレモを半々で回し、課題指標を狙って改善しましょう。")

    if casual_player is not None and rank_match_player is not None and casual_player > rank_match_player * 1.5:
        advices.append("カジュアル比率が高めです。ランク到達を目的にする期間は、ランクマッチ比率を上げると昇格速度が安定します。")

    if not advices:
        advices.append("プレイ時間系データは取得できました。課題指標の改善状況と合わせて実戦/トレモ配分を調整してください。")
    return advices


def show_personal_coaching_section(
    df: pd.DataFrame,
    columns: list[str],
    features: list[str],
    lp_results: list[dict[str, Any]],
    mr_results: list[dict[str, Any]],
) -> None:
    st.header("個別トレーニング提案（ユーザーコード指定）")
    st.caption("傾向分析用CSVに本人がいなくても、ユーザーコードから最新データを取得して診断します")

    rank_options = build_rank_options(df)
    default_rank_idx = len(rank_options) - 1 if rank_options else 0

    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        short_id = st.text_input("ユーザーコード (short_id)", value="")
    with c2:
        target_rank = st.selectbox("目標ランク", options=rank_options, index=default_rank_idx if rank_options else None)
    with c3:
        timeout = int(st.number_input("timeout(秒)", min_value=10, max_value=60, value=30, step=5))

    cookie_input = st.text_input("Cookie（空なら .buckler_cookie.txt を使用）", type="password")
    run_diag = st.button("個別診断を実行", use_container_width=True, type="primary")

    if not run_diag:
        return

    if not short_id.strip():
        st.warning("ユーザーコードを入力してください。")
        return

    try:
        cookie = sp.load_cookie_text(cookie_input, sp.DEFAULT_COOKIE_FILE)
        sp.validate_cookie_text(cookie)
        if not cookie:
            st.error("Cookie が見つかりません。.buckler_cookie.txt か Cookie 入力欄を確認してください。")
            return

        player_row = build_player_row_from_short_id(short_id.strip(), cookie, timeout, columns)
    except Exception as exc:
        st.error(f"ユーザーデータ取得に失敗しました: {exc}")
        return

    target_df, target_note = pick_target_population(df, str(target_rank))
    if target_df.empty:
        st.error("目標ランクの基準サンプルが見つからないため診断できません。")
        return

    target_lp_median = pd.to_numeric(target_df["リーグポイント"], errors="coerce").median()
    use_mr_model = pd.notna(target_lp_median) and float(target_lp_median) >= MASTER_LP_THRESHOLD
    model_results = mr_results if use_mr_model else lp_results

    coaching_features = [
        f
        for f in features
        if f not in CHAR_DEPENDENT_FEATURES and f not in PLAY_VOLUME_FEATURES and f not in {"No", "データ取得日", "プレイヤー名", "リーグポイント", "ランク", "MR"}
    ]
    gap_df = compute_feature_gap_table(player_row, target_df, model_results, coaching_features)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("プレイヤー名", str(player_row.get("プレイヤー名", "-")))
    k2.metric("現在ランク", str(player_row.get("ランク", "-")))
    k3.metric("LP", f"{int(player_row['リーグポイント']) if isinstance(player_row.get('リーグポイント'), (int, float)) else '-'}")
    mr_val = player_row.get("MR")
    k4.metric("MR", f"{int(mr_val) if isinstance(mr_val, (int, float)) else '-'}")

    st.caption(f"比較基準: {target_note}")

    if gap_df.empty:
        st.info("比較可能な指標が不足しているため、個別課題を算出できませんでした。")
        return

    shortage_df = gap_df[gap_df["gap_z"] > 0].head(5).copy()
    strength_df = gap_df[gap_df["gap_z"] < 0].sort_values("gap_z").head(3).copy()

    if not shortage_df.empty:
        shortage_df["不足度(z)" ] = shortage_df["gap_z"].round(2)
        shortage_df["指標（区分/項目）"] = shortage_df["feature"].astype(str).apply(feature_label)
        shortage_df["アクション"] = shortage_df["feature"].apply(lambda f: build_action_text(str(f), True))
        st.subheader("不足上位（優先課題）")
        st.dataframe(
            shortage_df[["指標（区分/項目）", "player", "target_median", "correlation", "不足度(z)", "アクション"]],
            use_container_width=True,
            hide_index=True,
        )

    if not strength_df.empty:
        strength_df["優位度(z)"] = (-strength_df["gap_z"]).round(2)
        strength_df["指標（区分/項目）"] = strength_df["feature"].astype(str).apply(feature_label)
        strength_df["評価"] = "目標帯中央値以上。維持しつつ再現性を高める"
        st.subheader("良い部分（強み）")
        st.dataframe(
            strength_df[["指標（区分/項目）", "player", "target_median", "correlation", "優位度(z)", "評価"]],
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("プレイ時間系アドバイス（別軸）")
    for text in build_play_volume_advice(player_row, target_df):
        st.write(f"- {text}")


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
        y="feature_display",
        orientation="h",
        color="correlation",
        color_continuous_scale="RdBu",
        title=title,
        hover_data={
            "n": True,
            "r_squared": ":.3f",
            "correlation": ":.3f",
            "feature_display": False,
            "feature": True,
        },
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
            f"LPモデルでは「{feature_label(lp_neg['feature'])}」が最も強い負相関（r={lp_neg['correlation']:.3f}, n={lp_neg['n']}）を示しました。"
        )
    if lp_pos is not None:
        lines.append(
            f"LPモデルでは「{feature_label(lp_pos['feature'])}」が最も強い正相関（r={lp_pos['correlation']:.3f}, n={lp_pos['n']}）でした。"
        )
    if mr_neg is not None:
        lines.append(
            f"MRモデルでは「{feature_label(mr_neg['feature'])}」が最も強い負相関（r={mr_neg['correlation']:.3f}, n={mr_neg['n']}）でした。"
        )
    if mr_pos is not None:
        lines.append(
            f"MRモデルでは「{feature_label(mr_pos['feature'])}」が最も強い正相関（r={mr_pos['correlation']:.3f}, n={mr_pos['n']}）でした。"
        )

    lines.append("この結果は探索分析であり、次段階では検証用データで符号・効果量の再現確認を行う前提で解釈してください。")
    return "\n".join(lines)


def show_factor_section(title: str, results: list[dict[str, Any]], top_n: int) -> None:
    st.subheader(title)
    pos_df, neg_df = top_positive_negative(results, top_n)
    pos_df = add_display_feature(pos_df)
    neg_df = add_display_feature(neg_df)

    left, right = st.columns(2)
    with left:
        plot_factor_bar(pos_df, f"{title} 正相関 上位")
        st.dataframe(
            pos_df[["feature_display", "correlation", "n", "r_squared"]],
            use_container_width=True,
            hide_index=True,
        )
    with right:
        plot_factor_bar(neg_df, f"{title} 負相関 上位")
        st.dataframe(
            neg_df[["feature_display", "correlation", "n", "r_squared"]],
            use_container_width=True,
            hide_index=True,
        )


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

    st.divider()
    show_personal_coaching_section(df, columns, features, lp_results, mr_results)


if __name__ == "__main__":
    main()

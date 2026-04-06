from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st
from dotenv import load_dotenv
from supabase import create_client

import analyze_playlog as ap
import collect_playlog as cp
import playlog as pl
import scrape_profiles as sp

load_dotenv()
_SUPABASE_URL = os.getenv("SUPABASE_URL", "")
_SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

MASTER_LP_THRESHOLD = 25000
RANKING_OUTPUT_DIR = Path(__file__).parent / "ranking-output"
_MY_DATA_DIR = Path(__file__).parent / "data" / "my"

# Priority 2: 信頼度・改善閾値定数
CONFIDENCE_HIGH_N = 60     # n >= 60 → 高信頼
CONFIDENCE_MED_N = 30      # n >= 30 → 中信頼
IMPROVEMENT_DELTA = 0.20   # shortage_score が前回比 0.20 以上減少 → 「改善」
STABLE_DELTA = 0.05        # 変化が 0.05 未満 → 「安定」

CHAR_DEPENDENT_FEATURES = {
    "Lv1",
    "Lv2",
    "Lv3",
    "CA",
}

PLAY_VOLUME_FEATURES = {
    "ランクマッチプレイ回数",
    "カジュアルマッチプレイ回数",
    "ルームマッチプレイ回数",
    "バトルハブマッチプレイ回数",
    "累計プレイポイント",
}

# Supabase JSONB → 行辞書変換用フィールドマッピング（analyze_step1.py と同じ定義）
_PCT_FIELDS: list[tuple[str, str]] = [
    ("gauge_rate_drive_guard",            "ドライブパリィ"),
    ("gauge_rate_drive_impact",           "ドライブインパクト"),
    ("gauge_rate_drive_arts",             "オーバードライブアーツ"),
    ("gauge_rate_drive_rush_from_parry",  "パリィドライブラッシュ"),
    ("gauge_rate_drive_rush_from_cancel", "キャンセルドライブラッシュ"),
    ("gauge_rate_drive_reversal",         "ドライブリバーサル"),
    ("gauge_rate_drive_other",            "ダメージ"),
    ("gauge_rate_sa_lv1",                 "Lv1"),
    ("gauge_rate_sa_lv2",                 "Lv2"),
    ("gauge_rate_sa_lv3",                 "Lv3"),
    ("gauge_rate_ca",                     "CA"),
]
_NUM_FIELDS: list[tuple[str, str]] = [
    ("drive_reversal",                        "使用回数"),
    ("drive_parry",                           "成功回数"),
    ("throw_drive_parry",                     "相手のドライブパリィを投げた"),
    ("received_throw_drive_parry",            "自分のドライブパリィを投げられた"),
    ("just_parry",                            "ジャストパリィ回数"),
    ("drive_impact",                          "ドライブインパクト_決めた回数"),
    ("punish_counter",                        "パニッシュカウンターを決めた回数"),
    ("drive_impact_to_drive_impact",          "相手のドライブインパクトに決めた回数"),
    ("received_drive_impact",                 "ドライブインパクト_受けた回数"),
    ("received_punish_counter",               "パニッシュカウンターを受けた回数"),
    ("received_drive_impact_to_drive_impact", "相手にドライブインパクトで返された回数"),
    ("stun",                                  "スタンさせた回数"),
    ("received_stun",                         "スタンさせられた回数"),
    ("throw_count",                           "投げ_決めた回数"),
    ("received_throw_count",                  "投げ_受けた回数"),
    ("throw_tech",                            "投げ抜け回数"),
    ("corner_time",                           "相手を追い詰めている時間"),
    ("cornered_time",                         "相手に追い詰められている時間"),
    ("rank_match_play_count",                 "ランクマッチプレイ回数"),
    ("casual_match_play_count",               "カジュアルマッチプレイ回数"),
    ("custom_room_match_play_count",          "ルームマッチプレイ回数"),
    ("battle_hub_match_play_count",           "バトルハブマッチプレイ回数"),
    ("total_all_character_play_point",        "累計プレイポイント"),
]
_SUPABASE_COLUMNS: list[str] = (
    ["No", "データ取得日", "プレイヤー名", "リーグポイント", "ランク", "MR"]
    + [col for _, col in _PCT_FIELDS]
    + [col for _, col in _NUM_FIELDS]
)

MASTER_RANK_SET = {"MASTER", "HIGH", "GRAND", "ULTIMATE"}

MASTER_RANK_ORDER = {"MASTER": 0, "HIGH": 1, "GRAND": 2, "ULTIMATE": 3}
MASTER_RANK_LABELS = {
    "MASTER": "MASTER",
    "HIGH": "HIGH MASTER",
    "GRAND": "GRAND MASTER",
    "ULTIMATE": "ULTIMATE MASTER",
}

DRIVE_GAUGE_COLS = [
    "ドライブパリィ",
    "ドライブインパクト",
    "オーバードライブアーツ",
    "パリィドライブラッシュ",
    "キャンセルドライブラッシュ",
    "ドライブリバーサル",
    "ダメージ",
]

SA_GAUGE_COLS = ["Lv1", "Lv2", "Lv3", "CA"]

# 公式 Buckler's Boot Camp の配色に合わせた色定義
DRIVE_GAUGE_COLORS: dict[str, str] = {
    "ドライブパリィ":         "#F48FB1",  # ライトピンク
    "ドライブインパクト":     "#E91E8C",  # ホットピンク
    "オーバードライブアーツ": "#AD1457",  # ダークマゼンタ
    "パリィドライブラッシュ": "#1565C0",  # ダークブルー
    "キャンセルドライブラッシュ": "#90CAF9",  # ライトブルー
    "ドライブリバーサル":     "#4CAF50",  # グリーン
    "ダメージ":               "#9E9E9E",  # グレー
}

SA_GAUGE_COLORS: dict[str, str] = {
    "Lv1": "#F8BBD9",  # 淡いピンク
    "Lv2": "#F48FB1",  # ライトピンク
    "Lv3": "#AD1457",  # ダークマゼンタ
    "CA":  "#1565C0",  # ブルー
}

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

# 8バンド構成：ROOKIE / IRON / BRONZE / SILVER / GOLD / PLATINUM / DIAMOND / MASTER以上
LP_BANDS: list[tuple[int, int, str]] = [
    (0,     999,   "0k-1k"),
    (1000,  2999,  "1k-3k"),
    (3000,  4999,  "3k-5k"),
    (5000,  8999,  "5k-9k"),
    (9000,  12999, "9k-13k"),
    (13000, 18999, "13k-19k"),
    (19000, 24999, "19k-25k"),
    (25000, 10**9, "25k-"),
]

LP_BAND_RANK_HINT = {
    "0k-1k":   "ROOKIE1-5",
    "1k-3k":   "IRON1-5",
    "3k-5k":   "BRONZE1-5",
    "5k-9k":   "SILVER1-5",
    "9k-13k":  "GOLD1-5",
    "13k-19k": "PLATINUM1-5",
    "19k-25k": "DIAMOND1-5",
    "25k-":    "MASTER以上",
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


# rankからLP中央値を推定するマッピング
_RANK_LP_MIDPOINT: dict[str, int] = {
    "rookie": 500,
    "iron": 2000,
    "bronze": 4000,
    "silver": 7000,
    "gold": 11000,
    "platinum": 16000,
    "diamond": 22000,
    "master": 25000,
}


def _supabase_row_to_dict(row: dict[str, Any]) -> dict[str, Any]:
    """Supabase player_data行をdashboardが扱える行辞書に変換する。"""
    li = row.get("league_info") or {}
    bs = (row.get("play") or {}).get("battle_stats") or {}

    rank_name: str = (
        (li.get("league_rank_info") or {}).get("league_rank_name")
        or row.get("rank")
        or ""
    )

    # LP値がなければrankから推定
    lp = li.get("league_point")
    if lp is None:
        rank_key = (row.get("rank") or "").lower()
        # master_high, master_grand 等もmasterとして扱う
        if rank_key.startswith("master"):
            rank_key = "master"
        lp = _RANK_LP_MIDPOINT.get(rank_key)

    # MR: league_infoにあればそれを使い、なければcharacter_league_infosの最大値で補完
    mr: float | None = li.get("master_rating") or None
    if mr is None:
        cli = (row.get("play") or {}).get("character_league_infos") or []
        max_mr = max(
            ((c.get("league_info") or {}).get("master_rating") or 0) for c in cli
        ) if cli else 0
        if max_mr > 0:
            mr = float(max_mr)

    result: dict[str, Any] = {
        "No": None,
        "データ取得日": row.get("fetch_date", ""),
        "プレイヤー名": str(row.get("player_id", "")),
        "リーグポイント": lp,
        "ランク": rank_name,
        "MR": mr,
    }

    for json_key, col in _PCT_FIELDS:
        val = bs.get(json_key)
        result[col] = float(val) * 100 if val is not None else None

    for json_key, col in _NUM_FIELDS:
        val = bs.get(json_key)
        result[col] = float(val) if val is not None else None

    return result


@st.cache_data(show_spinner=False)
def load_from_supabase() -> tuple[list[str], list[dict[str, Any]]]:
    """Supabaseからdata_type=sampleのデータをページネーションで取得する。"""
    if not _SUPABASE_URL or not _SUPABASE_KEY:
        raise RuntimeError(".env に SUPABASE_URL / SUPABASE_KEY を設定してください")
    client = create_client(_SUPABASE_URL, _SUPABASE_KEY)
    rows_raw: list[dict[str, Any]] = []
    page_size = 50
    offset = 0
    while True:
        response = (
            client.table("player_data")
            .select("player_id,fetch_date,rank,data_type,league_info,play")
            .eq("data_type", "sample")
            .range(offset, offset + page_size - 1)
            .execute()
        )
        batch = response.data or []
        rows_raw.extend(batch)
        if len(batch) < page_size:
            break
        offset += page_size
    rows = [_supabase_row_to_dict(r) for r in rows_raw]
    return _SUPABASE_COLUMNS, rows


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
        # % 付き文字列を含む列を正しくパースするため parse_numeric を使う
        df[col] = df[col].apply(
            lambda x: x if isinstance(x, float) else ap.parse_numeric(str(x))
        )

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
        lambda x: f"{x} ({LP_BAND_RANK_HINT.get(x, '-')})"
    )
    counts["目標80との差"] = 80 - counts["件数"]
    counts["最低60との差"] = 60 - counts["件数"]
    counts["LP帯ソート"] = counts["LP帯"].apply(sort_key_for_band)
    counts = counts.sort_values("LP帯ソート").drop(columns=["LP帯ソート"])
    return counts


def mr_band_count_df(df: pd.DataFrame) -> pd.DataFrame:
    """MRモデル（MASTER以上）のランク別サンプル数を返す。"""
    master_df = df[df["モデル区分"] == "MRモデル"].copy()
    master_df["ランク"] = master_df["ランク"].astype(str).str.strip().str.upper()
    counts = (
        master_df["ランク"]
        .value_counts(dropna=True)
        .rename_axis("ランク")
        .reset_index(name="件数")
    )
    # MR値でサブ段階を分類（データがあれば）
    if "MR" in master_df.columns and master_df["MR"].notna().any():
        mr_bands = []
        for _, row in master_df.iterrows():
            mr = row.get("MR")
            if mr is not None and mr >= 1800:
                mr_bands.append("ULTIMATE")
            elif mr is not None and mr >= 1700:
                mr_bands.append("GRAND")
            elif mr is not None and mr >= 1600:
                mr_bands.append("HIGH")
            else:
                mr_bands.append("MASTER")
        master_df = master_df.copy()
        master_df["MRランク"] = mr_bands
        counts = (
            master_df["MRランク"]
            .value_counts(dropna=True)
            .rename_axis("ランク")
            .reset_index(name="件数")
        )
    counts = counts[counts["ランク"].isin(MASTER_RANK_SET)].copy()
    if counts.empty:
        return counts
    counts["ランク順"] = counts["ランク"].map(MASTER_RANK_ORDER).fillna(99)
    counts["表示ランク"] = counts["ランク"].map(MASTER_RANK_LABELS).fillna(counts["ランク"])
    counts = counts.sort_values("ランク順").drop(columns=["ランク順"])
    return counts


def lookup_short_id_by_name(player_name: str) -> tuple[str, str] | None:
    """ranking-output/ の CSV からプレイヤー名で short_id を検索する。
    見つかった場合は (short_id, player_name) を返す。見つからない場合は None。
    複数ヒットした場合は最初のものを返す。
    """
    if not RANKING_OUTPUT_DIR.exists():
        return None

    needle = player_name.strip().lower()
    for csv_path in sorted(RANKING_OUTPUT_DIR.glob("*.csv")):
        try:
            pairs = sp.load_short_ids_from_csv(csv_path)
        except Exception:
            continue
        for sid, name in pairs:
            if name.strip().lower() == needle:
                return sid, name
    return None


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
    # fighter_banner_info.personal_info.fighter_id が正しいプレイヤー名のパス
    player_name = (
        page_props.get("fighter_banner_info", {})
        .get("personal_info", {})
        .get("fighter_id", "")
    ) or str(profile_row.get("player_name", "") or "")
    fetch_date = str(profile_row.get("fetch_date", "") or pl.get_today_date())

    lp_master = pl.load_lp_master(pl.get_default_lp_master_path())
    mr_master = pl.load_mr_master(pl.get_default_mr_master_path())
    rank = cp._resolve_rank(lp, mr, lp_master, mr_master)

    lp_num = ap.parse_numeric(lp)
    mr_num = ap.parse_numeric(mr) if mr and mr != "0" else None
    # MR補完: master_ratingがなければcharacter_league_infosの最大値で補完
    if not mr_num:
        cli = page_props.get("play", {}).get("character_league_infos") or []
        max_mr = max(
            ((c.get("league_info") or {}).get("master_rating") or 0) for c in cli
        ) if cli else 0
        if max_mr > 0:
            mr_num = float(max_mr)

    # _supabase_row_to_dict() と同じフィールドマッピングで変換（スケール統一）
    parsed_row: dict[str, Any] = {
        "No": None,
        "データ取得日": fetch_date,
        "プレイヤー名": player_name,
        "リーグポイント": lp_num,
        "ランク": rank,
        "MR": mr_num,
    }
    for json_key, col in _PCT_FIELDS:
        val = stats.get(json_key)
        parsed_row[col] = float(val) * 100 if val is not None else None
    for json_key, col in _NUM_FIELDS:
        val = stats.get(json_key)
        parsed_row[col] = float(val) if val is not None else None

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
        gap_z = direction * (target_median - float(player_value)) / target_std
        shortage_score = gap_z * abs(corr)

        # 上位何%：「上位が低いほど優秀」に統一
        # corr > 0（多いほど良い）: 自分より低い人の割合 = 下位%、上位% = 100 - 下位%
        # corr < 0（少ないほど良い）: 自分より低い人の割合 = そのまま上位%
        pct_below = float((series <= float(player_value)).mean() * 100)
        upper_pct = round(100 - pct_below if corr > 0 else pct_below, 1)

        rows.append(
            {
                "feature": feature,
                "player": float(player_value),
                "target_median": target_median,
                "correlation": corr,
                "gap_z": gap_z,
                "shortage_score": shortage_score,
                "upper_pct": upper_pct,
                "n_target": int(len(series)),
            }
        )

    if not rows:
        return pd.DataFrame(columns=["feature", "player", "target_median", "correlation", "gap_z", "shortage_score", "upper_pct", "n_target"])
    return pd.DataFrame(rows).sort_values("shortage_score", ascending=False)


def confidence_label(n: int) -> str:
    """サンプル数から信頼度ラベルを返す。"""
    if n >= CONFIDENCE_HIGH_N:
        return "高"
    elif n >= CONFIDENCE_MED_N:
        return "中"
    else:
        return "⚠低"


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


def _show_gap_bar_chart(gap_df: pd.DataFrame, title: str) -> None:
    """不足（負）と強み（正）を横棒グラフで表示する。"""
    chart_df = gap_df.copy()
    chart_df["指標"] = chart_df["feature"].astype(str).apply(feature_label)
    # gap_z を反転: 正=強み、負=不足 として表示
    chart_df["差分スコア"] = -chart_df["gap_z"]
    chart_df = chart_df.sort_values("差分スコア", ascending=True)
    chart_df["判定"] = chart_df["差分スコア"].apply(lambda x: "強み" if x >= 0 else "不足")

    fig = px.bar(
        chart_df,
        x="差分スコア",
        y="指標",
        orientation="h",
        color="判定",
        color_discrete_map={"強み": "#4A90C4", "不足": "#E8734A"},
        title=title,
    )
    fig.add_vline(x=0, line_dash="dash", line_color="gray", opacity=0.5)
    fig.update_layout(height=max(300, len(chart_df) * 28), margin=dict(l=10, r=10, t=50, b=10), showlegend=True)
    st.plotly_chart(fig, use_container_width=True)


def _show_gap_table(gap_df: pd.DataFrame, top_n: int = 5) -> None:
    """不足上位・強み上位のマトリックステーブルを表示する。"""
    if gap_df.empty:
        return

    display_df = gap_df.copy()
    display_df["指標"] = display_df["feature"].astype(str).apply(feature_label)
    display_df["判定"] = display_df["shortage_score"].apply(lambda x: "不足" if x > 0 else "強み")
    display_df["信頼度"] = display_df["n_target"].apply(confidence_label)
    display_df["順位"] = display_df.apply(
        lambda r: f"下位{100 - r['upper_pct']:.1f}%" if r["shortage_score"] > 0
                  else f"上位{r['upper_pct']:.1f}%",
        axis=1,
    )

    shortage = display_df[display_df["shortage_score"] > 0].sort_values("shortage_score", ascending=False).head(top_n)
    strength = display_df[display_df["shortage_score"] < 0].sort_values("shortage_score", ascending=True).head(top_n)

    table_cols = ["指標", "判定", "player", "target_median", "順位", "信頼度"]
    rename_map = {"player": "あなた", "target_median": "基準中央値"}

    # 両テーブル合算で最長ラベルを計算（CJK≈15px・ASCII≈8px・余白24px）
    all_labels = [
        feature_label(str(f))
        for f in pd.concat([strength, shortage])["feature"]
    ]
    def _label_px(s: str) -> int:
        cjk = sum(1 for c in s if ord(c) > 0x2E7F)
        return cjk * 15 + (len(s) - cjk) * 8 + 24
    label_col_width = max((_label_px(lbl) for lbl in all_labels), default=200)

    col_config = {
        "指標": st.column_config.TextColumn(width=label_col_width),
        # 残り列は幅未指定 → use_container_width=True で等分
    }

    def _render(df: pd.DataFrame) -> None:
        t = df[table_cols].rename(columns=rename_map)
        t["あなた"] = t["あなた"].round(2)
        t["基準中央値"] = t["基準中央値"].round(2)
        st.dataframe(t, use_container_width=True, hide_index=True, column_config=col_config)


    st.markdown("**強み上位**")
    if strength.empty:
        st.info("強み指標なし")
    else:
        _render(strength)

    st.markdown("**不足上位**")
    if shortage.empty:
        st.info("不足指標なし")
    else:
        _render(shortage)


def show_play_volume_table(player_row: dict[str, Any], same_rank_df: pd.DataFrame, target_df: pd.DataFrame) -> None:
    """プレイ時間系指標の自分の値と同ランク帯・目標帯のサンプル平均を表形式で表示する。"""
    records = []
    for feature in PLAY_VOLUME_FEATURES:
        player_val = to_num(player_row.get(feature, ""))
        same_series = pd.to_numeric(same_rank_df[feature], errors="coerce").dropna() if not same_rank_df.empty and feature in same_rank_df.columns else pd.Series(dtype=float)
        target_series = pd.to_numeric(target_df[feature], errors="coerce").dropna()
        same_mean = round(float(same_series.mean()), 1) if not same_series.empty else None
        target_mean = round(float(target_series.mean()), 1) if not target_series.empty else None
        records.append({
            "指標": feature_label(feature),
            "あなた": player_val if player_val is not None else "-",
            "サンプル平均（同ランク帯）": same_mean if same_mean is not None else "-",
            "サンプル平均（目標帯）": target_mean if target_mean is not None else "-",
        })
    st.dataframe(pd.DataFrame(records), use_container_width=True, hide_index=True)


def show_gauge_pie_charts(player_row: dict[str, Any], target_df: pd.DataFrame) -> None:
    """ドライブゲージ・SAゲージの使用割合を4枚の円グラフで表示する。"""
    st.caption("※ ゲージ使用割合は使用キャラクターで異なるため参考値としてください")

    def make_pie_df(cols: list[str], source: dict[str, Any] | None, df: pd.DataFrame | None) -> pd.DataFrame:
        if source is not None:
            vals = {c: (to_num(source.get(c, "")) or 0.0) for c in cols}
        else:
            vals = {}
            for c in cols:
                if c not in df.columns:
                    vals[c] = 0.0
                    continue
                series = pd.to_numeric(df[c], errors="coerce").dropna()
                vals[c] = float(series.mean()) if not series.empty else 0.0
        total = sum(vals.values())
        if total <= 0:
            total = 1.0
        return pd.DataFrame({"項目": list(vals.keys()), "割合": [v / total * 100 for v in vals.values()]})

    def pie_chart(title: str, pie_df: pd.DataFrame, color_map: dict[str, str], col_order: list[str]) -> None:
        st.markdown(f"**{title}**")
        fig = px.pie(
            pie_df,
            names="項目",
            values="割合",
            hole=0.35,
            color="項目",
            color_discrete_map=color_map,
            category_orders={"項目": col_order},
        )
        fig.update_traces(textposition="inside", textinfo="percent")
        fig.update_layout(height=380, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig, use_container_width=True)

    # 1行目：あなたのドライブ・SAゲージ
    row1_col1, row1_col2 = st.columns(2)
    with row1_col1:
        pie_chart(
            "あなた：ドライブゲージ使用割合",
            make_pie_df(DRIVE_GAUGE_COLS, player_row, None),
            DRIVE_GAUGE_COLORS,
            DRIVE_GAUGE_COLS,
        )
    with row1_col2:
        pie_chart(
            "あなた：SAゲージ消費割合",
            make_pie_df(SA_GAUGE_COLS, player_row, None),
            SA_GAUGE_COLORS,
            SA_GAUGE_COLS,
        )

    # 2行目：目標ランク帯のドライブ・SAゲージ
    row2_col1, row2_col2 = st.columns(2)
    with row2_col1:
        pie_chart(
            "目標ランク帯：ドライブゲージ使用割合",
            make_pie_df(DRIVE_GAUGE_COLS, None, target_df),
            DRIVE_GAUGE_COLORS,
            DRIVE_GAUGE_COLS,
        )
    with row2_col2:
        pie_chart(
            "目標ランク帯：SAゲージ消費割合",
            make_pie_df(SA_GAUGE_COLS, None, target_df),
            SA_GAUGE_COLORS,
            SA_GAUGE_COLS,
        )


def load_my_history(short_id: str) -> list[dict[str, Any]]:
    """data/my/*_<short_id>.json を日付順に読み、行辞書のリストを返す。"""
    rows: list[dict[str, Any]] = []
    if not _MY_DATA_DIR.exists():
        return rows
    for path in sorted(_MY_DATA_DIR.glob(f"*_{short_id}.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            row = _supabase_row_to_dict(data)
            row["_date"] = data.get("fetch_date", path.stem.split("_")[0])
            rows.append(row)
        except Exception:
            continue
    return rows


def show_weekly_tracking_section(
    short_id: str,
    target_df: pd.DataFrame,
    model_results: list[dict[str, Any]],
    coaching_features: list[str],
) -> None:
    """data/my/ の履歴データで週次トレンドと改善ランキングを表示する。"""
    history = load_my_history(short_id)
    if len(history) < 2:
        st.info(f"週次比較には2件以上の履歴データが必要です（現在 {len(history)} 件）。毎日 fetch_my_data.py を実行してください。")
        return

    # 各日付の gap_df を計算
    dated_gaps: list[tuple[str, pd.DataFrame]] = []
    for row in history:
        date = str(row.get("_date", ""))
        g = compute_feature_gap_table(row, target_df, model_results, coaching_features)
        if not g.empty:
            dated_gaps.append((date, g))

    if len(dated_gaps) < 2:
        st.info("履歴データから課題指標を算出できませんでした。")
        return

    latest_date, latest_gap = dated_gaps[-1]
    prev_date, prev_gap = dated_gaps[-2]

    # ── 改善ランキングTOP3（強みを先に見せる） ──────────────────────────
    st.markdown("#### 今回の変化（強みを先に）")
    delta_records: list[dict[str, Any]] = []
    for _, g_row in latest_gap.iterrows():
        feat = str(g_row["feature"])
        prv = prev_gap[prev_gap["feature"] == feat]
        if prv.empty:
            continue
        cur_score = float(g_row["shortage_score"])
        prv_score = float(prv["shortage_score"].iloc[0])
        delta = prv_score - cur_score  # 正 = 改善（不足が減った）
        delta_records.append({
            "feature": feat,
            "delta": delta,
            "cur_score": cur_score,
            "prv_score": prv_score,
        })

    if delta_records:
        delta_df = pd.DataFrame(delta_records).sort_values("delta", ascending=False)
        # 改善TOP3
        improved = delta_df[delta_df["delta"] >= IMPROVEMENT_DELTA].head(3)
        if not improved.empty:
            st.markdown("**📈 改善 TOP3**")
            recs = []
            for _, r in improved.iterrows():
                recs.append({
                    "指標": feature_label(str(r["feature"])),
                    "改善幅": f"+{r['delta']:.3f}",
                    "今回スコア": round(r["cur_score"], 3),
                    "前回スコア": round(r["prv_score"], 3),
                })
            st.dataframe(pd.DataFrame(recs), use_container_width=True, hide_index=True)
        else:
            st.caption("前回比 0.20 以上の改善指標なし。")

        # 後退 or 注意
        worsened = delta_df[delta_df["delta"] <= -IMPROVEMENT_DELTA].tail(3)
        if not worsened.empty:
            st.markdown("**⚠ 要注意（後退）**")
            recs = []
            for _, r in worsened.iterrows():
                recs.append({
                    "指標": feature_label(str(r["feature"])),
                    "悪化幅": f"{r['delta']:.3f}",
                    "今回スコア": round(r["cur_score"], 3),
                    "前回スコア": round(r["prv_score"], 3),
                })
            st.dataframe(pd.DataFrame(recs), use_container_width=True, hide_index=True)

    # ── 不足上位3指標の推移ライン ────────────────────────────────────────
    top3_features = (
        latest_gap[latest_gap["shortage_score"] > 0]
        .head(3)["feature"]
        .tolist()
    )
    if top3_features:
        trend_records: list[dict[str, Any]] = []
        for date, g in dated_gaps:
            for feat in top3_features:
                feat_row = g[g["feature"] == feat]
                if feat_row.empty:
                    continue
                trend_records.append({
                    "日付": date,
                    "指標": feature_label(feat),
                    "shortage_score": round(float(feat_row["shortage_score"].iloc[0]), 3),
                })
        if trend_records:
            trend_df = pd.DataFrame(trend_records)
            fig = px.line(
                trend_df,
                x="日付",
                y="shortage_score",
                color="指標",
                markers=True,
                title="不足スコア推移（優先課題 上位3指標）",
            )
            fig.add_hline(y=0, line_dash="dash", line_color="gray", annotation_text="目標達成ライン")
            fig.update_layout(height=300, margin=dict(l=10, r=10, t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)

    # ── 週次アクションプラン ─────────────────────────────────────────────
    st.markdown("#### 週次アクションプラン")
    plan_records: list[dict[str, Any]] = []
    for feat in top3_features:
        cur_row = latest_gap[latest_gap["feature"] == feat]
        prv_row = prev_gap[prev_gap["feature"] == feat]
        if cur_row.empty or prv_row.empty:
            continue
        cur_score = float(cur_row["shortage_score"].iloc[0])
        prv_score = float(prv_row["shortage_score"].iloc[0])
        delta = prv_score - cur_score

        if cur_score <= 0:
            action, comment = "✅ keep", "目標帯に到達。再現性を維持してください。"
        elif delta >= IMPROVEMENT_DELTA:
            action, comment = "📈 improve", f"改善中（{delta:+.3f}）。このまま継続。"
        elif delta <= -IMPROVEMENT_DELTA:
            action, comment = "⚠ watch", f"後退（{delta:+.3f}）。原因を確認してください。"
        elif abs(delta) < STABLE_DELTA:
            action, comment = "➡ stable", "変化なし。引き続き取り組んでください。"
        else:
            action, comment = "🔄 improve", f"微改善（{delta:+.3f}）。継続が重要です。"

        plan_records.append({
            "指標": feature_label(feat),
            "アクション": action,
            "今回": round(cur_score, 3),
            "前回": round(prv_score, 3),
            "コメント": comment,
        })

    if plan_records:
        st.dataframe(pd.DataFrame(plan_records), use_container_width=True, hide_index=True)
    st.caption(f"比較: {prev_date} → {latest_date}")


def show_personal_coaching_section(
    df: pd.DataFrame,
    columns: list[str],
    features: list[str],
    lp_results: list[dict[str, Any]],
    mr_results: list[dict[str, Any]],
) -> None:
    st.header("個別データ診断")
    st.caption(
        "ユーザーコードを入力するとBucklerから最新データを取得して診断します。"
        "ユーザーコードはBucklerのプロフィールURL末尾の数字です: "
        "https://www.streetfighter.com/6/buckler/profile/**XXXXXXXX**"
    )

    rank_options = build_rank_options(df)
    default_rank_idx = len(rank_options) - 1 if rank_options else 0

    c1, c2 = st.columns([3, 3])
    with c1:
        short_id_input = st.text_input(
            "ユーザーコード (short_id)",
            value="",
            placeholder="10桁の数字",
            help="BucklerプロフィールURL末尾の数字を入力してください",
        )
    with c2:
        target_rank = st.selectbox("目標ランク", options=rank_options, index=default_rank_idx if rank_options else None)

    run_diag = st.button("個別データ診断を実行", use_container_width=True, type="primary")

    if not run_diag:
        return

    # short_id を確定する
    short_id = short_id_input.strip()
    if short_id and not short_id.isdigit():
        st.warning("ユーザーコードは数字のみで入力してください。BucklerのプロフィールURL末尾の数字を確認してください。")
        return

    if not short_id:
        st.warning("ユーザーコードを入力してください。")
        return

    try:
        cookie = sp.load_cookie_text("", sp.DEFAULT_COOKIE_FILE)
        sp.validate_cookie_text(cookie)
        if not cookie:
            st.error(
                "個別診断機能は現在ご利用いただけません。"
                "サーバー設定が必要なため、しばらくお待ちください。"
            )
            return

        player_row = build_player_row_from_short_id(short_id, cookie, 30, columns)
    except ValueError as exc:
        st.error(f"入力内容に問題があります: {exc}")
        return
    except Exception as exc:
        st.error(
            "ユーザーデータの取得に失敗しました。"
            "ユーザーコードが正しいか確認してください。"
            f"（詳細: {exc}）"
        )
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
    # --- 同LP帯のサンプルも取得 ---
    current_lp = player_row.get("リーグポイント")
    current_band = lp_band(current_lp if isinstance(current_lp, (int, float)) else None)
    same_rank_df = df[df["LP帯"] == current_band].copy() if current_band not in ("不明", "範囲外") else pd.DataFrame()

    gap_df = compute_feature_gap_table(player_row, target_df, model_results, coaching_features)
    same_gap_df = compute_feature_gap_table(player_row, same_rank_df, model_results, coaching_features) if not same_rank_df.empty else pd.DataFrame()

    # --- ランク表示 ---
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("プレイヤー名", str(player_row.get("プレイヤー名", "-")))
    k2.metric("現在ランク", str(player_row.get("ランク", "-")))
    k3.metric("LP", f"{int(player_row['リーグポイント']) if isinstance(player_row.get('リーグポイント'), (int, float)) else '-'}")
    mr_val = player_row.get("MR")
    k4.metric("MR", f"{int(mr_val) if isinstance(mr_val, (int, float)) else '-'}")

    st.caption(f"目標ランク比較基準: {target_note}")
    if not same_rank_df.empty:
        st.caption(f"同ランク帯比較基準: {current_band} の {len(same_rank_df)} 件")

    # --- 同ランク帯との比較（テーブルのみ） ---
    st.subheader("同ランク帯との比較")
    if same_gap_df.empty:
        st.info("同ランク帯のサンプルが不足しているため比較できませんでした。")
    else:
        _show_gap_table(same_gap_df)

    # --- 目標ランク帯との比較（テーブルのみ） ---
    st.subheader("目標ランク帯との比較")
    if gap_df.empty:
        st.info("比較可能な指標が不足しているため、個別課題を算出できませんでした。")
    else:
        _show_gap_table(gap_df)

    # --- プレイ時間系 ---
    st.subheader("プレイ時間系（別軸）")
    show_play_volume_table(player_row, same_rank_df, target_df)
    for text in build_play_volume_advice(player_row, target_df):
        st.write(f"- {text}")

    # --- ゲージ使用割合 ---
    st.subheader("ゲージ使用割合")
    show_gauge_pie_charts(player_row, target_df)

    # --- 週次トレンド（Priority 3） ---
    st.subheader("週次トレンド")
    show_weekly_tracking_section(short_id, target_df, model_results, coaching_features)


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
    # 相関の絶対値が大きい順に上から並べる
    chart_df = chart_df.sort_values("abs_corr", ascending=True)
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
    fig.update_yaxes(categoryorder="array", categoryarray=chart_df["feature_display"].tolist())
    fig.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10), coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True)


def make_summary_text(
    total: int,
    sub_n: int,
    master_n: int,
    band_df: pd.DataFrame,
    lp_results: list[dict[str, Any]],
    mr_results: list[dict[str, Any]],
    progress_results: list[dict[str, Any]] | None = None,
) -> tuple[str | None, str | None, str]:
    """(positive_takeaway, negative_takeaway, detail_text) を返す。"""
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

    # 総合ランク要因分析（progress_results）のみから最大正相関・最大負相関を探す
    all_results: list[dict[str, Any]] = list(progress_results or [])

    best_pos: dict[str, Any] | None = None
    best_neg: dict[str, Any] | None = None
    for r in all_results:
        corr = r["correlation"]
        if corr > 0 and (best_pos is None or corr > best_pos["correlation"]):
            best_pos = r
        if corr < 0 and (best_neg is None or abs(corr) > abs(best_neg["correlation"])):
            best_neg = r

    # レガシー用（詳細テキストで使用）
    lp_pos = top_pos(lp_results)
    lp_neg = top_neg(lp_results)
    mr_pos = top_pos(mr_results)
    mr_neg = top_neg(mr_results)

    shortage_bands = band_df[band_df["最低60との差"] > 0]
    band_ok = shortage_bands.empty

    # テイクアウェイ（総合結果から正相関・負相関それぞれ1行）
    positive_takeaway: str | None = None
    negative_takeaway: str | None = None
    if best_pos is not None:
        positive_takeaway = (
            f"最大正相関要因: 「{feature_label(best_pos['feature'])}」（r={best_pos['correlation']:.3f}）— 伸ばすほどランクが上がる傾向"
        )
    if best_neg is not None:
        negative_takeaway = (
            f"最大負相関要因: 「{feature_label(best_neg['feature'])}」（r={best_neg['correlation']:.3f}）— 下げるほどランクが上がる傾向"
        )

    # 詳細（3-4行）
    detail_lines = [
        f"サンプル: 総{total}件（LPモデル {sub_n}件 / MRモデル {master_n}件）。"
        + ("全LP帯で最低60件確保済み。" if band_ok else f"不足帯あり（{', '.join(shortage_bands['LP帯'].tolist())}）—追加取得推奨。"),
    ]
    lp_parts = []
    if lp_pos:
        lp_parts.append(f"正相関: {feature_label(lp_pos['feature'])} r={lp_pos['correlation']:.3f}")
    if lp_neg:
        lp_parts.append(f"負相関: {feature_label(lp_neg['feature'])} r={lp_neg['correlation']:.3f}")
    if lp_parts:
        detail_lines.append("LP要因 — " + " / ".join(lp_parts))
    mr_parts = []
    if mr_pos:
        mr_parts.append(f"正相関: {feature_label(mr_pos['feature'])} r={mr_pos['correlation']:.3f}")
    if mr_neg:
        mr_parts.append(f"負相関: {feature_label(mr_neg['feature'])} r={mr_neg['correlation']:.3f}")
    if mr_parts:
        detail_lines.append("MR要因 — " + " / ".join(mr_parts))
    detail_lines.append("※探索分析。相関≠因果。効果量（R²）と n も合わせて判断すること。")

    return positive_takeaway, negative_takeaway, "\n".join(detail_lines)


def show_factor_section(title: str, results: list[dict[str, Any]], top_n: int, description: str = "") -> None:
    st.subheader(title)
    if description:
        st.caption(description)
    pos_df, neg_df = top_positive_negative(results, top_n)
    pos_df = add_display_feature(pos_df)
    neg_df = add_display_feature(neg_df)

    if pos_df.empty and neg_df.empty:
        st.info("表示できる要因がありません。")
        return

    left, right = st.columns(2)
    with left:
        if not pos_df.empty and "feature_display" in pos_df.columns:
            plot_factor_bar(pos_df, f"{title} 正相関 上位")
            st.dataframe(
                pos_df[["feature_display", "correlation", "n", "r_squared"]],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("表示できる正相関要因がありません。")
    with right:
        if not neg_df.empty and "feature_display" in neg_df.columns:
            plot_factor_bar(neg_df, f"{title} 負相関 上位")
            st.dataframe(
                neg_df[["feature_display", "correlation", "n", "r_squared"]],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("表示できる負相関要因がありません。")


def main() -> None:
    st.set_page_config(page_title="SF6 分析ダッシュボード", layout="wide")
    st.title("SF6 分析ダッシュボード")
    st.caption("本ダッシュボードは個人利用を目的としています。商用利用はできません。データの二次配布・転載はご遠慮ください。")

    with st.sidebar:
        st.header("設定")
        top_n = st.slider("上位表示件数", min_value=3, max_value=12, value=8)
        exclude_char_dependent = st.checkbox("キャラ依存の強い指標を除外", value=True)
        exclude_play_volume = st.checkbox("プレイ量系指標を除外", value=False)
        if st.button("データ再取得", use_container_width=True):
            load_from_supabase.clear()

    try:
        columns, rows = load_from_supabase()
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

    # ── Section 1: サンプル概要 ──────────────────────────────────────
    st.header("サンプル概要")
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("総件数", f"{total} 件")
    kpi2.metric("LPモデル件数", f"{sub_n} 件")
    kpi3.metric("MRモデル件数", f"{master_n} 件")

    # LP帯サンプル数（折りたたみ）
    with st.expander("LP帯サンプル数", expanded=False):
        lp_category_order = (
            band_df.assign(_sort=band_df["LP帯"].apply(sort_key_for_band))
            .sort_values("_sort")["表示LP帯"]
            .tolist()
        )
        fig_band = px.bar(
            band_df,
            x="件数",
            y="表示LP帯",
            orientation="h",
            text="件数",
            color="件数",
            color_continuous_scale="Teal",
        )
        fig_band.add_vline(x=60, line_dash="dash", line_color="orange", annotation_text="最低60")
        fig_band.add_vline(x=80, line_dash="dot", line_color="green", annotation_text="目標80")
        fig_band.update_yaxes(categoryorder="array", categoryarray=lp_category_order)
        fig_band.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
        st.plotly_chart(fig_band, use_container_width=True)
        st.dataframe(band_df, use_container_width=True, hide_index=True)

    # MRサンプル数（折りたたみ）
    mr_counts = mr_band_count_df(df)
    with st.expander("MRサンプル数", expanded=False):
        if not mr_counts.empty:
            fig_mr = px.bar(
                mr_counts,
                x="件数",
                y="表示ランク",
                orientation="h",
                text="件数",
                color="件数",
                color_continuous_scale="Purples",
            )
            fig_mr.update_yaxes(
                categoryorder="array",
                categoryarray=["MASTER", "HIGH MASTER", "GRAND MASTER", "ULTIMATE MASTER"],
            )
            fig_mr.update_layout(height=260, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
            st.plotly_chart(fig_mr, use_container_width=True)
            st.dataframe(mr_counts[["表示ランク", "件数"]], use_container_width=True, hide_index=True)
        else:
            st.info("MRモデルのサンプルがありません。")

    # ── Section 2: 主要知見 ──────────────────────────────────────────
    st.header("主要知見")
    pos_takeaway, neg_takeaway, detail_text = make_summary_text(total, sub_n, master_n, band_df, lp_results, mr_results, progress_results)
    if pos_takeaway:
        st.info(f"**{pos_takeaway}**")
    if neg_takeaway:
        st.error(f"**{neg_takeaway}**")
    if not pos_takeaway and not neg_takeaway:
        st.info(f"**分析対象 {total} 件。有意な相関は検出されませんでした（サンプル追加を推奨）。**")
    with st.expander("詳細サマリー"):
        st.text(detail_text)

    show_factor_section(
        "総合ランク要因分析",
        progress_results,
        top_n,
        description=(
            "【対象】全サンプル（BRONZE〜DIAMOND + MASTER以上）。LP/MRをグループ内Zスコア化して統合した"
            "総合スコアを目的変数として各指標との相関係数を算出。"
            "【注意】LP/MR要因分析の補完として解釈すること。"
        ),
    )

    show_factor_section(
        "LP要因分析",
        lp_results,
        top_n,
        description=(
            "【対象】LP が 25,000 未満のサンプル（BRONZE〜DIAMOND帯）。"
            "【分析方法】各指標とリーグポイント（LP）のピアソン相関係数を算出し、"
            "絶対値が大きいものを上位として表示。正相関＝LPが高い人ほど多い／高い傾向、"
            "負相関＝LPが高い人ほど少ない／低い傾向を示す。"
            "【注意】相関であり因果ではない。サンプル数（n）と決定係数（R²）も併せて判断すること。"
        ),
    )

    show_factor_section(
        "MR要因分析",
        mr_results,
        top_n,
        description=(
            "【対象】MASTER以上（MR管理帯）のサンプル。"
            "【分析方法】各指標とマスターレーティング（MR）のピアソン相関係数を算出。"
            "正相関＝MRが高い人ほど多い／高い傾向、負相関＝MRが高い人ほど少ない／低い傾向。"
            "【注意】MRはキャラ・マッチング補正が入るため、キャラ依存指標は除外して読むこと。"
        ),
    )

    # ── Section 3: 個別データ診断 ────────────────────────────────────
    st.divider()
    show_personal_coaching_section(df, columns, features, lp_results, mr_results)

    st.divider()
    st.caption(
        "本データはBucklers Boot Camp（CAPCOM）の公開情報をもとに個人が収集・分析したものです。"
        "公式情報・CAPCOM社とは一切関係ありません。個人利用を目的としています。"
    )


if __name__ == "__main__":
    main()

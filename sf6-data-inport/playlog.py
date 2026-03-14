from __future__ import annotations

import argparse
import codecs
import csv
import datetime as dt
import io
import json
import re
import sys
from pathlib import Path


PLAYER_NAME_LINE_NUMBER = 62
KNOWN_COUNTRY_SUFFIXES = (
    "アメリカ合衆国",
    "日本",
    "韓国",
    "中国",
    "台湾",
    "香港",
    "カナダ",
    "メキシコ",
    "チリ",
    "アルゼンチン",
    "コロンビア",
    "ペルー",
    "ベネズエラ",
    "エクアドル",
    "ボリビア",
    "パラグアイ",
    "ウルグアイ",
    "ブラジル",
    "イギリス",
    "フランス",
    "ドイツ",
    "イタリア",
    "スペイン",
    "オーストラリア",
)


def get_default_lp_master_path() -> Path:
    return Path(__file__).resolve().parent / "LP_master.json"


def get_default_mr_master_path() -> Path:
    return Path(__file__).resolve().parent / "MR_master.json"


def detect_text_encoding(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(codecs.BOM_UTF8):
        return "utf-8-sig"
    try:
        raw.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        return "cp932"


def read_text(path: Path) -> tuple[str, str]:
    encoding = detect_text_encoding(path)
    text = path.read_text(encoding=encoding)
    return text, encoding


def can_encode_text(text: str, encoding: str) -> bool:
    try:
        text.encode(encoding)
        return True
    except UnicodeEncodeError:
        return False


def migrate_text_file_to_utf8(path: Path) -> None:
    text, _ = read_text(path)
    path.write_bytes(codecs.BOM_UTF8 + text.encode("utf-8"))


def append_csv_row(path: Path, row: list[str]) -> None:
    out_encoding = detect_text_encoding(path)
    row_text = ",".join(row)

    if out_encoding == "cp932" and not can_encode_text(row_text, out_encoding):
        migrate_text_file_to_utf8(path)
        out_encoding = "utf-8-sig"

    with path.open("a", encoding=out_encoding, newline="") as f:
        writer = csv.writer(f)
        writer.writerow(row)


def extract_block(text: str, start_pat: str, end_pat: str) -> str:
    pattern = re.compile(start_pat + r"\s*(.*?)\s*" + end_pat, re.S | re.M)
    match = pattern.search(text)
    return match.group(1) if match else ""


def to_stripped_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def extract_following_value(text: str, label: str, value_pattern: str) -> str:
    lines = to_stripped_lines(text)
    pattern = re.compile(value_pattern)

    for index, line in enumerate(lines):
        if line != label:
            continue

        for candidate in lines[index + 1 : index + 5]:
            match = pattern.fullmatch(candidate)
            if match:
                return match.group(1)

    return ""


def extract_percent(text: str, label: str) -> str:
    pattern = re.compile(rf"(?m)^{re.escape(label)}\s*([0-9]+(?:\.[0-9]+)?)\s*%")
    match = pattern.search(text)
    if match:
        return f"{match.group(1)}%"

    following = extract_following_value(text, label, r"([0-9]+(?:\.[0-9]+)?)%")
    return f"{following}%" if following else ""


def extract_number(text: str, label: str, unit: str | None = None) -> str:
    if unit:
        pattern = re.compile(
            rf"(?m)^{re.escape(label)}\s*([0-9]+(?:\.[0-9]+)?)\s*{re.escape(unit)}"
        )
    else:
        pattern = re.compile(rf"(?m)^{re.escape(label)}\s*([0-9]+(?:\.[0-9]+)?)")
    match = pattern.search(text)
    if match:
        return match.group(1)

    if unit:
        return extract_following_value(
            text,
            label,
            rf"([0-9]+(?:\.[0-9]+)?){re.escape(unit)}",
        )

    return extract_following_value(text, label, r"([0-9]+(?:\.[0-9]+)?)")


def extract_int_like(text: str, label: str, unit: str | None = None) -> str:
    if unit:
        pattern = re.compile(rf"(?m)^{re.escape(label)}\s*([0-9,]+)\s*{re.escape(unit)}")
    else:
        pattern = re.compile(rf"(?m)^{re.escape(label)}\s*([0-9,]+)")
    match = pattern.search(text)
    if match:
        return match.group(1).replace(",", "")

    if unit:
        following = extract_following_value(text, label, rf"([0-9,]+){re.escape(unit)}")
    else:
        following = extract_following_value(text, label, r"([0-9,]+)")

    return following.replace(",", "") if following else ""


def get_next_no(output_path: Path) -> int:
    text, _ = read_text(output_path)
    rows = list(csv.reader(io.StringIO(text)))
    max_no = 0
    for row in rows[4:]:
        if row and row[0].strip().isdigit():
            max_no = max(max_no, int(row[0].strip()))
    return max_no + 1


def parse_lp_value(lp_text: str) -> int | None:
    cleaned = lp_text.strip().replace(",", "")
    if not cleaned:
        return None
    if not cleaned.isdigit():
        return None
    return int(cleaned)


def load_lp_master(lp_master_path: Path) -> list[dict[str, object]]:
    if not lp_master_path.exists():
        raise FileNotFoundError(f"LPマスターが見つかりません: {lp_master_path}")

    text, _ = read_text(lp_master_path)
    data = json.loads(text)
    if not isinstance(data, list):
        raise ValueError("LPマスターの形式が不正です")
    return data


def load_mr_master(mr_master_path: Path) -> list[dict[str, object]]:
    if not mr_master_path.exists():
        raise FileNotFoundError(f"MRマスターが見つかりません: {mr_master_path}")

    text, _ = read_text(mr_master_path)
    data = json.loads(text)
    if not isinstance(data, list):
        raise ValueError("MRマスターの形式が不正です")
    return data


def resolve_rank(lp_text: str, lp_master: list[dict[str, object]]) -> str:
    lp = parse_lp_value(lp_text)
    if lp is None:
        return ""

    sorted_master = sorted(lp_master, key=lambda item: int(item["min_lp"]), reverse=True)
    for item in sorted_master:
        min_lp = int(item["min_lp"])
        max_lp_raw = item.get("max_lp")
        max_lp = int(max_lp_raw) if max_lp_raw is not None else None

        if lp >= min_lp and (max_lp is None or lp < max_lp):
            rank = str(item["rank"])
            star = item.get("star")
            return rank if star is None else f"{rank}{int(star)}"

    return ""


def resolve_mr(lp_text: str) -> str:
    lp = parse_lp_value(lp_text)
    if lp is None or lp < 25000:
        return "0"

    if not sys.stdin.isatty():
        return "0"

    mr = input("MRを入力してください: ").strip()
    return mr if mr else "0"


def resolve_master_rank(
    lp_text: str,
    mr_text: str,
    mr_master: list[dict[str, object]],
) -> str:
    lp = parse_lp_value(lp_text)
    mr = parse_lp_value(mr_text)
    if lp is None or lp < 25000 or mr is None:
        return "MASTER"

    for item in mr_master:
        lp_min_raw = item.get("lp_min")
        lp_max_raw = item.get("lp_max")
        mr_min_raw = item.get("mr_min")
        mr_max_raw = item.get("mr_max")

        lp_min = int(lp_min_raw) if lp_min_raw is not None else None
        lp_max = int(lp_max_raw) if lp_max_raw is not None else None
        mr_min = int(mr_min_raw) if mr_min_raw is not None else None
        mr_max = int(mr_max_raw) if mr_max_raw is not None else None

        if lp_min is not None and lp < lp_min:
            continue
        if lp_max is not None and lp > lp_max:
            continue
        if mr_min is not None and mr < mr_min:
            continue
        if mr_max is not None and mr > mr_max:
            continue

        return str(item.get("abbr", "MASTER"))

    return "MASTER"


def build_row(
    in_text: str,
    no: int,
    date: str,
    player: str,
    league_points: str,
    rank: str,
    mr: str,
) -> list[str]:
    drive_reversal_block = extract_block(in_text, r"^ドライブリバーサル\s*$", r"^ドライブパリィ\s*$")
    drive_parry_block = extract_block(in_text, r"^ドライブパリィ\s*$", r"^ドライブインパクト\s*$")
    di_self_block = extract_block(in_text, r"【自分の使用】", r"【対戦相手の使用】")
    di_opp_block = extract_block(in_text, r"【対戦相手の使用】", r"^SAゲージ使用割合\s*$")
    throw_block = extract_block(in_text, r"^投げ\s*$", r"^壁際\s*$")
    wall_block = extract_block(in_text, r"^壁際\s*$", r"^ランクマッチプレイ回数\s*$")

    row = [""] * 40
    row[0] = str(no)
    row[1] = date
    row[2] = player
    row[3] = league_points
    row[4] = rank
    row[5] = mr

    row[6] = extract_percent(in_text, "ドライブパリィ")
    row[7] = extract_percent(in_text, "ドライブインパクト")
    row[8] = extract_percent(in_text, "オーバードライブアーツ")
    row[9] = extract_percent(in_text, "パリィドライブラッシュ")
    row[10] = extract_percent(in_text, "キャンセルドライブラッシュ")
    row[11] = extract_percent(in_text, "ドライブリバーサル")
    row[12] = extract_percent(in_text, "ダメージ")

    row[13] = extract_number(in_text, "使用回数", "回")
    row[14] = extract_number(in_text, "成功回数", "回")
    row[15] = extract_number(in_text, "相手のドライブパリィを投げた", "回")
    row[16] = extract_number(in_text, "自分のドライブパリィを投げられた", "回")
    row[17] = extract_number(in_text, "ジャストパリィ回数", "回")

    row[18] = extract_number(di_self_block, "決めた回数", "回")
    row[19] = extract_number(di_self_block, "パニッシュカウンターを決めた回数", "回")
    row[20] = extract_number(di_self_block, "相手のドライブインパクトに決めた回数", "回")
    row[21] = extract_number(di_opp_block, "受けた回数", "回")
    row[22] = extract_number(di_opp_block, "パニッシュカウンターを受けた回数", "回")
    row[23] = extract_number(di_opp_block, "相手にドライブインパクトで返された回数", "回")

    row[24] = extract_percent(in_text, "Lv1")
    row[25] = extract_percent(in_text, "Lv2")
    row[26] = extract_percent(in_text, "Lv3")
    row[27] = extract_percent(in_text, "CA")

    row[28] = extract_number(in_text, "スタンさせた回数", "回")
    row[29] = extract_number(in_text, "スタンさせられた回数", "回")

    row[30] = extract_number(throw_block, "決めた回数", "回")
    row[31] = extract_number(throw_block, "受けた回数", "回")
    row[32] = extract_number(throw_block, "投げ抜け回数", "回")

    row[33] = extract_number(wall_block, "相手を追い詰めている時間", "秒")
    row[34] = extract_number(wall_block, "相手に追い詰められている時間", "秒")

    row[35] = extract_int_like(in_text, "ランクマッチプレイ回数", "回")
    row[36] = extract_int_like(in_text, "カジュアルマッチプレイ回数", "回")
    row[37] = extract_int_like(in_text, "ルームマッチプレイ回数", "回")
    row[38] = extract_int_like(in_text, "バトルハブマッチプレイ回数", "回")
    row[39] = extract_int_like(in_text, "累計プレイポイント", "PT")

    return row


def normalize_date(raw: str) -> str:
    if not raw.strip():
        today = dt.date.today()
        return f"{today.year}/{today.month}/{today.day}"
    return raw.strip()


def get_today_date() -> str:
    today = dt.date.today()
    return f"{today.year}/{today.month}/{today.day}"


def repair_mojibake(text: str) -> str:
    if "ユーザーコード:" in text and "ドライブ" in text:
        return text

    try:
        repaired = text.encode("cp932", errors="ignore").decode("utf-8", errors="ignore")
    except Exception:
        return text

    if "ユーザーコード:" in repaired and "ドライブ" in repaired:
        return repaired

    return text


def is_valid_playlog_text(text: str) -> bool:
    required = ("ユーザーコード:", "ドライブゲージ使用実績", "ランクマッチプレイ回数")
    return all(token in text for token in required)


def prompt_paste_text() -> str:
    if sys.stdin.isatty():
        import queue
        import threading

        print("実績テキストを張り付けて、Enter で確定します。")

        PAUSE_TIMEOUT = 0.5  # 貼り付け後この秒数無入力で終了とみなす

        line_queue: queue.Queue[str | None] = queue.Queue()

        def _reader() -> None:
            try:
                while True:
                    line_queue.put(input())
            except EOFError:
                line_queue.put(None)

        threading.Thread(target=_reader, daemon=True).start()

        lines: list[str] = []
        first = line_queue.get()  # 最初の行は無制限に待つ
        if first is not None:
            lines.append(first)
            while True:
                try:
                    item = line_queue.get(timeout=PAUSE_TIMEOUT)
                    if item is None:
                        break
                    lines.append(item)
                except queue.Empty:
                    break  # PAUSE_TIMEOUT 秒無入力 = 貼り付け完了

        pasted = "\n".join(lines)
    else:
        pasted = sys.stdin.read()

    if not pasted.strip():
        raise ValueError("実績テキストが入力されていません")
    normalized = pasted.encode("utf-8", errors="replace").decode("utf-8")
    fixed = repair_mojibake(normalized)
    if not is_valid_playlog_text(fixed):
        raise ValueError(
            "貼り付けテキストを正しく読み取れませんでした。"
            "入力文字コードが崩れている可能性があります。"
        )
    return fixed


def write_input_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8", errors="replace")


def normalize_player_name(raw: str) -> str:
    player = raw.strip()
    if not player:
        return ""

    for country in sorted(KNOWN_COUNTRY_SUFFIXES, key=len, reverse=True):
        if player.endswith(country) and len(player) > len(country):
            return player[: -len(country)].strip()

    return player


def extract_player_name(in_text: str) -> str:
    raw_lines = to_stripped_lines(in_text)

    for i, line in enumerate(raw_lines):
        if line.startswith("ユーザーコード:"):
            j = i - 1
            if j >= 0:
                candidate = raw_lines[j]
                if candidate in KNOWN_COUNTRY_SUFFIXES and j - 1 >= 0:
                    candidate = raw_lines[j - 1]
                return normalize_player_name(candidate)

    raw_input_lines = [line.strip() for line in in_text.splitlines()]
    if 1 <= PLAYER_NAME_LINE_NUMBER <= len(raw_input_lines):
        fixed_line = raw_input_lines[PLAYER_NAME_LINE_NUMBER - 1].strip()
        if fixed_line:
            return normalize_player_name(fixed_line)

    return ""


def main() -> None:
    parser = argparse.ArgumentParser(
        description="sf6_playlog_in.txt の貼り付けテキストを sf6-playlog-out.csv へ1行追記する"
    )
    parser.add_argument("--input", default="sf6_playlog_in.txt", help="入力テキストファイル")
    parser.add_argument("--output", default="sf6-playlog-out.csv", help="出力CSVファイル")
    parser.add_argument("--lp-master", default="", help="LPマスターファイル (既定: LP_master.json)")
    parser.add_argument("--mr-master", default="", help="MRマスターファイル (既定: MR_master.json)")
    parser.add_argument("--player", default="", help="プレイヤー名 (通常は自動取得。必要時のみ上書き)")
    parser.add_argument("--league-points", default="", help="リーグポイント (任意)")
    parser.add_argument("--dry-run", action="store_true", help="追記せず、生成行だけ表示")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)
    lp_master_path = Path(args.lp_master) if args.lp_master.strip() else get_default_lp_master_path()
    mr_master_path = Path(args.mr_master) if args.mr_master.strip() else get_default_mr_master_path()

    if not output_path.exists():
        raise FileNotFoundError(f"出力ファイルが見つかりません: {output_path}")

    lp_master = load_lp_master(lp_master_path)

    in_text = prompt_paste_text()
    write_input_text(input_path, in_text)

    league_points = args.league_points.strip()
    if not league_points:
        league_points = input("リーグポイントを入力してください [不明ならEnter]: ").strip()

    rank = resolve_rank(league_points, lp_master)

    mr = resolve_mr(league_points)
    if rank == "MASTER":
        mr_master = load_mr_master(mr_master_path)
        rank = resolve_master_rank(league_points, mr, mr_master)

    player = args.player.strip() or extract_player_name(in_text)

    if not player:
        raise ValueError(
            "プレイヤー名を自動取得できませんでした。"
            f"入力テキストの{PLAYER_NAME_LINE_NUMBER}行目を確認するか --player を指定してください"
        )

    date = get_today_date()

    next_no = get_next_no(output_path)
    row = build_row(in_text, next_no, date, player, league_points, rank, mr)

    if args.dry_run:
        print(",".join(row))
        return

    try:
        append_csv_row(output_path, row)
    except PermissionError as exc:
        raise PermissionError(
            f"{output_path.name} に書き込めません。Excel など他のアプリで開いている場合は閉じて再実行してください"
        ) from exc

    print(f"追記しました: {output_path} (No={next_no}, Player={player}, Date={date})")


if __name__ == "__main__":
    main()

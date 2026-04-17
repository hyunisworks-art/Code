"""SUPABASE_URLのホスト名だけ抽出してDNS解決を試す。"""
from __future__ import annotations

import os
import socket
import sys
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()

OUT = Path("supabase_dns_report.txt")
lines: list[str] = []


def log(msg: str) -> None:
    lines.append(msg)
    print(msg)


def main() -> None:
    url = os.getenv("SUPABASE_URL") or ""
    parsed = urlparse(url)
    host = parsed.hostname or ""
    log(f"host: {host}")

    if not host:
        log("NG: host抽出失敗")
        return

    try:
        ip = socket.gethostbyname(host)
        log(f"OK: DNS解決成功 -> {ip}")
    except Exception as e:
        log(f"NG: DNS解決失敗 ({type(e).__name__}: {e})")

    # 対照実験: supabase.com 本体（存在するはず）
    try:
        ip2 = socket.gethostbyname("supabase.com")
        log(f"対照: supabase.com -> {ip2} (ネット側は生存)")
    except Exception as e:
        log(f"対照NG: supabase.com解決失敗 ({e}) → ネット全体の問題かも")


if __name__ == "__main__":
    try:
        main()
    finally:
        OUT.write_text("\n".join(lines), encoding="utf-8")

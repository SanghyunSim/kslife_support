#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Telegram Bot API로 알림 전송 (표준 라이브러리만 사용, 추가 의존성 없음)."""
import json
import os
import urllib.request

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")


def _send(text: str) -> None:
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print(f"[telegram 설정 없음, 알림 생략] {text}")
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = json.dumps({"chat_id": TELEGRAM_CHAT_ID, "text": text}).encode("utf-8")
    req = urllib.request.Request(
        url, data=payload, headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        resp.read()


def send_telegram_message(text: str) -> None:
    _send(text)


def send_telegram_error(text: str) -> None:
    _send("[오류] " + text)

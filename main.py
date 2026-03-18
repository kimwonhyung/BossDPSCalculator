# -*- coding: utf-8 -*-
"""
🔥 6인 보스 DPS 계산기  |  시즌2
CustomTkinter 기반 GUI  →  PyInstaller로 exe 변환 가능
"""

import customtkinter as ctk
import tkinter as tk
import json
import os
import sys
import importlib
import math
import re
import time
import threading
import ctypes
import hashlib
import tempfile
import subprocess
from ctypes import wintypes
from pathlib import Path
from urllib import request as urlrequest
from urllib import error as urlerror

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  공식 / 상수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
WEAPON_BASE = {
    "없음":       lambda lvl: (0 * 0),
    "아비터":       lambda lvl: (4675 + 295 * 100) * 4.8,
    "뮤탈":       lambda lvl: (6000 + 400 * 100) * 8,
    "뮤탈(후원)": lambda lvl: (6000 + 400 * 115) * 8,
    "배틀":       lambda lvl: (13500 + 520 * 100) * 12,
    "배틀(후원)": lambda lvl: (13500 + 520 * 115) * 12,
    "다칸":       lambda lvl: (65000 + 550 * 100) * 14,
    "다칸(후원)": lambda lvl: (65000 + 550 * 115) * 14,
    "롱기": lambda lvl: (80000 + 800 * 100) * 20,
    "롱기(후원)": lambda lvl: (80000 + 800 * 115) * 20,
}
WEAPONS = list(WEAPON_BASE.keys())

SUB_BASE = {
    "없음": 0,
    "레어": 27800 * 2,
    "에픽": 34780 * 2.4,
    "유니크": 41700 * 2.67,
    "레전드리": 34780 * 8,
    "태초": 69600 * 12,
    "신화": 139000 * 12,
}
SUB_STACKS = ["X", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]

BOSSES = [
    ("이지 아르카인", 7.5), ("노말 아르카인", 15.5),
    ("하드 아르카인", 33), ("익스 아르카인", 53),("헬 아르카인", 100),
    ("이지 이터모르트", 9), ("노말 이터모르트", 18.75),
    ("하드 이터모르트", 35.5), ("익스 이터모르트", 60),("헬 이터모르트", 108),
    ("이지 크라켄드", 10.5), ("노말 크라켄드", 23.3),
    ("하드 크라켄드", 38), ("익스 크라켄드", 67),("헬 크라켄드", 115),
    ("이지 노바리스", 12), ("노말 노바리스", 26.2),
    ("하드 노바리스", 41), ("익스 노바리스", 75),("헬 노바리스", 125),
    ("이지 엔드레이어", 14.5), ("노말 엔드레이어", 30),
    ("하드 엔드레이어", 46), ("익스 엔드레이어", 85),("헬 엔드레이어", 150),
]

SAVE_FILE = Path(os.path.expanduser("~")) / ".dps_calc_save.json"
# 업데이트 시 latest.json과 함께 버전, URL, SHA256 해시 갱신 필요  
APP_VERSION = "1.5"
# GitHub에서 자동 업데이트 확인 (배포된 exe에서만 작동)
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/kimwonhyung/BossDPSCalculator/main/latest.json"
UPDATE_CHECK_TIMEOUT_SEC = 4
UPDATE_DOWNLOAD_TIMEOUT_SEC = 25


def _resource_path(filename: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / filename


def _to_int(value, default=0, min_value=None, max_value=None):
    try:
        val = int(value)
    except (TypeError, ValueError):
        val = default
    if min_value is not None:
        val = max(min_value, val)
    if max_value is not None:
        val = min(max_value, val)
    return val


def _parse_version_tuple(version_text: str):
    nums = re.findall(r"\d+", str(version_text or ""))
    if not nums:
        return (0,)
    return tuple(int(n) for n in nums[:4])


def _is_remote_version_newer(remote_version: str, local_version: str) -> bool:
    return _parse_version_tuple(remote_version) > _parse_version_tuple(local_version)


def _sha256_file(file_path: Path) -> str:
    h = hashlib.sha256()
    with file_path.open("rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _normalize_subs(raw_subs):
    normalized = []
    for i in range(6):
        grade = "없음"
        stack = "없음"

        if isinstance(raw_subs, list) and i < len(raw_subs):
            item = raw_subs[i]
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                candidate_grade = str(item[0])
                candidate_stack = str(item[1])

                if candidate_grade in SUB_BASE:
                    grade = candidate_grade

                if candidate_stack in SUB_STACKS:
                    stack = candidate_stack
                elif candidate_stack == "없음":
                    stack = "없음"

        # 스택이 X/없음이면 계산에서는 미보유로 통일
        if stack in ("X", "없음") or grade == "없음":
            normalized.append(("없음", "없음"))
        else:
            normalized.append((grade, stack))

    return normalized


def _normalize_player(raw_player, idx):
    p = raw_player if isinstance(raw_player, dict) else {}
    weapon = p.get("weapon") if p.get("weapon") in WEAPON_BASE else "배틀"
    legacy_seal = _to_int(p.get("seal", 10), default=10, min_value=0, max_value=999)
    seal_re = _to_int(p.get("seal_re", legacy_seal), default=legacy_seal, min_value=0, max_value=999)
    seal_shin = _to_int(p.get("seal_shin", 0), default=0, min_value=0, max_value=999)
    seal = seal_re + seal_shin
    seal_lvl = _to_int(p.get("seal_lvl", 110), default=110, min_value=0, max_value=110)
    name = str(p.get("name") or f"{idx + 1}P")
    subs = _normalize_subs(p.get("subs", []))
    return {
        "weapon": weapon,
        "seal_re": seal_re,
        "seal_shin": seal_shin,
        "seal": seal,
        "seal_lvl": seal_lvl,
        "subs": subs,
        "name": name,
    }


# ── 계산 함수 ────────────────────────────────────────────────────
def calc_dps(weapon, seal_re_cnt, seal_shin_cnt, seal_lvl, subs):
    """총 DPS 반환 (단위: 절대값, /1e8 하면 억)"""
    seal_lvl = _to_int(seal_lvl, default=0, min_value=0, max_value=110)
    seal_re_cnt = _to_int(seal_re_cnt, default=0, min_value=0)
    seal_shin_cnt = _to_int(seal_shin_cnt, default=0, min_value=0)
    w_base = WEAPON_BASE.get(weapon, lambda x: 0)(seal_lvl)
    # [레]인: 기존 인장 공식
    seal_re_damage = seal_re_cnt * (9900 + 495 * seal_lvl) * 2.67
    # [신]인: 기본 공격력 20000, 1업당 550, 공속 8
    seal_shin_damage = seal_shin_cnt * (20000 + 550 * seal_lvl) * 8
    sub_total = 0
    for item in (subs if isinstance(subs, list) else []):
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            continue
        grade, stack = str(item[0]), str(item[1])
        if grade == "없음" or stack in ("없음", "X"):
            continue
        base = SUB_BASE.get(grade, 0)
        stack_i = _to_int(stack, default=0, min_value=0, max_value=9)
        sub_total += base * (1 + stack_i * 0.5)
    return (w_base + seal_re_damage + seal_shin_damage + sub_total) * 180


def default_player(i):
    return dict(weapon="다칸", seal_re=15, seal_shin=0, seal=15, seal_lvl=110,
                subs=[("없음", "없음")] * 6,
                name=f"{i+1}P")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  색상 팔레트
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
C = dict(
    bg="#0C0E1A",
    fg_color="#2A2F4A",
    panel="#12152A",
    card="#181B2E",
    input="#1E2238",
    border="#252840",
    fire1="#FF3D3D",
    fire2="#FF7A00",
    gold="#FFD600",
    blue="#4D9FFF",
    blue2="#0004FF",
    purple="#B06EFF",
    purple2="#7300FF",
    teal="#00D4C8",
    green="#3DFF8F",
    white="#FFFFFF",
    gray="#7A7F9A",
    dim="#3A3F5A",
    dim2="#4B2800",
    bossname="#7EF6FF",
    boss_highlight="#66F2FF",
    need_interceptor="#FFD600",
    green2="#24C024",
    red="#FF0000"
)

PLAYER_COLORS = [C["fire1"], C["blue"], C["green"],
                 C["purple"], C["fire2"], C["dim2"]]

DIFF_COLOR = {"이지": C["gray"], "노말": C["green"],
              "하드": C["fire2"], "익스": C["fire1"],
              "헬": C["red"]}

INFO_SPECIAL_TEXT_COLORS = [
    ("태초 조건", C["purple"]),
    ("태초 무기 강화 확률 증가+5%", C["purple"]),
    ("태초 확률+5%", C["purple"]),
    ("강화 확률+0.01%", C["gold"]),
    ("강화 확률+0.02%", C["gold"]),
    ("강화 확률+0.03%", C["gold"]),
    ("강화 확률+0.05%", C["gold"]),
    ("신화 조건", C["red"]),
    ("신화 확률+10%", C["red"]),
    ("신화 확률+5%", C["red"]),
    ("신화의 파편+1", C["green2"]),
    ("초월석+1", C["blue"]),
    ("초월석+2", C["blue"]),
    ("초월석+3", C["blue"]),
    ("초월석+4", C["blue"]),
    ("초월석 교환 확률+0.5%", C["blue"]),
    ("초월석 교환 확률+1%", C["blue"]),
    ("곡괭이 단계+1", C["red"]),
    ("곡괭이 공업+1", C["red"]),
    ("상자 레벨 최대치+1", C["purple2"]),
    ("다이아몬드 상자", C["blue2"]),
    ("행운의 상자", C["green2"]),   
]

INFO_HIGHLIGHT_LEGEND = [
    ("초월석 관련", C["blue"]),
    ("강화 확률 관련", C["gold"]),
    ("태초 관련", C["purple"]),
    ("신화 관련", C["red"]),
]

SCORE_TABLE_COLUMNS = [
    ("아르카인", "#8A2BE2"),
    ("이터모르트", "#E53935"),
    ("크라켄드", "#00C853"),
    ("노바리스", "#4A90E2"),
    ("엔드레이어", "#E8B8B8"),
    ("종합 점수", "#E89A3D"),
]

SCORE_TABLE_ROWS = [
    (5, [
        "마력의 흔적+10\n인장+5",
        "마력의 흔적+10\n인장+5",
        "마력의 흔적+15\n인장+7",
        "마력의 흔적+20\n인장+10",
        "마력의 흔적+25\n인장+15",
        "마력의 흔적+30\n인장+20\n강화 확률+0.01%",
    ]),
    (25, [
        "초월석+1",
        "초월석 교환 확률+0.5%",
        "초월석 교환 확률+0.5%",
        "강화 확률+0.01%",
        "강화 확률+0.01%",
        "초월석+1\n강화 확률+0.01%",
    ]),
    (50, [
        "인장+10\n인장 공업+1",
        "초월석+1",
        "인장+10\n인장 공업+1",
        "인장+15\n인장 공업+1",
        "인장+20\n인장 공업+1",
        "인장+30\n인장 공업+1\n강화 확률+0.01%",
    ]),
    (75, [
        "곡괭이 단계+1",
        "곡괭이 공업+1",
        "초월석+1",
        "초월석 교환 확률+0.5%",
        "초월석 교환 확률+0.5%",
        "초월석 교환 확률+0.5%",
    ]),
    (100, [
        "초월석 교환 확률+0.5%",
        "곡괭이 단계+1",
        "곡괭이 공업+1",
        "초월석+1",
        "초월석 교환 확률+0.5%",
        "초월석+1",
    ]),
    (150, [
        "초월석 교환 확률+0.5%",
        "초월석 교환 확률+0.5%",
        "곡괭이 단계+1",
        "곡괭이 공업+1",
        "초월석+1",
        "초월석+1",
    ]),
    (200, [
        "초월석 교환 확률+1%",
        "행운의 상자",
        "상자 레벨 최대치+1",
        "곡괭이 단계+1",
        "곡괭이 공업+1",
        "초월석+1",
    ]),
    (250, [
        "초월석 교환 확률+1%",
        "초월석 교환 확률+1%",
        "행운의 상자",
        "상자 레벨 최대치+1",
        "곡괭이 단계+1",
        "초월석+1",
    ]),
    (300, [
        "초월석 교환 확률+1%\n강화 확률+0.01%",
        "초월석 교환 확률+1%\n강화 확률+0.01%",
        "초월석 교환 확률+1%\n강화 확률+0.01%",
        "초월석 교환 확률+1%\n강화 확률+0.01%",
        "초월석 교환 확률+1%\n강화 확률+0.01%",
        "다이아몬드 상자\n[§ 초월 §]",
    ]),
    (500, [
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "초월석+3\n[§ 고대 §]",
    ]),
    (750, [
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "초월석+3\n[§ 흑초 §]",
    ]),
    (1000, [
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "마력의 흔적+10\n초월석 교환 확률+0.5%\n강화 확률+0.01%",
        "초월석+4\n[§ 금태 §]",
    ]),
]


def _split_info_special_segments(text):
    text = str(text)
    if not text or text == "-":
        return [(text, None)]

    ordered_items = sorted(INFO_SPECIAL_TEXT_COLORS, key=lambda item: len(item[0]), reverse=True)
    pattern = "|".join(re.escape(keyword) for keyword, _ in ordered_items)
    if not pattern:
        return [(text, None)]

    color_map = {keyword: color for keyword, color in ordered_items}
    segments = []
    last_index = 0

    for match in re.finditer(pattern, text):
        start, end = match.span()
        if start > last_index:
            segments.append((text[last_index:start], None))
        matched_text = match.group(0)
        segments.append((matched_text, color_map.get(matched_text)))
        last_index = end

    if last_index < len(text):
        segments.append((text[last_index:], None))

    return segments if segments else [(text, None)]

INFO_DB_COLUMNS = [
    ("난이도", 44),
    ("체력(억)", 58),
    ("필요 어기량", 72),
    ("3회 클리어 보상", 360),
    ("3회 칭호명", 170),
    ("10회 클리어 보상 (이지 6회)", 420),
    ("10회 칭호명", 190),
]

INFO_DB_ROWS = [
    {
        "boss": "경계의 파수자 아르카인",
        "color": "#8A2BE2",
        "rows": [
            ("Easy", "7.5", "80", "마흔+10, 인장+1", "[경계를 넘은 자]", "마흔+10, 인장+2, 잔재+1, 초월석+1", "[금벽 돌파자]"),
            ("Normal", "15.5", "100", "마흔+15, 인장+3, 초월석 교환 확률+1%", "[균열을 헤친 자]", "마흔+15, 인장+6, 잔재+1, 인공+1, 초월석 교환 확률+1%", "[균열 지배자]"),
            ("Hard", "33", "135", "인장+18, 인공+1, 태초 조건", "[금단을 파쇄한 자]", "인장+25, 초월석+1, 태초 무기 강화 확률 증가+5%", "[금단 종결자]"),
            ("Extreme", "53", "250", "마흔+15, 인장+15, 초월석 교환 확률+1%, 태초 확률+5%", "[◇경계의 지배자◇]", "강화 확률+0.02%, 인장+15, 초월석+1, 행운의 상자+1", "[〈 I 아르카인 I 〉]"),
            ("Hell", "100", "275", "-", "-", "-", "-"),
        ],
    },
    {
        "boss": "운명 포식자 이터모르트",
        "color": "#E53935",
        "rows": [
            ("Easy", "9", "85", "마흔+10, 인장+1", "[운명을 거스른 자]", "마흔+10, 잔재+0, 인장+3, 인공+1, 초월석+1", "[운명 인도자]"),
            ("Normal", "18.75", "105", "마흔+15, 인장+4, 초월석 교환 확률+1%", "[전명을 뒤엎은 자]", "마흔+15, 인장+7, 잔재+1, 인공+1, 초월석 교환 확률+2%", "[숙명 종결자]"),
            ("Hard", "35.5", "140", "인장+19, 인공+1, 초월석+1", "[운명을 삼킨 자]", "인장+25, 초월석+1, 태초 무기 강화 확률 증가+5%", "[전명의 군주]"),
            ("Extreme", "60", "250", "인장+20, 초월석 교환 확률+1%, 초월석+1", "[운명의 관측자]", "신화 조건", "[《 ◎ 이터모르트 ◎ 》]"),
            ("Hell", "108", "275", "-", "-", "-", "-"),
        ],
    },
    {
        "boss": "폭군의 화신 크라켄드",
        "color": "#00C853",
        "rows": [
            ("Easy", "10.5", "90", "마흔+10, 인장+3, 초월석 교환 확률+1%", "[폭풍을 밀어뜨린 자]", "마흔+10, 인장+4, 잔재+1, 초월석+1", "[포군 종결자]"),
            ("Normal", "23.3", "110", "마흔+20, 인장+5, 강화 확률+0.01%", "[파도를 꺾은 자]", "마흔+20, 인장+8, 잔재+1, 인공+1, 초월석+1", "[파도의 화신]"),
            ("Hard", "38", "150", "인장+20, 인공+1, 초월석+1", "[폭군의 시대를 끝낸 자]", "인공+2, 다수+1, 태초 무기 강화 확률 증가+5%", "[철권 군주]"),
            ("Extreme", "67", "250", "신화 확률+5%, 초월석+1", "[◆무기의 왕제◆]", "마흔+30, 초월석+1, 강화 확률+0.05%", "[§ 크라켄드 §]"),
            ("Hell", "115", "275", "-", "-", "-", "-"),
        ],
    },
    {
        "boss": "파멸의 기사 노바리스",
        "color": "#4A90E2",
        "rows": [
            ("Easy", "12", "95", "마흔+15, 인장+4, 초월석 교환 확률+1%", "[파멸을 저지한 자]", "마흔+15, 인장+5, 인공+1, 강화 확률+0.03%", "[선봉 용사]"),
            ("Normal", "26.2", "120", "마흔+25, 인장+6, 강화 확률+0.01%", "[멸망을 꺾은 자]", "마흔+25, 인장+9, 잔재+1, 인공+1, 초월석+1", "[종언 수호자]"),
            ("Hard", "41", "180", "인장+21, 인공+1, 초월석+1", "[종말의 행진을 늦춘 자]", "행상+1, 초월석+1, 태초 무기 강화 확률 증가+5%", "[단절의 수라]"),
            ("Extreme", "75", "250", "신화 확률+5%", "[◆파멸의 구원자◆]", "신화의 파편+1", "[† 노바리스 †]"),
            ("Hell", "125", "275", "-", "-", "-", "-"),
        ],
    },
    {
        "boss": "초월자 엔드레이어",
        "color": "#E8B8B8",
        "rows": [
            ("Easy", "14.5", "100", "마흔+15, 인장+5, 강화 확률+0.03%", "[세계의 수호자]", "마흔+15, 인장+6, 잔재+2, 인공+3, 강화 확률+0.02%, 초월석+1", "[천계 강림]"),
            ("Normal", "30", "130", "마흔+30, 인장+7, 강화 확률+0.01%, 초월석 교환 확률+1%", "[세계를 뒤흔든 자]", "인장+10, 잔재+2, 인공+2, 행상+1, 초월석+1", "[만물의 주인]"),
            ("Hard", "46", "200", "인장+22, 인공+1, 초월석+1, 강화 확률+0.02%", "[세계를 초월한 자]", "초월석+1, 강화 확률+0.03%, 태초 무기 강화 확률 증가+5%", "[권능의 왕]"),
            ("Extreme", "85", "250", "신화 확률+10%", "[◇영겁의 절대자◇]", "신화의 파편+1", "[★ 엔드레이어 ★]"),
            ("Hell", "150", "275", "-", "-", "-", "-"),
        ],
    },
]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  온보딩 화면
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class OnboardWindow(ctk.CTkToplevel):
    def __init__(self, parent, on_done):
        super().__init__(parent)
        self.on_done = on_done
        self.title("처음 오셨나요?")
        self.geometry("560x520")
        self.resizable(False, False)
        self.configure(fg_color=C["bg"])
        self.grab_set()
        self.step = 0
        self._build_steps()
        self._show(0)
        self.after(10, self._center_on_screen)

    def _center_on_screen(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_steps(self):
        self.steps_data = [
            {
                "icon": "📋",
                "title": "입력 방법",
                "sub": "3단계면 충분합니다",
                "body": (
                    "① 무기 선택\n"
                    "   드롭다운에서 각 플레이어 주무기 선택\n\n"
                    "② 인장 입력\n"
                    "   [레]인 / [신]인 개수와 인장 공업 수치 입력\n\n"
                    "③ 보조무기 스택 선택\n"
                    "   등급별 스택(X / 0~9) 선택\n"
                    "   X = 해당 등급 미보유\n\n"
                    "💡 무기/인장/공업/보조무기는\n"
                    "   【마우스 휠】로도 값 조정 가능\n\n"
                    "→ 결과는 오른쪽 패널에서 실시간 확인"
                ),
            },
            {
                "icon": "💡",
                "title": "알아두면 좋은 것",
                "sub": "오버레이 & 팁",
                "body": (
                    "📌 '위에 고정' → 타이틀바 제거 + 게임 위 오버레이\n"
                    "   헤더 왼쪽 드래그로 창 위치 이동\n"
                    "   ESC 키로 오버레이 모드 해제 가능\n\n"
                    "⌨ 단축키\n"
                    "  1~6 : 해당 플레이어 카드 ON/OFF\n"
                    "  F6 : 스타 포커스 해제 후 계산기 포커스\n"
                    "  - : 현재 딜량/트라이 정보 채팅 전송\n"
                    "  F7 : 매크로 커스텀 5줄 채팅 전송\n"
                    "  F12 : 창 숨김/다시 표시\n"
                    "  ESC : 오버레이(위에 고정) 해제\n\n"
                    "💾 설정은 자동 저장되며, 다음 실행 시\n"
                    "   그대로 불러옵니다."
                ),
            },
            {
                "icon": "💝",
                "title": "후원",
                "sub": "개발 후원 계좌 안내",
                "body": (
                    "카카오뱅크\n"
                    "3333-22-5948-231\n"
                    "예금주: 김원형\n\n"
                    "항상 사용해 주셔서 감사합니다."
                ),
            },
        ]
        self.n = len(self.steps_data)
        self.frames = [self._make_step_frame(d) for d in self.steps_data]

    def _make_step_frame(self, d):
        f = ctk.CTkFrame(self, fg_color="transparent")
        ctk.CTkLabel(f, text=d["icon"], font=("맑은 고딕", 48),
                     text_color=C["white"]).pack(pady=(40, 4))
        ctk.CTkLabel(f, text=d["title"],
                     font=("맑은 고딕", 20, "bold"),
                     text_color=C["white"]).pack()
        ctk.CTkLabel(f, text=d["sub"],
                     font=("맑은 고딕", 11),
                     text_color=C["gray"]).pack(pady=(2, 18))
        box = ctk.CTkFrame(f, fg_color=C["card"], corner_radius=10)
        box.pack(padx=40, fill="x")
        ctk.CTkLabel(box, text=d["body"],
                     font=("맑은 고딕", 11),
                     text_color=C["gray"],
                     justify="left").pack(padx=20, pady=16)
        return f

    def _show(self, idx):
        for f in self.frames:
            f.pack_forget()
        self.frames[idx].pack(fill="both", expand=True)
        self.step = idx
        self._rebuild_nav()

    def _rebuild_nav(self):
        if hasattr(self, "_nav"):
            self._nav.destroy()
        self._nav = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=72)
        self._nav.pack(side="bottom", fill="x")
        self._nav.pack_propagate(False)

        dots = ctk.CTkFrame(self._nav, fg_color="transparent")
        dots.place(relx=0.5, rely=0.28, anchor="center")
        for i in range(self.n):
            color = C["fire1"] if i == self.step else C["dim"]
            ctk.CTkLabel(dots, text="●", font=("맑은 고딕", 8),
                         text_color=color, width=16).pack(side="left")

        btn_row = ctk.CTkFrame(self._nav, fg_color="transparent")
        btn_row.place(relx=0.5, rely=0.70, anchor="center")

        if self.step > 0:
            ctk.CTkButton(btn_row, text="← 이전", width=100,
                          fg_color=C["input"], hover_color=C["border"],
                          text_color=C["gray"], font=("맑은 고딕", 11),
                          command=lambda: self._show(self.step - 1)).pack(side="left", padx=6)

        is_last = self.step == self.n - 1
        ctk.CTkButton(
            btn_row,
            text="계산기 시작 🚀" if is_last else "다음 →",
            width=140,
            fg_color=C["fire1"], hover_color="#CC2A2A",
            text_color=C["white"], font=("맑은 고딕", 11, "bold"),
            command=self._finish if is_last else lambda: self._show(self.step + 1),
        ).pack(side="left", padx=6)

    def _finish(self):
        self.on_done()
        self.destroy()


class MacroWindow(ctk.CTkToplevel):
    PRESET_COUNT = 3
    LINE_COUNT = 8

    def __init__(self, parent, initial_presets=None, active_preset=0, on_save=None, on_close=None):
        super().__init__(parent)
        self._on_save_cb = on_save
        self._on_close_cb = on_close
        self._presets = self._normalize_presets(initial_presets)
        self._active_preset = max(0, min(self.PRESET_COUNT - 1, int(active_preset) if isinstance(active_preset, int) else 0))
        self._entries = []
        self._preset_buttons = []
        self.title("매크로")
        self.geometry("640x510")
        self.resizable(False, False)
        self.configure(fg_color=C["bg"])
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._build()
        self.after(10, self._center_on_screen)

    def _center_on_screen(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _normalize_lines(self, lines):
        normalized = []
        src = lines if isinstance(lines, list) else []
        for i in range(self.LINE_COUNT):
            text = ""
            if i < len(src):
                text = str(src[i] or "")
            normalized.append(text)
        return normalized

    def _normalize_presets(self, presets):
        normalized = []
        src = presets if isinstance(presets, list) else []
        for i in range(self.PRESET_COUNT):
            lines = []
            if i < len(src):
                lines = src[i]
            normalized.append(self._normalize_lines(lines))
        return normalized

    def _sync_entries_to_current_preset(self):
        if not self._entries:
            return
        self._presets[self._active_preset] = [entry.get() for entry in self._entries]

    def _apply_current_preset_to_entries(self):
        if not self._entries:
            return
        lines = self._presets[self._active_preset]
        for i, entry in enumerate(self._entries):
            entry.delete(0, "end")
            if i < len(lines):
                entry.insert(0, lines[i])

    def _refresh_preset_buttons(self):
        for i, btn in enumerate(self._preset_buttons):
            if i == self._active_preset:
                btn.configure(
                    fg_color=C["fire1"],
                    hover_color="#CC2A2A",
                    text_color=C["white"],
                    border_color="#CC2A2A",
                )
            else:
                btn.configure(
                    fg_color=C["input"],
                    hover_color=C["border"],
                    text_color=C["gray"],
                    border_color=C["border"],
                )

    def _select_preset(self, idx):
        idx = max(0, min(self.PRESET_COUNT - 1, idx))
        if idx == self._active_preset:
            return
        self._sync_entries_to_current_preset()
        self._active_preset = idx
        self._apply_current_preset_to_entries()
        self._refresh_preset_buttons()

    def _save(self):
        self._sync_entries_to_current_preset()
        if callable(self._on_save_cb):
            self._on_save_cb(self._presets, self._active_preset)

    def _close(self):
        if callable(self._on_close_cb):
            self._on_close_cb()
        self.destroy()

    def _build(self):
        root = ctk.CTkFrame(self, fg_color="transparent")
        root.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            root,
            text="⚙ 매크로",
            font=("맑은 고딕", 18, "bold"),
            text_color=C["white"],
        ).pack(pady=(2, 8))

        body = ctk.CTkFrame(root, fg_color=C["card"], corner_radius=10)
        body.pack(fill="both", expand=True, padx=2)

        ctk.CTkLabel(
            body,
            text="프리셋(1~3)마다 문구 8줄 저장 가능 (F7 단축키로 전송)",
            font=("맑은 고딕", 11, "bold"),
            text_color=C["gray"],
        ).pack(anchor="w", padx=14, pady=(10, 6))

        preset_row = ctk.CTkFrame(body, fg_color="transparent")
        preset_row.pack(fill="x", padx=12, pady=(0, 4))

        ctk.CTkLabel(
            preset_row,
            text="프리셋",
            width=56,
            anchor="w",
            font=("맑은 고딕", 10, "bold"),
            text_color=C["gray"],
        ).pack(side="left", padx=(2, 6))

        for i in range(self.PRESET_COUNT):
            btn = ctk.CTkButton(
                preset_row,
                text=f"{i + 1}",
                width=44,
                height=26,
                fg_color=C["input"],
                hover_color=C["border"],
                text_color=C["gray"],
                border_width=1,
                border_color=C["border"],
                font=("Consolas", 11, "bold"),
                command=lambda idx=i: self._select_preset(idx),
            )
            btn.pack(side="left", padx=(0, 6))
            self._preset_buttons.append(btn)

        for i in range(self.LINE_COUNT):
            row = ctk.CTkFrame(body, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=4)

            ctk.CTkLabel(
                row,
                text=f"{i + 1}.",
                width=26,
                anchor="w",
                font=("맑은 고딕", 11, "bold"),
                text_color=C["gray"],
            ).pack(side="left", padx=(2, 6))

            entry = ctk.CTkEntry(
                row,
                height=30,
                fg_color=C["input"],
                border_color=C["border"],
                text_color=C["white"],
                font=("맑은 고딕", 11),
                placeholder_text="매크로 문구 입력",
            )
            entry.pack(side="left", fill="x", expand=True)
            self._entries.append(entry)

        self._apply_current_preset_to_entries()
        self._refresh_preset_buttons()

        btn_row = ctk.CTkFrame(root, fg_color="transparent")
        btn_row.pack(fill="x", pady=(10, 2))

        ctk.CTkButton(
            btn_row,
            text="저장",
            width=84,
            height=30,
            fg_color=C["fire1"],
            hover_color="#CC2A2A",
            text_color=C["white"],
            font=("맑은 고딕", 10, "bold"),
            command=self._save,
        ).pack(side="right", padx=(6, 0))

        ctk.CTkButton(
            btn_row,
            text="닫기",
            width=84,
            height=30,
            fg_color=C["input"],
            hover_color=C["border"],
            text_color=C["gray"],
            font=("맑은 고딕", 10, "bold"),
            command=self._close,
        ).pack(side="right")

# 정보 버튼/정보창 기능은 초기화되었습니다.
# 필요 시 여기부터 다시 새로 구현하면 됩니다.


class InfoWindow(ctk.CTkToplevel):
    def __init__(self, parent, on_close=None):
        super().__init__(parent)
        self._parent_ref = parent
        self._on_close_cb = on_close
        # 창 높이 = 화면 높이 - 오프셋. 숫자만 바꿔서 쉽게 조절할 수 있다.
        self._screen_height_offset = 150
        self._table_width = self._calc_table_width()
        self._page_titles = ["보스 칭호 보상", "업적 스코어제 보상 목록"]
        self._page = 0
        self._pages = []
        self.title("보스 칭호 보상")
        screen_h = self.winfo_screenheight()
        initial_h = max(680, min(screen_h - self._screen_height_offset, screen_h - 24))
        self.geometry(f"1540x{initial_h}")
        self.minsize(max(1320, self._table_width + 56), 760)
        self.resizable(True, True)
        self.configure(fg_color=C["bg"])
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._build()
        self.after(10, self._fit_to_screen)

    def _calc_table_width(self):
        # Label padding and separator padding are reflected so every row uses identical width.
        col_sum = sum(width for _, width in INFO_DB_COLUMNS)
        col_count = len(INFO_DB_COLUMNS)
        label_pad = col_count * (7 + 2)
        sep_count = max(0, col_count - 1)
        sep_with_pad = sep_count * (2 + 2 + 2)
        # ScrollableFrame viewport reserves extra pixels near the scrollbar area.
        # Add a small compensation so the table reaches the visible right edge.
        viewport_compensation = 36
        return col_sum + label_pad + sep_with_pad + viewport_compensation

    def _close(self):
        if callable(self._on_close_cb):
            self._on_close_cb()
        self.destroy()

    def _build(self):
        title_bar = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=44)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)

        self._title_label = ctk.CTkLabel(
            title_bar,
            text="보스 칭호 보상",
            font=("맑은 고딕", 16, "bold"),
            text_color=C["white"],
        )
        self._title_label.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkButton(
            title_bar,
            text="닫기",
            width=72,
            height=28,
            fg_color=C["input"],
            hover_color=C["border"],
            text_color=C["gray"],
            font=("맑은 고딕", 10, "bold"),
            command=self._close,
        ).pack(side="right", padx=12)

        self._btn_page_switch = ctk.CTkButton(
            title_bar,
            text="업적 보기 →",
            width=84,
            height=28,
            fg_color=C["fire1"],
            hover_color="#CC2A2A",
            text_color=C["white"],
            font=("맑은 고딕", 10, "bold"),
            command=self._on_page_switch,
        )
        self._btn_page_switch.pack(side="right", padx=(0, 6))

        legend_row = ctk.CTkFrame(title_bar, fg_color="transparent")
        legend_row.pack(side="left", padx=8)

        for label, color in INFO_HIGHLIGHT_LEGEND:
            item = ctk.CTkFrame(legend_row, fg_color="transparent")
            item.pack(side="left", padx=(0, 12), pady=(2, 0))

            ctk.CTkFrame(
                item,
                width=10,
                height=10,
                fg_color=color,
                corner_radius=2,
            ).pack(side="left", padx=(0, 5), pady=2)

            ctk.CTkLabel(
                item,
                text=label,
                font=("맑은 고딕", 9, "bold"),
                text_color=C["gray"],
            ).pack(side="left")

        self._content_host = ctk.CTkFrame(self, fg_color="transparent")
        self._content_host.pack(fill="both", expand=True, padx=0, pady=0)
        self._build_pages()
        self._show_page(0)

    def _build_pages(self):
        page1 = ctk.CTkFrame(self._content_host, fg_color="transparent")
        content = ctk.CTkScrollableFrame(
            page1,
            fg_color="transparent",
            scrollbar_button_color=C["border"],
            scrollbar_button_hover_color=C["dim"],
        )
        content.pack(fill="both", expand=True, padx=0, pady=0)
        self._build_table_header(content)
        for block in INFO_DB_ROWS:
            self._build_boss_block(content, block)

        page2 = ctk.CTkFrame(self._content_host, fg_color="transparent")
        self._build_score_table(page2)

        self._pages = [page1, page2]

        self._page_nav = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=24)
        self._page_nav.pack(fill="x")
        self._page_nav.pack_propagate(False)

    def _build_score_table(self, parent):
        col_width = 210
        table_width = col_width * len(SCORE_TABLE_COLUMNS)

        root = ctk.CTkFrame(parent, fg_color="transparent")
        root.pack(fill="both", expand=True, padx=0, pady=0)

        # Fixed boss header area (always visible while scrolling body rows).
        # Scrollable body is centered within viewport excluding scrollbar width.
        # Keep fixed header aligned by syncing right padding with actual scrollbar width.
        fixed_hdr_wrap = ctk.CTkFrame(root, fg_color="transparent")
        fixed_hdr_wrap.pack(fill="x", padx=(1, 0), pady=(2, 1))

        fixed_hdr = ctk.CTkFrame(
            fixed_hdr_wrap,
            fg_color="transparent",
            corner_radius=0,
            height=30,
            width=table_width,
        )
        fixed_hdr.pack(anchor="n")
        fixed_hdr.pack_propagate(False)

        for i, (name, bgc) in enumerate(SCORE_TABLE_COLUMNS):
            hdr_corner = 10
            cell = ctk.CTkFrame(
                fixed_hdr,
                fg_color=bgc,
                width=col_width,
                height=30,
                corner_radius=hdr_corner,
                border_width=1,
                border_color=C["border"],
            )
            cell.pack(side="left")
            cell.pack_propagate(False)
            ctk.CTkLabel(
                cell,
                text=name,
                font=("맑은 고딕", 12, "bold"),
                text_color=C["bg"],
            ).pack(expand=True)

        sc = ctk.CTkScrollableFrame(
            root,
            fg_color="transparent",
            scrollbar_button_color=C["border"],
            scrollbar_button_hover_color=C["dim"],
        )
        sc.pack(fill="both", expand=True, padx=0, pady=0)

        def _sync_score_header_alignment(_event=None):
            try:
                scrollbar = getattr(sc, "_scrollbar", None)
                right_pad = 0
                if scrollbar is not None:
                    right_pad = scrollbar.winfo_width()
                    if right_pad <= 1:
                        right_pad = scrollbar.winfo_reqwidth()
                fixed_hdr_wrap.pack_configure(padx=(1, max(0, right_pad)))
            except Exception:
                pass

        sc.bind("<Configure>", _sync_score_header_alignment, add="+")
        scrollbar = getattr(sc, "_scrollbar", None)
        if scrollbar is not None:
            scrollbar.bind("<Configure>", _sync_score_header_alignment, add="+")
        self.after(30, _sync_score_header_alignment)

        table_wrap = ctk.CTkFrame(sc, fg_color="transparent")
        table_wrap.pack(fill="x", padx=0, pady=(1, 1))

        grid = ctk.CTkFrame(table_wrap, fg_color="transparent", width=table_width)
        grid.pack(anchor="n")

        for score, rewards in SCORE_TABLE_ROWS:
            # Keep each score row visually consistent: fixed height for up to 3 lines.
            reward_row_height = 86

            # Score groups are visually separated like boss blocks.
            score_block = ctk.CTkFrame(
                grid,
                fg_color=C["panel"],
                corner_radius=10,
                width=table_width,
                border_width=1,
                border_color=C["border"],
            )
            score_block.pack(fill="x", pady=(0, 5))

            reward_row = ctk.CTkFrame(
                score_block,
                fg_color="transparent",
                corner_radius=0,
                height=reward_row_height,
                width=table_width,
            )
            reward_row.pack(fill="x", padx=0, pady=2)
            reward_row.pack_propagate(False)

            score_values = [f"{score}점"] * 5 + [f"{score * 5}점"]
            for i, reward in enumerate(rewards):
                cell_corner = 10
                self._pack_score_reward_cell(
                    reward_row,
                    text=reward,
                    width=col_width,
                    height=reward_row_height,
                    color=C["white"],
                    font=("맑은 고딕", 9),
                    cell_bg=C["card"],
                    score_text=score_values[i],
                    corner_radius=cell_corner,
                )

            self._add_score_vertical_guides(
                reward_row,
                col_width=col_width,
                col_count=len(SCORE_TABLE_COLUMNS),
                row_height=reward_row_height,
            )

    def _on_page_switch(self):
        if not self._pages:
            return
        self._show_page((self._page + 1) % len(self._pages))

    def _show_page(self, idx):
        for p in self._pages:
            p.pack_forget()
        self._pages[idx].pack(fill="both", expand=True)
        self._page = idx
        if 0 <= idx < len(self._page_titles):
            page_title = self._page_titles[idx]
            self.title(page_title)
            self._title_label.configure(text=page_title)
        self._rebuild_page_nav()

    def _rebuild_page_nav(self):
        for child in self._page_nav.winfo_children():
            child.destroy()

        dots = ctk.CTkFrame(self._page_nav, fg_color="transparent")
        dots.place(relx=0.5, rely=0.5, anchor="center")
        for i in range(len(self._pages)):
            color = C["fire1"] if i == self._page else C["dim"]
            ctk.CTkLabel(dots, text="●", font=("맑은 고딕", 8), text_color=color, width=16).pack(side="left")

        total_pages = len(self._pages)
        if total_pages <= 0:
            self._btn_page_switch.configure(text="업적 보기 →")
            return

        if self._page == total_pages - 1:
            self._btn_page_switch.configure(text="칭호 보기 ←")
        else:
            self._btn_page_switch.configure(text="업적 보기 →")

    def _pack_info_cell(self, parent, text, width, color, anchor, font, pady, pad_x=(7, 2)):
        segments = _split_info_special_segments(text)
        if anchor != "w":
            ctk.CTkLabel(
                parent,
                text=text,
                width=width,
                anchor=anchor,
                font=font,
                text_color=color,
                justify="left",
                wraplength=width - 8 if width > 80 else 0,
            ).pack(side="left", padx=pad_x, pady=pady)
            return

        text_value = str(text)
        line_count = max(1, str(text).count("\n") + 1)
        holder_height = max(18, (14 * line_count) + 4)
        holder = tk.Frame(parent, bg=C["card"], width=width, height=holder_height, bd=0, highlightthickness=0)
        holder.pack(side="left", padx=pad_x, pady=pady)
        holder.pack_propagate(False)

        # tk.Text + tag rendering keeps multiline flow stable while preserving keyword colors.
        body = tk.Text(
            holder,
            bg=C["card"],
            fg=color,
            font=font,
            bd=0,
            highlightthickness=0,
            relief="flat",
            wrap="word",
            padx=0,
            pady=0,
            cursor="arrow",
            takefocus=0,
        )
        body.place(x=0, y=0, relwidth=1, relheight=1)
        body.configure(spacing1=0, spacing2=0, spacing3=0)

        if not segments:
            segments = [(text_value, None)]

        color_tag_index = 0
        for segment_text, segment_color in segments:
            if not segment_text:
                continue
            if segment_color:
                tag_name = f"seg_{color_tag_index}"
                color_tag_index += 1
                body.insert("end", segment_text, tag_name)
                body.tag_configure(tag_name, foreground=segment_color)
            else:
                body.insert("end", segment_text)

        body.configure(state="disabled")

    def _pack_score_reward_cell(self, parent, text, width, height, color, font, cell_bg, score_text, corner_radius=0):
        strip_color = "#1A1F38"
        clamped_text = "\n".join(str(text).splitlines()[:3])
        cell = ctk.CTkFrame(
            parent,
            fg_color=cell_bg,
            width=width,
            height=height,
            corner_radius=corner_radius,
            border_width=1,
            border_color=C["border"],
        )
        cell.pack(side="left")
        cell.pack_propagate(False)

        score_strip = ctk.CTkFrame(cell, fg_color=strip_color, height=20, corner_radius=corner_radius)
        score_strip.pack(fill="x")
        score_strip.pack_propagate(False)
        ctk.CTkLabel(
            score_strip,
            text=score_text,
            anchor="center",
            font=("맑은 고딕", 12, "bold"),
            text_color=C["white"],
        ).pack(expand=True)

        dotted_divider = tk.Canvas(
            cell,
            height=2,
            bg=cell_bg,
            bd=0,
            highlightthickness=0,
            relief="flat",
        )
        dotted_divider.pack(fill="x")

        def _draw_dotted_divider(event, canvas=dotted_divider):
            canvas.delete("all")
            canvas.create_line(
                3,
                1,
                max(3, event.width - 3),
                1,
                fill=C["border"],
                dash=(1, 2),
            )

        dotted_divider.bind("<Configure>", _draw_dotted_divider, add="+")

        content = ctk.CTkFrame(cell, fg_color=cell_bg, corner_radius=0)
        content.pack(fill="both", expand=True)

        body = tk.Text(
            content,
            bg=cell_bg,
            fg=color,
            font=font,
            bd=0,
            highlightthickness=0,
            relief="flat",
            wrap="word",
            padx=0,
            pady=0,
            cursor="arrow",
            takefocus=0,
        )
        body.pack(fill="both", expand=True, padx=8, pady=(4, 5))
        body.configure(spacing1=0, spacing2=0, spacing3=0)

        segments = _split_info_special_segments(clamped_text)
        if not segments:
            segments = [(clamped_text, None)]

        color_tag_index = 0
        for segment_text, segment_color in segments:
            if not segment_text:
                continue
            if segment_color:
                tag_name = f"score_seg_{color_tag_index}"
                color_tag_index += 1
                body.insert("end", segment_text, tag_name)
                body.tag_configure(tag_name, foreground=segment_color)
            else:
                body.insert("end", segment_text)

        body.configure(state="disabled")

    def _add_score_vertical_guides(self, parent, col_width, col_count, row_height):
        guide_color = "#2A3154"
        for i in range(1, col_count):
            x = i * col_width
            ctk.CTkFrame(
                parent,
                fg_color=guide_color,
                width=1,
                height=max(0, row_height - 2),
                corner_radius=0,
            ).place(x=x, y=1)

    def _build_table_header(self, parent):
        row = ctk.CTkFrame(parent, fg_color=C["panel"], corner_radius=8, height=24, width=self._table_width)
        row.pack(fill="x", padx=0, pady=(0, 4))
        row.pack_propagate(False)
        for i, (title, width) in enumerate(INFO_DB_COLUMNS):
            ctk.CTkLabel(
                row,
                text=title,
                width=width,
                anchor="center",
                font=("맑은 고딕", 11, "bold"),
                text_color=C["white"],
            ).pack(side="left", padx=(7, 2), pady=0)

            if i < len(INFO_DB_COLUMNS) - 1:
                ctk.CTkFrame(
                    row,
                    width=2,
                    height=16,
                    fg_color=C["border"],
                    corner_radius=0,
                ).pack(side="left", pady=5, padx=2)

    def _build_boss_block(self, parent, block):
        boss_row = ctk.CTkFrame(parent, fg_color=block["color"], corner_radius=8, height=24, width=self._table_width)
        boss_row.pack(fill="x", padx=0, pady=(0, 1))
        boss_row.pack_propagate(False)
        ctk.CTkLabel(
            boss_row,
            text=block["boss"],
            font=("맑은 고딕", 12, "bold"),
            text_color=C["bg"],
        ).pack(side="left", padx=9, pady=3)

        for data in block["rows"]:
            row = ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=6, height=26, width=self._table_width)
            row.pack(fill="x", padx=0, pady=1)
            row.pack_propagate(False)
            cells = [
                (data[0], INFO_DB_COLUMNS[0][1], DIFF_COLOR.get(data[0], C["white"]), "center"),
                (data[1], INFO_DB_COLUMNS[1][1], C["white"], "center"),
                (data[2], INFO_DB_COLUMNS[2][1], C["gold"], "center"),
                (data[3], INFO_DB_COLUMNS[3][1], C["white"], "w"),
                (data[4], INFO_DB_COLUMNS[4][1], C["white"], "w"),
                (data[5], INFO_DB_COLUMNS[5][1], C["white"], "w"),
                (data[6], INFO_DB_COLUMNS[6][1], C["white"], "w"),
            ]
            for i, (text, width, color, anchor) in enumerate(cells):
                self._pack_info_cell(
                    row,
                    text=text,
                    width=width,
                    color=color,
                    anchor=anchor,
                    font=("맑은 고딕", 8),
                    pady=3,
                )

                if i < len(cells) - 1:
                    ctk.CTkFrame(
                        row,
                        width=2,
                        height=14,
                        fg_color=C["border"],
                        corner_radius=0,
                    ).pack(side="left", pady=4, padx=2)

    def _fit_to_screen(self):
        # 콘텐츠 크기 기준으로 창을 키우되, 화면 밖으로는 나가지 않게 제한.
        try:
            self.update_idletasks()
            req_w = self.winfo_reqwidth() + 20
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()

            min_w = max(1320, self._table_width + 56)
            width = min(max(min_w, req_w), max(min_w, screen_w - 24))
            target_h = screen_h - self._screen_height_offset
            height = max(680, min(target_h, screen_h - 24))
            x = max(0, (screen_w - width) // 2)
            y = max(0, (screen_h - height) // 2)
            self.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            pass


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  플레이어 카드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class PlayerCard(ctk.CTkFrame):
    SUB_ORDER = ["레어", "에픽", "유니크", "레전드리", "태초", "신화"]
    SUB_COLORS = {
        "레어": "#4D9FFF", "에픽": "#B06EFF",
        "유니크": "#FF7A00", "레전드리": "#FFD600",
        "태초": "#00D4C8", "신화": "#FF3D3D",
    }
    STACK_TEXT_COLORS = {
        "X": C["gray"],
        "0": C["white"],
        "1": C["white"],
        "2": C["white"],
        "3": C["white"],
        "4": C["white"],
        "5": C["white"],
    }

    def __init__(self, parent, app, idx, on_change):
        super().__init__(parent, fg_color=C["card"],
                         border_color=C["border"], border_width=1,
                         corner_radius=10)
        self.app = app
        self.idx = idx
        self.cb = on_change
        self.color = PLAYER_COLORS[idx]
        self._vars = {}
        self._sub_stack = []
        self._sub_menus = []
        self._loading = False
        self._build()
        self.load()

    def _build(self):
        p = self.app.players[self.idx]

        ctk.CTkFrame(self, height=4, corner_radius=10,
                     fg_color=self.color).pack(fill="x")

        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=12, pady=(0, 0))
        ctk.CTkLabel(hdr, text=p["name"],
                     font=("맑은 고딕", 12, "bold"),
                     text_color=self.color).pack(side="left")
        self._lbl_personal_dps = ctk.CTkLabel(
            hdr,
            text="0.00억",
            font=("Consolas", 11, "bold"),
            text_color=C["gray"],
        )
        self._lbl_personal_dps.pack(side="left", padx=(8, 0))

        r1 = ctk.CTkFrame(self, fg_color="transparent")
        r1.pack(fill="x", padx=6, pady=0)

        field_label_w = 32
        field_font = ("맑은 고딕", 11, "bold")
        field_label_color = "#B8C2E8"

        ctk.CTkLabel(r1, text="무기", width=field_label_w,
                 font=field_font, text_color=field_label_color).pack(side="left")
        self._vars["weapon"] = ctk.StringVar(value=p["weapon"])
        weapon_menu = ctk.CTkOptionMenu(
            r1, variable=self._vars["weapon"], values=WEAPONS,
            fg_color=C["input"], button_color=C["border"],
            button_hover_color=C["dim"], text_color=C["white"],
            font=("맑은 고딕", 10, "bold"), width=86,
            dynamic_resizing=False, command=self._upd
        )
        weapon_menu.pack(side="left", padx=(2, 6))

        def on_weapon_wheel(event):
            try:
                idx = WEAPONS.index(self._vars["weapon"].get())
            except ValueError:
                idx = 0

            if event.delta > 0 and idx < len(WEAPONS) - 1:
                idx += 1
            elif event.delta < 0 and idx > 0:
                idx -= 1

            self._vars["weapon"].set(WEAPONS[idx])
            self._upd()
            return "break"

        weapon_menu.bind("<MouseWheel>", on_weapon_wheel)

        def _build_numeric_input(parent_row, key, label_text, min_value=0, max_value=None):
            ctk.CTkLabel(parent_row, text=label_text, width=field_label_w,
                         font=field_font, text_color=field_label_color).pack(side="left")

            self._vars[key] = ctk.StringVar(value="0")
            entry = ctk.CTkEntry(
                parent_row, textvariable=self._vars[key], width=32,
                fg_color=C["input"], border_color=C["border"],
                text_color=C["white"], font=("Consolas", 11, "bold")
            )
            entry.pack(side="left", padx=(2, 0))
            self._vars[key].trace_add("write", lambda *_: self._upd())

            def _up(event=None):
                try:
                    val = int(self._vars[key].get())
                except Exception:
                    val = min_value
                next_val = val + 1
                if max_value is not None:
                    next_val = min(max_value, next_val)
                self._vars[key].set(str(next_val))
                return "break"

            def _down(event=None):
                try:
                    val = int(self._vars[key].get())
                except Exception:
                    val = min_value
                self._vars[key].set(str(max(min_value, val - 1)))
                return "break"

            entry.bind("<Up>", _up)
            entry.bind("<Down>", _down)
            entry.bind("<KeyPress-plus>", _up)
            entry.bind("<KeyPress-KP_Add>", _up)

            def _on_wheel(event):
                if event.delta > 0:
                    return _up(event)
                if event.delta < 0:
                    return _down(event)
                return "break"

            entry.bind("<MouseWheel>", _on_wheel)

            spin = ctk.CTkFrame(parent_row, fg_color="transparent")
            spin.pack(side="left", padx=(0, 4))
            ctk.CTkButton(spin, text="▲", width=18, height=18,
                          font=("Consolas", 10), fg_color=C["input"],
                          hover_color=C["border"], command=_up).pack(side="top")
            ctk.CTkButton(spin, text="▼", width=18, height=18,
                          font=("Consolas", 10), fg_color=C["input"],
                          hover_color=C["border"], command=_down).pack(side="bottom")

        _build_numeric_input(r1, "seal_re", "[레]인")
        _build_numeric_input(r1, "seal_shin", "[신]인")
        _build_numeric_input(r1, "seal_lvl", "공업", min_value=0, max_value=110)

        grid = ctk.CTkFrame(self, fg_color="transparent")
        grid.pack(fill="x", padx=10, pady=(0, 10))

        for col, name in enumerate(self.SUB_ORDER):
            col_f = ctk.CTkFrame(grid, fg_color="transparent")
            col_f.grid(row=0, column=col, padx=1, sticky="ew")
            grid.columnconfigure(col, weight=1)

            ctk.CTkLabel(col_f, text=name,
                         font=("맑은 고딕", 10, "bold"),
                         text_color=self.SUB_COLORS[name]).pack()

            sv = ctk.StringVar(value="X")
            om = ctk.CTkOptionMenu(
                col_f, variable=sv, values=SUB_STACKS,
                fg_color=C["fg_color"], button_color=C["border"],
                button_hover_color=C["dim"], text_color=C["white"],
                font=("Consolas", 12), width=52, command=self._upd
            )
            om.pack(fill="x")

            # 보조무기 스택: 마우스 휠로 X/0~9 변경
            def on_sub_wheel(event, var=sv):
                try:
                    idx = SUB_STACKS.index(var.get())
                except ValueError:
                    idx = 0

                # Windows: wheel up -> delta>0, down -> delta<0
                if event.delta > 0 and idx < len(SUB_STACKS) - 1:
                    idx += 1
                elif event.delta < 0 and idx > 0:
                    idx -= 1

                var.set(SUB_STACKS[idx])
                self._upd()
                return "break"

            om.bind("<MouseWheel>", on_sub_wheel)
            self._sub_stack.append(sv)
            self._sub_menus.append(om)

    def _apply_sub_stack_colors(self):
        for i, om in enumerate(self._sub_menus):
            stack = self._sub_stack[i].get()
            txt = self.STACK_TEXT_COLORS.get(stack, C["white"])
            om.configure(text_color=txt)

    def load(self):
        p = self.app.players[self.idx]
        self._loading = True
        self._vars["weapon"].set(p["weapon"])
        seal_re = _to_int(p.get("seal_re", p.get("seal", 0)), default=0, min_value=0)
        seal_shin = _to_int(p.get("seal_shin", 0), default=0, min_value=0)
        self._vars["seal_re"].set(str(seal_re))
        self._vars["seal_shin"].set(str(seal_shin))
        self._vars["seal_lvl"].set(str(min(110, max(0, int(p.get("seal_lvl", 0))))))

        for i, sv in enumerate(self._sub_stack):
            if i < len(p["subs"]):
                grade, stack = p["subs"][i]
                # 기존 저장값("없음")도 UI에서는 "X"로 보여준다.
                sv.set("X" if (grade == "없음" or str(stack) == "없음") else str(stack))
            else:
                sv.set("X")
        self._apply_sub_stack_colors()
        self._loading = False

    def _upd(self, *_):
        if self._loading:
            return

        p = self.app.players[self.idx]
        p["weapon"] = self._vars["weapon"].get()
        try:
            seal_re = max(0, int(self._vars["seal_re"].get()))
            if self._vars["seal_re"].get() != str(seal_re):
                self._loading = True
                self._vars["seal_re"].set(str(seal_re))
                self._loading = False
            p["seal_re"] = seal_re
        except ValueError:
            pass
        try:
            seal_shin = max(0, int(self._vars["seal_shin"].get()))
            if self._vars["seal_shin"].get() != str(seal_shin):
                self._loading = True
                self._vars["seal_shin"].set(str(seal_shin))
                self._loading = False
            p["seal_shin"] = seal_shin
        except ValueError:
            pass

        p["seal"] = _to_int(p.get("seal_re", 0), default=0, min_value=0) + _to_int(p.get("seal_shin", 0), default=0, min_value=0)
        try:
            lvl_val = int(self._vars["seal_lvl"].get())
            lvl_val = min(110, max(0, lvl_val))
            if self._vars["seal_lvl"].get() != str(lvl_val):
                self._loading = True
                self._vars["seal_lvl"].set(str(lvl_val))
                self._loading = False
            p["seal_lvl"] = lvl_val
        except ValueError:
            pass

        p["subs"] = []
        for i in range(6):
            stack = self._sub_stack[i].get()
            if stack == "X":
                p["subs"].append(("없음", "없음"))
            else:
                p["subs"].append((self.SUB_ORDER[i], stack))
        self._apply_sub_stack_colors()
        self.cb()

    def sync_to_model(self):
        """현재 카드 UI 값을 app.players에 즉시 반영한다."""
        self._upd()

    def set_personal_dps(self, dps_eok: float, enabled: bool = True):
        if enabled:
            self._lbl_personal_dps.configure(
                text=f"[{dps_eok:.2f}억]",
                text_color=C["gold"] if dps_eok > 0 else C["gray"],
            )
        else:
            self._lbl_personal_dps.configure(text="OFF", text_color=C["dim"])


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  메인 앱
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self._is_frozen_exe = bool(getattr(sys, "frozen", False))

        self.title("🔥 5인 보스 DPS 계산기  |  시즌2")
        window_width = 720
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f"{window_width}x{screen_height}+{max(0, screen_width - window_width)}+0")
        self.configure(fg_color=C["bg"])
        self._window_icon_image = None
        self._apply_window_icon()

        self._topmost = False
        self._drag_x = 0
        self._drag_y = 0
        self._keyboard_mod = None
        self._pyautogui_mod = None
        self._pyautogui_tried_import = False
        self._pystray_mod = None
        self._pil_image_mod = None
        self._tray_import_tried = False
        self._tray_icon = None
        self._tray_thread = None
        self._tray_image_cache = None
        self._toast_label = None
        self._toast_after_id = None
        self._macro_send_thread = None
        self._macro_send_result = None
        self._macro_send_success_toast = "보스 트라이 정보 전송 완료 "
        self._global_hotkey_registered = False
        self._global_custom_hotkey_registered = False
        self._global_window_toggle_hotkey_registered = False
        self._native_window_toggle_hotkey_registered = False
        self._native_window_toggle_hotkey_id = 0xB612
        self._zero_send_guard_until = 0.0
        self._zero_send_guard_lock = threading.Lock()
        self._last_hotkey_trigger_at = 0.0
        self._last_custom_hotkey_trigger_at = 0.0
        self._last_focus_hotkey_trigger_at = 0.0
        self._last_window_toggle_hotkey_trigger_at = 0.0
        self._refresh_job = None
        self._info_window = None
        self._macro_window = None
        self._macro_custom_presets = [[""] * 8 for _ in range(3)]
        self._macro_active_preset = 0
        self._update_check_started = False
        self._update_downloading = False

        self.players = [default_player(i) for i in range(6)]
        self.player_enabled = [True] * 6
        self._load()
        self._build_main()

        self.bind_all("<KeyRelease>", self._on_key_toggle)
        self.bind_all("<ButtonRelease-1>", self._blur_text_input_on_click, add="+")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.bind("<Escape>",
                  lambda e: self._toggle_topmost() if self._topmost else None)

        self._setup_external_macro_tools()
        self._setup_native_window_toggle_hotkey()

        # 시작 직후 위에 고정(타이틀바 제거) 모드 강제 적용.
        self.after(120, self._apply_startup_topmost_for_exe)
        self.after(350, self._apply_startup_topmost_for_exe)
        self.after(800, self._apply_startup_topmost_for_exe)
        self.after(1400, self._start_auto_update_check)

        if not SAVE_FILE.exists():
            self.after(200, self._onboard)

    def _apply_startup_topmost_for_exe(self):
        if self._topmost:
            return
        try:
            self.withdraw()
            self.overrideredirect(True)
            self.attributes("-topmost", True)
            self.deiconify()
            self._topmost = True
            self._make_draggable()
            self._btn_topmost.configure(
                fg_color=C["teal"], hover_color="#00A89E",
                text_color=C["bg"], text="📌 고정 ON",
            )
            self._hdr_sub.configure(text="✥ 드래그로 이동  │  ESC 해제", text_color=C["teal"])
        except Exception:
            pass

    def _apply_window_icon(self):
        # EXE/파이썬 실행 모두에서 창 상단 아이콘 적용
        if os.name == "nt":
            for ico_name in ("app.ico",):
                ico_path = _resource_path(ico_name)
                if not ico_path.exists():
                    continue
                try:
                    self.iconbitmap(default=str(ico_path))
                    return
                except Exception:
                    pass

        try:
            from tkinter import PhotoImage
            for png_name in ("main_image.png", "1.png", "2.png"):
                png_path = _resource_path(png_name)
                if not png_path.exists():
                    continue
                img = PhotoImage(file=str(png_path))
                self.iconphoto(True, img)
                self._window_icon_image = img
                return
        except Exception:
            pass

    def _start_auto_update_check(self):
        if self._update_check_started:
            return
        self._update_check_started = True

        # 개발 실행(main.py)에서는 자동 업데이트를 건너뛴다.
        if not self._is_frozen_exe:
            return

        manifest_url = str(UPDATE_MANIFEST_URL or "").strip()
        if not manifest_url:
            return

        def _worker():
            result = {"status": "none"}
            try:
                req = urlrequest.Request(manifest_url, headers={"Cache-Control": "no-cache"})
                with urlrequest.urlopen(req, timeout=UPDATE_CHECK_TIMEOUT_SEC) as resp:
                    payload = resp.read().decode("utf-8", errors="replace")
                data = json.loads(payload)

                remote_ver = str(data.get("version") or "").strip()
                download_url = str(data.get("url") or "").strip()
                sha256 = str(data.get("sha256") or "").strip().lower()

                if remote_ver and download_url and _is_remote_version_newer(remote_ver, APP_VERSION):
                    result = {
                        "status": "update",
                        "manifest": {
                            "version": remote_ver,
                            "url": download_url,
                            "sha256": sha256,
                            "notes": str(data.get("notes") or "").strip(),
                        }
                    }
                else:
                    result = {"status": "latest"}
            except Exception as ex:
                result = {"status": "error", "error": str(ex)}

            try:
                self.after(0, lambda: self._on_update_check_done(result))
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _on_update_check_done(self, result):
        if not isinstance(result, dict):
            return
        if result.get("status") != "update":
            return

        manifest = result.get("manifest") or {}
        remote_ver = str(manifest.get("version") or "").strip()
        if not remote_ver:
            return

        try:
            from tkinter import messagebox
            note = str(manifest.get("notes") or "").strip()
            msg = (
                f"새 버전이 있습니다.\n\n"
                f"현재: v{APP_VERSION}\n"
                f"최신: v{remote_ver}\n"
            )
            if note:
                msg += f"\n변경사항:\n{note}\n"
            msg += "\n지금 업데이트할까요?"

            do_update = messagebox.askyesno("업데이트", msg)
            if do_update:
                self._begin_update_download(manifest)
        except Exception:
            pass

    def _begin_update_download(self, manifest):
        if self._update_downloading:
            self._show_toast("이미 업데이트 다운로드 중입니다 ", C["gold"], C["bg"], 1400)
            return

        if not isinstance(manifest, dict):
            return

        download_url = str(manifest.get("url") or "").strip()
        if not download_url:
            return

        self._update_downloading = True
        self._show_toast(" 업데이트 다운로드 시작 ", C["teal"], C["bg"], 1300)

        def _worker():
            outcome = {"status": "failed", "error": "unknown"}
            tmp_dir = None
            try:
                tmp_dir = Path(tempfile.mkdtemp(prefix="bossdps_update_"))
                downloaded_exe = tmp_dir / "update_new.exe"

                req = urlrequest.Request(download_url, headers={"Cache-Control": "no-cache"})
                with urlrequest.urlopen(req, timeout=UPDATE_DOWNLOAD_TIMEOUT_SEC) as resp:
                    with downloaded_exe.open("wb") as f:
                        while True:
                            chunk = resp.read(1024 * 128)
                            if not chunk:
                                break
                            f.write(chunk)

                expected_sha256 = str(manifest.get("sha256") or "").strip().lower()
                if expected_sha256:
                    actual_sha256 = _sha256_file(downloaded_exe)
                    if actual_sha256.lower() != expected_sha256:
                        raise RuntimeError("다운로드 파일 해시가 일치하지 않습니다.")

                script_path = tmp_dir / "apply_update.ps1"
                script_path.write_text(
                    self._build_update_apply_script(downloaded_exe),
                    encoding="utf-8",
                )

                outcome = {
                    "status": "ready",
                    "script": str(script_path),
                    "version": str(manifest.get("version") or ""),
                }
            except Exception as ex:
                outcome = {"status": "failed", "error": str(ex)}

            try:
                self.after(0, lambda: self._on_update_download_done(outcome))
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _build_update_apply_script(self, downloaded_exe: Path) -> str:
        app_exe = Path(sys.executable).resolve()
        backup_exe = Path(str(app_exe) + ".bak")

        def _ps_quote(text: str) -> str:
            return str(text).replace("'", "''")

        pid_val = os.getpid()
        return (
            "$ErrorActionPreference = 'Stop'\n"
            f"$pidToWait = {pid_val}\n"
            f"$appExe = '{_ps_quote(str(app_exe))}'\n"
            f"$newExe = '{_ps_quote(str(downloaded_exe))}'\n"
            f"$backupExe = '{_ps_quote(str(backup_exe))}'\n"
            "Write-Host '=== 업데이트 스크립트 시작 ==='\n"
            "Write-Host 'appExe:' $appExe\n"
            "Write-Host 'newExe:' $newExe\n"
            "Write-Host 'backupExe:' $backupExe\n"
            "while (Get-Process -Id $pidToWait -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 400 }\n"
            "Start-Sleep -Milliseconds 300\n"
            "try {\n"
            "  Write-Host '백업/교체 시작'\n"
            "  if (Test-Path -LiteralPath $backupExe) { Remove-Item -LiteralPath $backupExe -Force -ErrorAction SilentlyContinue }\n"
            "  if (Test-Path -LiteralPath $appExe) { Move-Item -LiteralPath $appExe -Destination $backupExe -Force }\n"
            "  Move-Item -LiteralPath $newExe -Destination $appExe -Force\n"
            "  Write-Host 'Start-Process 실행'\n"
            "  Start-Process -FilePath $appExe\n"
            "  Write-Host 'Start-Process 완료'\n"
            "}\n"
            "catch {\n"
            "  Write-Host '업데이트 스크립트 오류:' $_\n"
            "  [System.Windows.Forms.MessageBox]::Show('업데이트 실패: ' + $_, '업데이트 오류', 'OK', 'Error')\n"
            "  if ((-not (Test-Path -LiteralPath $appExe)) -and (Test-Path -LiteralPath $backupExe)) {\n"
            "    Move-Item -LiteralPath $backupExe -Destination $appExe -Force\n"
            "  }\n"
            "}\n"
            "Write-Host '=== 업데이트 스크립트 종료 ==='\n"
        )

    def _on_update_download_done(self, outcome):
        self._update_downloading = False
        if not isinstance(outcome, dict):
            self._show_toast(" 업데이트 실패 ", C["fire1"], C["white"], 1800)
            return

        if outcome.get("status") != "ready":
            self._show_toast(" 업데이트 실패 ", C["fire1"], C["white"], 1800)
            return

        script_path = str(outcome.get("script") or "").strip()
        if not script_path:
            self._show_toast(" 업데이트 실패 ", C["fire1"], C["white"], 1800)
            return

        try:
            flags = 0
            if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
                flags |= subprocess.CREATE_NEW_PROCESS_GROUP
            if hasattr(subprocess, "DETACHED_PROCESS"):
                flags |= subprocess.DETACHED_PROCESS

            subprocess.Popen(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy", "Bypass",
                    "-File", script_path,
                ],
                creationflags=flags,
                close_fds=True,
            )

            self.save()
            self._show_toast(" 업데이트 적용을 위해 재시작합니다 ", C["teal"], C["bg"], 1200)
            self.after(300, self._on_close)
        except Exception:
            self._show_toast(" 업데이트 실행 실패 ", C["fire1"], C["white"], 2000)

    def _setup_external_macro_tools(self):
        try:
            self._keyboard_mod = importlib.import_module("keyboard")
        except Exception:
            self._keyboard_mod = None

        if not self._keyboard_mod:
            return

        self._global_hotkey_registered = False
        self._global_custom_hotkey_registered = False
        self._global_focus_hotkey_registered = False
        # Keep registrations minimal; too many aliases can fire duplicate callbacks.
        for hotkey_name in ("-", "num -"):
            try:
                self._keyboard_mod.add_hotkey(hotkey_name, self._on_global_hotkey_zero)
                self._global_hotkey_registered = True
            except Exception:
                pass

        # Send custom macro lines using global F7.
        for hotkey_name in ("f7",):
            try:
                self._keyboard_mod.add_hotkey(hotkey_name, self._on_global_hotkey_custom)
                self._global_custom_hotkey_registered = True
            except Exception:
                pass

        # Bring calculator to front from game window using global F6.
        for hotkey_name in ("f6",):
            try:
                self._keyboard_mod.add_hotkey(hotkey_name, self._on_global_hotkey_focus)
                self._global_focus_hotkey_registered = True
            except Exception:
                pass

        # Toggle calculator visibility using global F12.
        for hotkey_name in ("f12",):
            try:
                self._keyboard_mod.add_hotkey(hotkey_name, self._on_global_hotkey_window_toggle)
                self._global_window_toggle_hotkey_registered = True
            except Exception:
                pass

    def _setup_native_window_toggle_hotkey(self):
        if os.name != "nt":
            return

        try:
            user32 = ctypes.windll.user32
            MOD_NOREPEAT = 0x4000
            VK_F12 = 0x7B
            ok = user32.RegisterHotKey(None, self._native_window_toggle_hotkey_id, MOD_NOREPEAT, VK_F12)
            self._native_window_toggle_hotkey_registered = bool(ok)
            if self._native_window_toggle_hotkey_registered:
                self.after(120, self._poll_native_hotkeys)
        except Exception:
            self._native_window_toggle_hotkey_registered = False

    def _poll_native_hotkeys(self):
        if not self._native_window_toggle_hotkey_registered:
            return

        try:
            user32 = ctypes.windll.user32
            PM_REMOVE = 0x0001
            WM_HOTKEY = 0x0312
            msg = wintypes.MSG()

            while user32.PeekMessageW(ctypes.byref(msg), None, WM_HOTKEY, WM_HOTKEY, PM_REMOVE):
                if msg.message == WM_HOTKEY and int(msg.wParam) == self._native_window_toggle_hotkey_id:
                    self._request_window_toggle()
        except Exception:
            pass
        finally:
            try:
                if self.winfo_exists():
                    self.after(120, self._poll_native_hotkeys)
            except Exception:
                pass

    def _ensure_pyautogui(self):
        if self._pyautogui_tried_import:
            return self._pyautogui_mod
        self._pyautogui_tried_import = True
        try:
            self._pyautogui_mod = importlib.import_module("pyautogui")
        except Exception:
            self._pyautogui_mod = None
        return self._pyautogui_mod

    def _ensure_tray_modules(self):
        if self._tray_import_tried:
            return self._pystray_mod, self._pil_image_mod

        self._tray_import_tried = True
        try:
            self._pystray_mod = importlib.import_module("pystray")
            pil_mod = importlib.import_module("PIL.Image")
            self._pil_image_mod = pil_mod
        except Exception:
            self._pystray_mod = None
            self._pil_image_mod = None
        return self._pystray_mod, self._pil_image_mod

    def _create_tray_image(self):
        if self._tray_image_cache is not None:
            return self._tray_image_cache

        _, pil_image_mod = self._ensure_tray_modules()
        if not pil_image_mod:
            return None

        # Prefer packaged ico/png; fallback to generated solid icon.
        for filename in ("app.ico", "main_image.png", "1.png", "2.png"):
            p = _resource_path(filename)
            if p.exists():
                try:
                    self._tray_image_cache = pil_image_mod.open(str(p))
                    return self._tray_image_cache
                except Exception:
                    pass

        try:
            new_image = getattr(pil_image_mod, "new", None)
            if callable(new_image):
                self._tray_image_cache = new_image("RGBA", (64, 64), (255, 61, 61, 255))
                return self._tray_image_cache
        except Exception:
            pass
        return None

    def _show_tray_icon(self):
        if self._tray_icon is not None:
            return

        pystray_mod, _ = self._ensure_tray_modules()
        if not pystray_mod:
            return

        image = self._create_tray_image()
        if image is None:
            return

        def _on_restore(icon, item):
            try:
                self.after(0, self._toggle_window_visibility)
            except Exception:
                pass

        def _on_quit(icon, item):
            try:
                self.after(0, self._on_close)
            except Exception:
                pass

        menu = pystray_mod.Menu(
            pystray_mod.MenuItem("복원", _on_restore, default=True),
            pystray_mod.MenuItem("종료", _on_quit),
        )
        self._tray_icon = pystray_mod.Icon("BossDPSCalculator", image, "6인 보스 DPS 계산기", menu)

        def _run_tray():
            try:
                self._tray_icon.run()
            except Exception:
                pass

        self._tray_thread = threading.Thread(target=_run_tray, daemon=True)
        self._tray_thread.start()

    def _hide_tray_icon(self):
        icon = self._tray_icon
        self._tray_icon = None
        if icon is None:
            return
        try:
            icon.stop()
        except Exception:
            pass

    def _on_global_hotkey_zero(self):
        # keyboard callback runs in worker thread.
        now = time.monotonic()
        with self._zero_send_guard_lock:
            if (now - self._last_hotkey_trigger_at) < 0.25:
                return
            self._last_hotkey_trigger_at = now
        self._queue_macro_send_from_hotkey()

    def _on_global_hotkey_focus(self):
        # keyboard callback runs in worker thread.
        now = time.monotonic()
        if (now - self._last_focus_hotkey_trigger_at) < 0.2:
            return
        self._last_focus_hotkey_trigger_at = now

        try:
            self.after(0, self._focus_app_window)
        except Exception:
            pass

    def _on_global_hotkey_custom(self):
        # keyboard callback runs in worker thread.
        now = time.monotonic()
        if (now - self._last_custom_hotkey_trigger_at) < 0.25:
            return
        self._last_custom_hotkey_trigger_at = now
        self._queue_custom_macro_send_from_hotkey()

    def _on_global_hotkey_window_toggle(self):
        # keyboard callback runs in worker thread.
        self._request_window_toggle()

    def _request_window_toggle(self):
        now = time.monotonic()
        if (now - self._last_window_toggle_hotkey_trigger_at) < 0.25:
            return
        self._last_window_toggle_hotkey_trigger_at = now

        try:
            self.after(0, self._toggle_window_visibility)
        except Exception:
            pass

    def _toggle_window_visibility(self):
        try:
            if self.state() == "withdrawn" or not self.winfo_viewable():
                self._hide_tray_icon()
                self.deiconify()
                self.update_idletasks()
                self.attributes("-topmost", True)
                self.lift()
                self.focus_force()
                self.after(150, lambda: self.attributes("-topmost", self._topmost))
                self._show_toast(" 창 표시 ", C["teal"], C["bg"], 800)
            else:
                self.withdraw()
                self._show_tray_icon()
        except Exception:
            pass

    def _focus_app_window(self):
        try:
            self.deiconify()
            self.update_idletasks()
        except Exception:
            pass

        # Temporarily lift window priority, then restore user's topmost state.
        try:
            self.attributes("-topmost", True)
            self.lift()
            self.focus_force()
            self.after(150, lambda: self.attributes("-topmost", self._topmost))
        except Exception:
            pass

        if os.name == "nt":
            try:
                hwnd = int(self.winfo_id())
                user32 = ctypes.windll.user32

                # Soft focus-out: tap Alt only to release some game input capture.
                VK_MENU = 0x12
                KEYEVENTF_KEYUP = 0x0002
                user32.keybd_event(VK_MENU, 0, 0, 0)
                user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)

                # Bring calculator to front directly (no Alt+Tab fallback).
                user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                user32.SetForegroundWindow(hwnd)
                user32.SetActiveWindow(hwnd)
                user32.SetFocus(hwnd)
                user32.ClipCursor(None)

                # Retry foreground + clip release once more during transition.
                def _retry_focus_release():
                    try:
                        user32.SetForegroundWindow(hwnd)
                        user32.ClipCursor(None)
                    except Exception:
                        pass

                self.after(60, _retry_focus_release)
            except Exception:
                pass

        self._show_toast(" 계산기 포커스 ", C["teal"], C["bg"], 800)

    def _queue_macro_send_from_hotkey(self):
        now = time.monotonic()
        blocked = False

        with self._zero_send_guard_lock:
            if now < self._zero_send_guard_until:
                blocked = True
            else:
                # Lock until the current send routine releases the guard.
                self._zero_send_guard_until = now + 10.0

        if blocked:
            try:
                self.after(0, lambda: self._show_toast("전송중입니다 ", C["gold"], C["bg"], 900))
            except Exception:
                pass
            return

        try:
            self.after(0, self._start_macro_send_job)
        except Exception:
            with self._zero_send_guard_lock:
                self._zero_send_guard_until = 0.0

    def _queue_custom_macro_send_from_hotkey(self):
        now = time.monotonic()
        blocked = False

        with self._zero_send_guard_lock:
            if now < self._zero_send_guard_until:
                blocked = True
            else:
                self._zero_send_guard_until = now + 10.0

        if blocked:
            try:
                self.after(0, lambda: self._show_toast("전송중입니다 ", C["gold"], C["bg"], 900))
            except Exception:
                pass
            return

        try:
            self.after(0, self._start_custom_macro_send_job)
        except Exception:
            with self._zero_send_guard_lock:
                self._zero_send_guard_until = 0.0

    def _start_macro_send_job(self):
        self._sync_all_cards_to_model()
        messages = self._build_macro_chat_messages()
        first_msg = messages[0]
        self._macro_send_success_toast = "보스 트라이 정보 전송 완료 "

        try:
            self.clipboard_clear()
            self.clipboard_append(first_msg)
            self.update_idletasks()
        except Exception:
            pass

        if not self._keyboard_mod:
            self._show_toast("keyboard 미설치, 딜량은 복사됨 ", C["fire1"], C["white"], 2400)
            with self._zero_send_guard_lock:
                self._zero_send_guard_until = 0.0
            return

        self._macro_send_result = None
        self._macro_send_thread = threading.Thread(
            target=self._send_macro_chat_worker,
            args=(messages,),
            daemon=True,
        )
        self._macro_send_thread.start()
        self.after(80, self._check_macro_send_done)

    def _start_custom_macro_send_job(self):
        messages = self._build_custom_macro_messages()
        if not messages:
            self._show_toast("매크로 텍스트가 비어있습니다 ", C["gold"], C["bg"], 1500)
            with self._zero_send_guard_lock:
                self._zero_send_guard_until = 0.0
            return

        first_msg = messages[0]
        self._macro_send_success_toast = "커스텀 매크로 전송 완료 "

        try:
            self.clipboard_clear()
            self.clipboard_append(first_msg)
            self.update_idletasks()
        except Exception:
            pass

        if not self._keyboard_mod:
            self._show_toast("keyboard 미설치, 첫 줄은 복사됨 ", C["fire1"], C["white"], 2400)
            with self._zero_send_guard_lock:
                self._zero_send_guard_until = 0.0
            return

        self._macro_send_result = None
        self._macro_send_thread = threading.Thread(
            target=self._send_macro_chat_worker,
            args=(messages,),
            daemon=True,
        )
        self._macro_send_thread.start()
        self.after(80, self._check_macro_send_done)

    def _send_macro_chat_worker(self, messages):
        # Worker thread: only performs keyboard I/O and sleeps.
        open_chat_delay = 0.02
        line_type_delay = 0.002
        close_chat_delay = 0.015
        between_lines_delay = 0.03
        pyautogui_type_delay = 0.005

        try:
            for line in messages:
                self._keyboard_mod.press_and_release("enter")
                time.sleep(open_chat_delay)
                self._keyboard_mod.write(line, delay=line_type_delay, exact=True)
                time.sleep(close_chat_delay)
                self._keyboard_mod.press_and_release("enter")
                time.sleep(between_lines_delay)
            self._macro_send_result = "success"
            return
        except Exception:
            pass

        if self._ensure_pyautogui():
            try:
                for line in messages:
                    self._keyboard_mod.press_and_release("enter")
                    time.sleep(open_chat_delay)
                    self._pyautogui_mod.write(line, interval=pyautogui_type_delay)
                    time.sleep(close_chat_delay)
                    self._keyboard_mod.press_and_release("enter")
                    time.sleep(between_lines_delay)
                self._macro_send_result = "fallback_success"
                return
            except Exception:
                pass

        self._macro_send_result = "failed"

    def _check_macro_send_done(self):
        t = self._macro_send_thread
        if t is not None and t.is_alive():
            self.after(80, self._check_macro_send_done)
            return

        result = self._macro_send_result
        self._macro_send_thread = None
        self._macro_send_result = None

        if result == "success":
            self._show_toast(self._macro_send_success_toast, C["green"], C["bg"])
        elif result == "fallback_success":
            self._show_toast(self._macro_send_success_toast, C["green"], C["bg"])
        else:
            self._show_toast("자동전송 실패, 첫 줄은 복사됨 ", C["fire1"], C["white"], 2200)

        with self._zero_send_guard_lock:
            self._zero_send_guard_until = 0.0

    def _load(self):
        try:
            data = json.loads(SAVE_FILE.read_text(encoding="utf-8"))
            for i, p in enumerate(data.get("players", [])):
                if i < 6:
                    self.players[i] = _normalize_player(p, i)
            enabled = data.get("player_enabled", [True] * 6)
            if isinstance(enabled, list):
                for i in range(min(6, len(enabled))):
                    self.player_enabled[i] = bool(enabled[i])

            raw_presets = data.get("macro_custom_presets")
            if isinstance(raw_presets, list):
                normalized_presets = []
                for p_idx in range(3):
                    preset_src = raw_presets[p_idx] if p_idx < len(raw_presets) else []
                    lines = []
                    if not isinstance(preset_src, list):
                        preset_src = []
                    for line_idx in range(8):
                        text = ""
                        if line_idx < len(preset_src):
                            text = str(preset_src[line_idx] or "")
                        lines.append(text)
                    normalized_presets.append(lines)
                self._macro_custom_presets = normalized_presets
            else:
                # 하위 호환: 기존 5줄 단일 매크로는 프리셋 1로 마이그레이션.
                custom_lines = data.get("macro_custom_lines", [""] * 5)
                if isinstance(custom_lines, list):
                    migrated = [""] * 8
                    for i in range(min(8, len(custom_lines))):
                        migrated[i] = str(custom_lines[i] or "")
                    self._macro_custom_presets[0] = migrated

            self._macro_active_preset = _to_int(
                data.get("macro_active_preset", 0),
                default=0,
                min_value=0,
                max_value=2,
            )
        except Exception:
            pass

    def save(self):
        SAVE_FILE.write_text(
            json.dumps({
                "players": self.players,
                "player_enabled": self.player_enabled,
                "macro_custom_presets": self._macro_custom_presets,
                "macro_active_preset": self._macro_active_preset,
                # 하위 호환용 필드(최근 프리셋 기준)
                "macro_custom_lines": self._macro_custom_presets[self._macro_active_preset],
            }, ensure_ascii=False, indent=2),
            encoding="utf-8")

    def _onboard(self):
        OnboardWindow(self, self.save)

    def _open_title_info(self):
        if self._info_window is not None and self._info_window.winfo_exists():
            try:
                self._info_window.lift()
                self._info_window.focus_force()
            except Exception:
                pass
            return

        def _clear_info_ref():
            self._info_window = None

        self._info_window = InfoWindow(self, on_close=_clear_info_ref)

    def _open_macro_info(self):
        if self._macro_window is not None and self._macro_window.winfo_exists():
            try:
                self._macro_window.lift()
                self._macro_window.focus_force()
            except Exception:
                pass
            return

        def _clear_macro_ref():
            self._macro_window = None

        self._macro_window = MacroWindow(
            self,
            initial_presets=self._macro_custom_presets,
            active_preset=self._macro_active_preset,
            on_save=self._save_macro_custom_lines,
            on_close=_clear_macro_ref,
        )

    def _save_macro_custom_lines(self, presets, active_preset):
        normalized_presets = []
        src_presets = presets if isinstance(presets, list) else []
        for p_idx in range(3):
            preset_src = src_presets[p_idx] if p_idx < len(src_presets) else []
            if not isinstance(preset_src, list):
                preset_src = []
            normalized_lines = []
            for line_idx in range(8):
                text = ""
                if line_idx < len(preset_src):
                    text = str(preset_src[line_idx] or "")
                normalized_lines.append(text)
            normalized_presets.append(normalized_lines)

        self._macro_custom_presets = normalized_presets
        self._macro_active_preset = _to_int(active_preset, default=0, min_value=0, max_value=2)
        self.save()
        self._show_toast(f" 매크로 프리셋 {self._macro_active_preset + 1} 저장 완료 ", C["green"], C["bg"], 1200)

    def _build_main(self):
        self._hdr = ctk.CTkFrame(self, fg_color=C["panel"], corner_radius=0, height=52)
        self._hdr.pack(fill="x")
        self._hdr.pack_propagate(False)

        self._hdr_title = ctk.CTkLabel(
            self._hdr, text="",
            font=("맑은 고딕", 13, "bold"), text_color=C["white"])
        # Title label is kept only for drag binding compatibility.

        btn_r = ctk.CTkFrame(self._hdr, fg_color="transparent")
        btn_r.pack(side="left", padx=3)

        self._btn_topmost = ctk.CTkButton(
            btn_r, text="📌 위에 고정", width=80,
            fg_color=C["input"], hover_color=C["border"],
            text_color=C["gray"], font=("맑은 고딕", 10, "bold"), height=28,
            command=self._toggle_topmost,
        )
        self._btn_topmost.pack(side="left", padx=4)

        ctk.CTkFrame(btn_r, width=2, fg_color=C["border"],
                 corner_radius=0).pack(side="left", fill="y", pady=8, padx=4)

        def _set_header_btn_active(btn):
            btn.configure(
                fg_color="#2A335A",
                hover_color="#2A335A",
                text_color=C["white"],
                border_width=1,
                border_color=C["teal"],
            )

        def _set_header_btn_inactive(btn, fg_color, hover_color, text_color, border_width=1, border_color=None):
            btn.configure(
                fg_color=fg_color,
                hover_color=hover_color,
                text_color=text_color,
                border_width=border_width,
                border_color=(border_color if border_color is not None else C["border"]),
            )

        def _bind_header_btn_focus(btn, restore_inactive=None, activate_style=None):
            if restore_inactive is None:
                restore_inactive = lambda: _set_header_btn_inactive(
                    btn,
                    fg_color=C["input"],
                    hover_color=C["border"],
                    text_color=C["gray"],
                    border_width=1,
                    border_color=C["border"],
                )
            if activate_style is None:
                activate_style = lambda: _set_header_btn_active(btn)

            restore_inactive()

            def _activate(_event):
                activate_style()

            def _deactivate(_event):
                if self.focus_get() is not btn:
                    restore_inactive()

            btn.bind("<Enter>", _activate, add="+")
            btn.bind("<Leave>", _deactivate, add="+")
            btn.bind("<FocusIn>", _activate, add="+")
            btn.bind("<FocusOut>", _deactivate, add="+")

        btn_reset = ctk.CTkButton(btn_r, text="초기화", width=60, height=28,
                                  fg_color=C["input"], hover_color=C["border"],
                                  text_color=C["gray"], font=("맑은 고딕", 10, "bold"),
                                  command=self._reset)
        btn_reset.pack(side="left", padx=4)

        btn_help = ctk.CTkButton(btn_r, text="❓도움말", width=60, height=28,
                                 fg_color=C["input"], hover_color=C["border"],
                                 text_color=C["gray"], font=("맑은 고딕", 10, "bold"),
                                 command=self._onboard)
        btn_help.pack(side="left", padx=4)

        btn_info = ctk.CTkButton(btn_r, text="정보", width=60, height=28,
                                 fg_color=C["input"], hover_color=C["border"],
                                 text_color=C["gray"], font=("맑은 고딕", 10, "bold"),
                                 command=self._open_title_info)
        btn_info.pack(side="left", padx=4)

        btn_macro = ctk.CTkButton(btn_r, text="매크로", width=60, height=28,
                                  fg_color=C["input"], hover_color=C["border"],
                                  text_color=C["gray"], font=("맑은 고딕", 10, "bold"),
                                  command=self._open_macro_info)
        btn_macro.pack(side="left", padx=4)

        def _restore_topmost_button():
            if self._topmost:
                _set_header_btn_inactive(
                    self._btn_topmost,
                    fg_color=C["teal"],
                    hover_color="#00A89E",
                    text_color=C["bg"],
                    border_width=1,
                    border_color="#00A89E",
                )
            else:
                _set_header_btn_inactive(
                    self._btn_topmost,
                    fg_color=C["input"],
                    hover_color=C["border"],
                    text_color=C["gray"],
                    border_width=1,
                    border_color=C["border"],
                )

        def _activate_topmost_button():
            _set_header_btn_inactive(
                self._btn_topmost,
                fg_color="#1ECFC0",
                hover_color="#1ECFC0",
                text_color=C["bg"],
                border_width=1,
                border_color="#A8FFF7",
            )

        btn_save = ctk.CTkButton(btn_r, text="💾저장", width=60, height=28,
                                 fg_color=C["fire1"], hover_color="#CC2A2A",
                                 text_color=C["white"], font=("맑은 고딕", 10, "bold"),
                                 command=self._save_click)
        btn_save.pack(side="left", padx=4)

        def _restore_save_button():
            _set_header_btn_inactive(
                btn_save,
                fg_color=C["fire1"],
                hover_color="#CC2A2A",
                text_color=C["white"],
                border_width=1,
                border_color="#CC2A2A",
            )

        def _activate_save_button():
            _set_header_btn_inactive(
                btn_save,
                fg_color="#E13A3A",
                hover_color="#E13A3A",
                text_color=C["white"],
                border_width=1,
                border_color="#FF0000",
            )

        for top_btn in (btn_reset, btn_help, btn_info, btn_macro):
            _bind_header_btn_focus(top_btn)
        _bind_header_btn_focus(
            self._btn_topmost,
            restore_inactive=_restore_topmost_button,
            activate_style=_activate_topmost_button,
        )
        _bind_header_btn_focus(
            btn_save,
            restore_inactive=_restore_save_button,
            activate_style=_activate_save_button,
        )

        # 위에 고정 안내 문구는 저장 버튼 오른쪽에 표시.
        self._hdr_sub = ctk.CTkLabel(
            btn_r, text="",
            font=("맑은 고딕", 10), text_color=C["dim"])
        self._hdr_sub.pack(side="left", padx=(8, 0))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=3, pady=3)

        left = ctk.CTkFrame(body, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True, padx=(0,3))

        self.cards = []
        for i in range(6):
            card = PlayerCard(left, self, i, self._on_change)
            card.grid(row=i, column=0, padx=3, pady=3, sticky="nsew")
            self.cards.append(card)
        left.columnconfigure(0, weight=1)
        for r in range(6):
            left.rowconfigure(r, weight=1)

        # 저장된 활성/비활성 상태를 UI에 반영
        for i in range(6):
            self._apply_player_visibility(i)

        right = ctk.CTkFrame(body, fg_color=C["panel"], corner_radius=10, width=300)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)
        self._build_result(right)
        self._refresh()

    def _toggle_topmost(self):
        self._topmost = not self._topmost

        self.withdraw()
        self.overrideredirect(self._topmost)
        self.attributes("-topmost", self._topmost)
        self.deiconify()

        if self._topmost:
            self._make_draggable()
            self._btn_topmost.configure(
                fg_color=C["teal"], hover_color="#00A89E",
                text_color=C["bg"], text="📌 고정 ON",
            )
            self._hdr_sub.configure(text="✥ 드래그로 이동  │  ESC 해제", text_color=C["teal"])
        else:
            self._remove_draggable()
            self._btn_topmost.configure(
                fg_color=C["input"], hover_color=C["border"],
                text_color=C["gray"], text="📌 위에 고정",
            )
            self._hdr_sub.configure(text="", text_color=C["dim"])

    def _make_draggable(self):
        for w in [self._hdr, self._hdr_title, self._hdr_sub]:
            w.bind("<ButtonPress-1>", self._drag_start)
            w.bind("<B1-Motion>", self._drag_motion)
            w.configure(cursor="fleur")

    def _remove_draggable(self):
        for w in [self._hdr, self._hdr_title, self._hdr_sub]:
            w.unbind("<ButtonPress-1>")
            w.unbind("<B1-Motion>")
            w.configure(cursor="")

    def _drag_start(self, event):
        self._drag_x = event.x_root - self.winfo_x()
        self._drag_y = event.y_root - self.winfo_y()

    def _drag_motion(self, event):
        x = event.x_root - self._drag_x
        y = event.y_root - self._drag_y
        self.geometry(f"+{x}+{y}")

    def _build_result(self, parent):
        pad = dict(padx=7)
        header_pad = dict(padx=(7, 7))

        total_card = ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=8)
        total_card.pack(fill="x", **header_pad, pady=(7, 7))

        total_row = ctk.CTkFrame(total_card, fg_color="transparent")
        total_row.pack(fill="x", padx=10, pady=7)
        total_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(total_row, text="총 딜량",
                 width=62,
                 anchor="w",
                 font=("맑은 고딕", 11, "bold"),
                 text_color=C["gray"]).grid(row=0, column=0, sticky="w")
        self.lbl_total = ctk.CTkLabel(total_row, text="0.00 억",
                                      font=("Consolas", 22, "bold"),
                                      text_color=C["gold"])
        self.lbl_total.grid(row=0, column=1)

        ctk.CTkLabel(total_row, text="", width=62).grid(row=0, column=2)

        top_sep = ctk.CTkFrame(parent, height=1, fg_color=C["border"], corner_radius=0)
        top_sep.pack(fill="x", **header_pad, pady=(0, 0))

        scroll_frame = ctk.CTkScrollableFrame(
            parent, fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color=C["border"],
            scrollbar_button_hover_color=C["dim"])
        scroll_frame.pack(fill="both", expand=True, **pad, pady=(0, 0))

        def _sync_result_header_alignment(_event=None):
            try:
                scrollbar = getattr(scroll_frame, "_scrollbar", None)
                right_pad = 0
                if scrollbar is not None:
                    right_pad = scrollbar.winfo_width()
                    if right_pad <= 1:
                        right_pad = scrollbar.winfo_reqwidth()
                # Match header width to row content width: scrollbar width + row right inset(8)
                aligned_right = 7 + max(0, right_pad + 8)
                total_card.pack_configure(padx=(7, aligned_right))
                top_sep.pack_configure(padx=(7, aligned_right))
            except Exception:
                pass

        scroll_frame.bind("<Configure>", _sync_result_header_alignment, add="+")
        _scrollbar = getattr(scroll_frame, "_scrollbar", None)
        if _scrollbar is not None:
            _scrollbar.bind("<Configure>", _sync_result_header_alignment, add="+")
        self.after(30, _sync_result_header_alignment)

        self._boss_rows = []
        self._boss_group_labels = {}
        prev_group = ""
        for name, hp in BOSSES:
            diff, *rest = name.split(" ")
            group = " ".join(rest)
            if group != prev_group:
                group_label = ctk.CTkLabel(scroll_frame, text=f"  ── {group} ──",
                                           font=("맑은 고딕", 8),
                                           text_color=C["gray"])
                if prev_group == "":
                    group_label.pack(fill="x", pady=(0, 0))
                else:
                    group_label.pack(fill="x", pady=(0, 0))
                self._boss_group_labels[group] = group_label
                prev_group = group

            row = ctk.CTkFrame(scroll_frame, fg_color=C["card"], corner_radius=6)
            # Keep horizontal margins so highlight border is not visually clipped
            # by the scrollable viewport edge/scrollbar area.
            row.pack(fill="x", padx=(2, 8), pady=1)

            # Keep content inset by 1px on all sides so highlight border stays continuous.
            row_inner = ctk.CTkFrame(row, fg_color="transparent", corner_radius=0)
            row_inner.pack(fill="both", expand=True, padx=1, pady=(0.5, 1))

            lbl_diff = ctk.CTkLabel(row_inner, text=diff, width=36,
                                    font=("맑은 고딕", 9, "bold"),
                                    text_color=DIFF_COLOR.get(diff, C["gray"]),
                                    fg_color="transparent")
            lbl_diff.pack(side="left", padx=(8, 0), pady=0)
            lbl_hp = ctk.CTkLabel(row_inner, text=f"{hp:.1f}억", width=48,
                                  anchor="e",
                                  font=("Consolas", 9),
                                  text_color=C["gray"],
                                  fg_color="transparent")
            lbl_hp.pack(side="left", padx=(0, 6), pady=0)
            lbl_need = ctk.CTkLabel(row_inner, text="", anchor="center",
                                    width=52,
                                    font=("맑은 고딕", 9, "bold"),
                                    text_color=C["need_interceptor"],
                                    fg_color="transparent")
            lbl_need.pack(side="left", pady=0)
            lbl_sec = ctk.CTkLabel(row_inner, text="", width=40,
                                   font=("맑은 고딕", 9, "bold"),
                                   text_color=C["dim"],
                                   fg_color="transparent")
            lbl_sec.pack(side="left", pady=0)
            lbl_res = ctk.CTkLabel(row_inner, text="──", width=60,
                                   font=("맑은 고딕", 9, "bold"),
                                   text_color=C["dim"],
                                   fg_color="transparent")
            lbl_res.pack(side="left", padx=(0, 8), pady=0)
            self._boss_rows.append((row, group, lbl_diff, lbl_hp, lbl_res, lbl_sec, lbl_need, hp * 1e8))

    def _on_change(self):
        self._schedule_refresh()

    def _schedule_refresh(self, delay_ms: int = 25):
        if self._refresh_job is not None:
            try:
                self.after_cancel(self._refresh_job)
            except Exception:
                pass
        self._refresh_job = self.after(delay_ms, self._refresh)

    def _is_text_input_focused(self) -> bool:
        w = self.focus_get()
        if w is None:
            return False
        if isinstance(w, ctk.CTkEntry):
            return True
        return w.winfo_class() in ("Entry", "TEntry", "Text")

    def _blur_text_input_on_click(self, event):
        if not self._is_text_input_focused():
            return

        w = event.widget
        if isinstance(w, ctk.CTkEntry):
            return
        if w is not None and w.winfo_class() in ("Entry", "TEntry", "Text"):
            return

        # Keep button/menu click behavior intact, then drop focus from Entry.
        self.after_idle(self.focus_set)

    def _on_key_toggle(self, event):
        key = event.keysym
        if key == "F12":
            # Local F12 can always hide the window, but restore requires global hook.
            # If global hook is unavailable, block hiding to avoid "window lost" situations.
            is_hidden = (self.state() == "withdrawn") or (not self.winfo_viewable())
            can_restore_with_hotkey = (
                self._global_window_toggle_hotkey_registered
                or self._native_window_toggle_hotkey_registered
            )
            if is_hidden:
                self._request_window_toggle()
            else:
                if not can_restore_with_hotkey:
                    self._show_toast("F12 전역단축키 비활성: 숨김 차단", C["gold"], C["bg"], 1800)
                    return
                self._request_window_toggle()
            return

        if key == "F7":
            # Avoid double-send when global hotkey is already active.
            if self._global_custom_hotkey_registered:
                return
            self._queue_custom_macro_send_from_hotkey()
            return

        if key in ("minus", "KP_Subtract"):
            # Avoid double-send when global hotkey is already active.
            if self._global_hotkey_registered:
                return
            self._queue_macro_send_from_hotkey()
            return

        if self._is_text_input_focused():
            return

        key_to_idx = {
            "1": 0, "2": 1, "3": 2, "4": 3, "5": 4, "6": 5,
            "KP_1": 0, "KP_2": 1, "KP_3": 2,
            "KP_4": 3, "KP_5": 4, "KP_6": 5,
        }
        if key not in key_to_idx:
            return

        self._toggle_player(key_to_idx[key])

    def _apply_player_visibility(self, idx: int):
        card = self.cards[idx]
        if self.player_enabled[idx]:
            card.grid()
        else:
            card.grid_remove()

    def _toggle_player(self, idx: int):
        self.player_enabled[idx] = not self.player_enabled[idx]
        self._apply_player_visibility(idx)
        state = "ON" if self.player_enabled[idx] else "OFF"
        self._show_toast(f" {idx + 1}P {state} ", C["teal"], C["bg"])
        self._refresh()

    def _calc_total_eok(self) -> float:
        return sum(
            calc_dps(
                p["weapon"],
                _to_int(p.get("seal_re", p.get("seal", 0)), default=0, min_value=0),
                _to_int(p.get("seal_shin", 0), default=0, min_value=0),
                p["seal_lvl"],
                p["subs"],
            ) / 1e8
            for i, p in enumerate(self.players)
            if self.player_enabled[i]
        )

    def _format_macro_chat_message(self) -> str:
        return f"현재 총 딜량: {self._calc_total_eok():.2f}억"

    def _estimate_interceptors_for_lack(self, lack_eok: float):
        enabled_count = sum(1 for enabled in self.player_enabled if enabled)
        if enabled_count <= 0:
            return None

        # Fixed assumption: interceptor seal level is always 110.
        gain_per_interceptor = ((9900 + 495 * 110) * 2.67) * 180
        if gain_per_interceptor <= 0:
            return None

        lack_raw = max(0.0, lack_eok) * 1e8
        return max(0, int(math.ceil(lack_raw / gain_per_interceptor)))

    def _build_macro_chat_messages(self):
        personal_info = []
        total = 0.0
        for i, p in enumerate(self.players):
            seal_re = _to_int(p.get("seal_re", p.get("seal", 0)), default=0, min_value=0)
            seal_shin = _to_int(p.get("seal_shin", 0), default=0, min_value=0)
            dps_eok = calc_dps(p["weapon"], seal_re, seal_shin, p["seal_lvl"], p["subs"]) / 1e8
            personal_info.append((
                self.player_enabled[i],
                p.get("weapon", "없음"),
                seal_re + seal_shin,
            ))
            if self.player_enabled[i]:
                total += dps_eok

        ordered = sorted(BOSSES, key=lambda x: x[1])

        current_name = None
        current_hp = 0.0
        next_name = None
        next_hp = 0.0

        for name, hp in ordered:
            if total >= hp:
                current_name = name
                current_hp = hp
            elif next_name is None:
                next_name = name
                next_hp = hp
                break

        lines = [f"현재 총 딜량: {total:.2f}억"]

        weapon_alias = {
            "배틀(후원)": "배틀(후)",
            "뮤탈(후원)": "뮤탈(후)",
            "다칸(후원)": "다칸(후)",
        }
        weapon_count = {}
        seal_values = []
        for enabled, weapon, seal in personal_info:
            if not enabled:
                continue
            shown_weapon = weapon_alias.get(str(weapon), str(weapon))
            weapon_count[shown_weapon] = weapon_count.get(shown_weapon, 0) + 1
            seal_values.append(str(seal))

        if weapon_count:
            weapon_parts = [f"{name} {cnt}" for name, cnt in weapon_count.items()]
            lines.append(f"무기 - {' '.join(weapon_parts)}")
            lines.append(f"인장 - {' '.join(seal_values)}")
        else:
            lines.append("무기 - 없음")
            lines.append("인장 - 없음")

        if current_name is None:
            lines.append("[현재] 가능 보스: 없음")
        else:
            lines.append(f"[현재] {current_name} ({current_hp:.1f}억)")

        if next_name is None:
            lines.append("[다음] 최종 보스까지 트라이 가능")
        else:
            lack = max(0.0, next_hp - total)
            need_interceptors = self._estimate_interceptors_for_lack(lack)
            if need_interceptors is not None:
                lines.append(f"[다음] {next_name}까지 {lack:.2f}억 부족 - 인셉 {need_interceptors}개 필요")
            else:
                lines.append(f"[다음] {next_name}까지 {lack:.2f}억 부족")

        return lines

    def _build_custom_macro_messages(self):
        lines = []
        presets = self._macro_custom_presets if isinstance(self._macro_custom_presets, list) else []
        preset_idx = _to_int(self._macro_active_preset, default=0, min_value=0, max_value=2)
        current_lines = presets[preset_idx] if preset_idx < len(presets) and isinstance(presets[preset_idx], list) else []
        for raw in current_lines:
            text = str(raw or "").strip()
            if text:
                lines.append(text)
        return lines

    def _sync_all_cards_to_model(self):
        # Global hotkey can fire right after UI input change; force latest values.
        for card in self.cards:
            card.sync_to_model()

    def _send_macro_chat_current_dps(self):
        # Backward compatibility path used by existing direct callers.
        self._start_macro_send_job()

    def _refresh(self):
        self._refresh_job = None

        if not hasattr(self, "lbl_total") or not hasattr(self, "_boss_rows"):
            return

        total = 0.0
        for i, p in enumerate(self.players):
            dps_eok = calc_dps(
                p["weapon"],
                _to_int(p.get("seal_re", p.get("seal", 0)), default=0, min_value=0),
                _to_int(p.get("seal_shin", 0), default=0, min_value=0),
                p["seal_lvl"],
                p["subs"],
            ) / 1e8
            if self.player_enabled[i]:
                total += dps_eok
            self.cards[i].set_personal_dps(dps_eok, self.player_enabled[i])

        self.lbl_total.configure(
            text=f"{total:.2f} 억",
            text_color=C["green"] if total >= BOSSES[0][1] else C["gold"]
        )

        total_raw = total * 1e8
        dps_s = total_raw / 180 if total_raw else 0
        next_boss_hp = None
        max_possible_hp = None
        max_possible_group = None
        for _, group, _, _, _, _, _, hp in self._boss_rows:
            if total_raw < hp:
                if next_boss_hp is None or hp < next_boss_hp:
                    next_boss_hp = hp
            else:
                if max_possible_hp is None or hp > max_possible_hp:
                    max_possible_hp = hp
                    max_possible_group = group

        for group_name, group_label in self._boss_group_labels.items():
            if max_possible_group is not None and group_name == max_possible_group:
                group_label.configure(text_color=C["boss_highlight"], font=("맑은 고딕", 9, "bold"))
            else:
                group_label.configure(text_color=C["gray"], font=("맑은 고딕", 8))

        next_boss_hint_shown = False
        for row, _, lbl_diff, lbl_hp, lbl_res, lbl_sec, lbl_need, hp in self._boss_rows:
            if (max_possible_hp is not None) and (hp == max_possible_hp):
                row.configure(fg_color="#1F2B52", border_width=2, border_color=C["teal"], corner_radius=4)
                lbl_diff.configure(fg_color="#1F2B52")
                lbl_hp.configure(fg_color="#1F2B52")
                lbl_need.configure(fg_color="#1F2B52")
                lbl_sec.configure(fg_color="#1F2B52")
                lbl_res.configure(fg_color="#1F2B52")
            else:
                row.configure(fg_color=C["card"], border_width=0, corner_radius=6)
                lbl_diff.configure(fg_color="transparent")
                lbl_hp.configure(fg_color="transparent")
                lbl_need.configure(fg_color="transparent")
                lbl_sec.configure(fg_color="transparent")
                lbl_res.configure(fg_color="transparent")

            if total_raw >= hp:
                lbl_res.configure(text="✅ 가능", text_color=C["green"])
                secs = hp / dps_s if dps_s else 0
                lbl_sec.configure(text=f"{secs:.0f}초", text_color=C["gray"])
                lbl_need.configure(text="")
            else:
                lbl_res.configure(text="❌ 불가", text_color=C["fire1"])
                lbl_sec.configure(text="")
                if (not next_boss_hint_shown) and (next_boss_hp is not None) and (hp == next_boss_hp):
                    lack_eok = (hp - total_raw) / 1e8
                    need_interceptors = self._estimate_interceptors_for_lack(lack_eok)
                    if need_interceptors is not None:
                        lbl_need.configure(text=f"셉{need_interceptors}개", text_color=C["need_interceptor"])
                    else:
                        lbl_need.configure(text="")
                    next_boss_hint_shown = True
                else:
                    lbl_need.configure(text="")

    def _save_click(self):
        self.save()
        self._show_toast(" 💾 저장되었습니다 ", C["green"], C["bg"])

    def _reset(self):
        from tkinter import messagebox
        if messagebox.askyesno("초기화", "모든 데이터를 초기화할까요?"):
            for i in range(6):
                self.players[i] = default_player(i)
                self.player_enabled[i] = True
            for card in self.cards:
                card.load()
            for i in range(6):
                self._apply_player_visibility(i)
            self._refresh()

    def _show_toast(self, text: str, bg_color: str, text_color: str, duration_ms: int = 1800):
        if self._toast_label is None or not self._toast_label.winfo_exists():
            self._toast_label = ctk.CTkLabel(
                self,
                text=text,
                fg_color=bg_color,
                text_color=text_color,
                font=("맑은 고딕", 10, "bold"),
                corner_radius=8,
            )
            self._toast_label.place(relx=0.5, rely=0.96, anchor="center")
        else:
            self._toast_label.configure(text=text, fg_color=bg_color, text_color=text_color)
            self._toast_label.place(relx=0.5, rely=0.96, anchor="center")

        if self._toast_after_id is not None:
            try:
                self.after_cancel(self._toast_after_id)
            except Exception:
                pass

        def _hide_toast():
            if self._toast_label is not None and self._toast_label.winfo_exists():
                self._toast_label.place_forget()
            self._toast_after_id = None

        self._toast_after_id = self.after(duration_ms, _hide_toast)

    def _on_close(self):
        self._hide_tray_icon()

        if self._native_window_toggle_hotkey_registered and os.name == "nt":
            try:
                ctypes.windll.user32.UnregisterHotKey(None, self._native_window_toggle_hotkey_id)
            except Exception:
                pass

        if self._keyboard_mod and self._global_hotkey_registered:
            try:
                self._keyboard_mod.clear_all_hotkeys()
            except Exception:
                pass
        self.destroy()


if __name__ == "__main__":
    App().mainloop()


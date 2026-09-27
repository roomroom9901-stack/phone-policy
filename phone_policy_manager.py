#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
폰정책·정산 매니저 v2 (단일 파일)
------------------------------------------------------------
실행: 이 파일을 더블클릭 (또는 python phone_policy_manager.py)
      처음 실행 때 필요한 구성요소(PySide6, openpyxl, xlrd)를 자동으로 설치합니다.

기능
  - 대리점별 정책(무선/유선) 관리, 날짜별 변경 이력, 대리점 비교
  - 대리점 정책표(엑셀·사진·PDF·캡처)를 올리면 AI(Claude)가 자동으로 표로 옮김
  - 판매가·마진 계산, 고객용 가격표(HTML) 출력
  - 개통 등록, 대리점 입금 대조, 직원 수당, 월별 순이익 리포트
  - 프로그램 안에 '사용법' 탭 포함

데이터는 이 파일과 같은 폴더의 phone_policy.db 에 저장됩니다.
"""
import sys
import os
import csv
import json
import re
import base64
import shutil
import sqlite3
import datetime
import html
import subprocess
import socket
import urllib.request
import urllib.error
import glob
import zipfile
import hashlib
from concurrent.futures import ThreadPoolExecutor


# ------------------------------------------------------------
# 오류가 나면 '아무 반응 없음'이 되지 않도록: 기록 파일 + 윈도우 알림창
# ------------------------------------------------------------
_LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "PhonePolicyManager")
try:
    os.makedirs(_LOG_DIR, exist_ok=True)
except Exception:
    _LOG_DIR = os.path.expanduser("~")
ERROR_LOG = os.path.join(_LOG_DIR, "error_log.txt")
START_LOG = os.path.join(_LOG_DIR, "start_log.txt")


def _winbox(text, title="폰정책매니저", icon=0x10):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, str(text)[:1800], title, icon)
    except Exception:
        try:
            sys.stderr.write(str(text) + "\n")
        except Exception:
            pass


def _log(path, text):
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {text}\n")
    except Exception:
        pass


def step(msg):
    """시작 단계 기록 (멈추는 곳을 찾기 위해)"""
    _log(START_LOG, msg)


def fatal(title, exc_text):
    _log(ERROR_LOG, f"{title}\n{exc_text}")
    _winbox(f"{title}\n\n{exc_text[-1200:]}\n\n(기록 파일: {ERROR_LOG})\n이 창을 캡처해서 보내주세요.")


def _excepthook(t, v, tb):
    import traceback
    txt = "".join(traceback.format_exception(t, v, tb))
    _log(ERROR_LOG, "처리 안 된 오류\n" + txt)
    try:
        from PySide6.QtWidgets import QApplication as _QA, QMessageBox as _QM
        if _QA.instance():
            _QM.warning(None, "폰정책매니저 오류", f"오류가 났습니다. 이 창을 캡처해서 보내주세요.\n\n{txt[-1000:]}")
            return
    except Exception:
        pass
    _winbox("오류가 났습니다.\n\n" + txt[-1200:])


sys.excepthook = _excepthook


def _ensure_packages():
    missing = []
    for mod, pkg in (("PySide6", "PySide6"), ("openpyxl", "openpyxl"), ("xlrd", "xlrd")):
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    if not missing:
        return
    if sys.stdout is None:          # 바탕화면 아이콘(창 없는 실행)으로 켰을 때
        _winbox("필요한 구성요소(" + ", ".join(missing) + ")를 설치합니다.\n확인을 누르고 1~3분 기다려 주세요.",
                icon=0x40)
        py = os.path.join(os.path.dirname(sys.executable), "python.exe")
        try:
            subprocess.check_call([py if os.path.isfile(py) else sys.executable, "-m", "pip", "install", *missing],
                                  creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            return
        except Exception as ex:
            fatal("구성요소 설치 실패 — 폰정책매니저_설치.bat 을 다시 더블클릭해 주세요.", str(ex))
            sys.exit(1)
    print("=" * 60)
    print(" 처음 실행이라 필요한 구성요소를 설치합니다. (1~3분)")
    print(" 이 창을 닫지 말고 기다려 주세요...")
    print("=" * 60)
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
    except Exception as ex:
        print("\n자동 설치에 실패했습니다:", ex)
        print("명령 프롬프트(cmd)를 열고 아래를 입력해 주세요:")
        print("   pip install " + " ".join(missing))
        input("\n엔터를 누르면 종료합니다.")
        sys.exit(1)


FROZEN = getattr(sys, "frozen", False)          # 설치형(exe)으로 실행 중인지

if __name__ == "__main__" and not FROZEN:
    _ensure_packages()

step(f"시작: python {sys.version.split()[0]} / {sys.executable} / {os.path.abspath(sys.argv[0]) if sys.argv else ''}")
try:
    from PySide6.QtCore import Qt, QDate, QUrl, QTimer, QThread, Signal, QBuffer, QByteArray, QIODevice
    from PySide6.QtGui import QColor, QFont, QDesktopServices, QBrush, QImage, QKeySequence, QShortcut, QIcon, QPixmap, \
        QTextDocument, QPainter
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
        QGridLayout, QGroupBox, QLabel, QPushButton, QComboBox, QLineEdit, QSpinBox,
        QDateEdit, QTableWidget, QTableWidgetItem, QAbstractItemView, QMessageBox,
        QFileDialog, QSplitter, QInputDialog, QListWidget, QTextBrowser, QCheckBox, QDialog, QPlainTextEdit,
    )
except Exception as _ex:
    import traceback as _tb
    fatal("화면 구성요소(PySide6)를 불러오지 못했습니다.\n폰정책매니저_설치.bat 을 다시 더블클릭해 주세요.\n"
          f"(파이썬 {sys.version.split()[0]})", _tb.format_exc())
    sys.exit(1)
step("PySide6 불러오기 완료")

APP_NAME = "폰정책·정산 매니저"
APP_VERSION = "2026.09.26.3"     # ← 새 버전을 배포할 때 이 숫자만 올리면 직원 PC에 업데이트 창이 뜹니다
UPDATE_REPO = "roomroom9901-stack/phone-policy"                  # 인터넷 업데이트 저장소 (사장님이 [업데이트 배포] 누르면 자동으로 채워짐)
DEFAULT_ROLE = "직원"            # 직원용 설치파일·배포본에는 "직원"으로 바뀌어 들어감 (직원 PC는 처음부터 직원 화면)
SHARE_POLICY_FILE = "정책공유.json"
SHARE_SALES_PREFIX = "개통_"
EXE_NAME = "PhonePolicyManager.exe"
SETUP_NAME = "폰정책매니저-Setup.exe"
SHORTCUT_NAME = "폰정책매니저"
_LOCAL = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
INSTALL_DIR = os.path.join(_LOCAL, "Programs", "PhonePolicyManager")   # 설치형: 프로그램 위치 (관리자 권한 불필요)
INSTALLED_EXE = os.path.join(INSTALL_DIR, EXE_NAME)
SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0] or __file__))
if FROZEN or os.name == "nt":
    # 윈도우에서는 설치형이든 파일 실행이든 데이터는 항상 같은 곳에 (파일을 어디서 받든 데이터가 이어짐)
    BASE_DIR = os.path.join(_LOCAL, "PhonePolicyManager")
    os.makedirs(os.path.join(BASE_DIR, "app"), exist_ok=True)
    _old_db = os.path.join(SCRIPT_DIR, "phone_policy.db")
    if not FROZEN and os.path.isfile(_old_db) and not os.path.isfile(os.path.join(BASE_DIR, "phone_policy.db")):
        shutil.copyfile(_old_db, os.path.join(BASE_DIR, "phone_policy.db"))      # 예전 데이터 이사
else:
    BASE_DIR = SCRIPT_DIR
APP_PY = os.path.join(BASE_DIR, "app", "phone_policy_manager.py")      # 설치형: 업데이트되는 프로그램 본체
DB_PATH = os.path.join(BASE_DIR, "phone_policy.db")

CARRIERS = ["SKT", "KT", "LGU+"]
CATEGORIES = ["무선", "유선"]
JOIN_BY_CAT = {"무선": ["신규", "번호이동", "기기변경"], "유선": ["신규", "재약정", "전환"]}
ALL_JOINS = ["신규", "번호이동", "기기변경", "재약정", "전환"]
DISCOUNT_TYPES = ["공시지원", "선택약정"]

POLICY_COLS = ["구분", "통신사", "모델/상품", "요금제/약정", "가입유형", "출고가", "공시지원금", "리베이트", "차감", "조건/메모", "그룹"]
MONEY_COLS = (5, 6, 7, 8)

# 손님 상황 체크 항목 (부가·차감 규칙의 '조건')
FLAGS = [
    ("부가", "부가서비스 가입"),
    ("보험", "폰보험 가입"),
    ("필링", "컬러링·음악 등 콘텐츠 가입"),
    ("결합", "유무선 결합"),
    ("카드", "제휴카드"),
    ("시니어", "시니어·청소년·키즈 요금제"),
    ("지원금소액", "유통망지원금 10만원 미만 입력"),
    ("단기기변", "기변: 이전 폰 13개월 미만 사용"),
    ("초단기기변", "기변: 이전 폰 183일 미만 사용"),
]
FLAG_KEYS = [k for k, _ in FLAGS]
RULE_KINDS = ["자동", "주의"] + FLAG_KEYS
RULE_COLS = ["규칙 이름", "조건", "해당 시", "미해당 시", "적용 그룹", "모델 포함", "모델 제외", "가입유형",
             "요금 최소(천원)", "요금 최대(천원)", "메모"]
RULE_MARK = "### 부가·차감 규칙 ###"
MONEY_FIELDS = ("release_price", "public_subsidy", "rebate", "deduction")

DEFAULT_MODEL = "claude-sonnet-5"
# AI 키 자동 감지: Claude(sk-ant-) / OpenAI GPT(sk-) / Gemini(AIza) — 직원마다 다른 키를 넣어도 알아서 맞는 AI 사용
AI_DEFAULTS = {"claude": "claude-sonnet-5", "openai": "gpt-5", "gemini": "gemini-2.5-flash"}
AI_NAMES = {"claude": "Claude", "openai": "OpenAI GPT", "gemini": "Google Gemini"}


def ai_provider(key):
    k = (key or "").strip()
    if k.startswith("sk-ant-"):
        return "claude"
    if k.startswith("AIza"):
        return "gemini"
    if k.startswith("sk-"):
        return "openai"
    return ""


def model_family(model):
    m = (model or "").lower()
    if m.startswith("claude"):
        return "claude"
    if m.startswith("gemini"):
        return "gemini"
    if m.startswith(("gpt", "o1", "o3", "o4", "chatgpt")):
        return "openai"
    return ""


def pick_model(key, model):
    """키 종류에 맞는 모델. 적어둔 모델이 다른 회사 것이면 그 회사 기본 모델로"""
    prov = ai_provider(key) or "claude"
    m = (model or "").strip()
    return m if m and model_family(m) == prov else AI_DEFAULTS[prov]
API_URL = "https://api.anthropic.com/v1/messages"
IMG_EXTS = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif")
SHEET_EXTS = (".xlsx", ".xlsm", ".xls", ".csv")

GREEN = QColor("#e3f6e9")
YELLOW = QColor("#fff4cc")
RED = QColor("#fde0de")
BLUE = QColor("#e4eefc")

SCHEMA = """
CREATE TABLE IF NOT EXISTS agencies(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT UNIQUE NOT NULL, carrier TEXT DEFAULT '', contact TEXT DEFAULT '', memo TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS policies(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  agency_id INTEGER NOT NULL, category TEXT NOT NULL DEFAULT '무선', carrier TEXT NOT NULL, model TEXT NOT NULL,
  plan TEXT NOT NULL DEFAULT '', join_type TEXT NOT NULL,
  release_price INTEGER DEFAULT 0, public_subsidy INTEGER DEFAULT 0,
  rebate INTEGER DEFAULT 0, deduction INTEGER DEFAULT 0, conditions TEXT DEFAULT '',
  valid_from TEXT NOT NULL, deleted INTEGER DEFAULT 0, created_at TEXT);
CREATE TABLE IF NOT EXISTS staff(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT UNIQUE NOT NULL, per_unit INTEGER DEFAULT 0, margin_rate REAL DEFAULT 0, active INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS sales(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sale_date TEXT NOT NULL, staff_id INTEGER, agency_id INTEGER, category TEXT NOT NULL DEFAULT '무선',
  carrier TEXT, model TEXT, plan TEXT, join_type TEXT, discount_type TEXT,
  customer TEXT DEFAULT '', phone4 TEXT DEFAULT '',
  release_price INTEGER DEFAULT 0, public_subsidy INTEGER DEFAULT 0,
  rebate INTEGER DEFAULT 0, deduction INTEGER DEFAULT 0,
  support INTEGER DEFAULT 0, extra_income INTEGER DEFAULT 0, margin INTEGER DEFAULT 0,
  paid INTEGER, paid_date TEXT, memo TEXT DEFAULT '', created_at TEXT);
CREATE TABLE IF NOT EXISTS expenses(
  id INTEGER PRIMARY KEY AUTOINCREMENT, month TEXT NOT NULL, item TEXT, amount INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS ai_cache(hash TEXT PRIMARY KEY, created TEXT, data TEXT);
CREATE TABLE IF NOT EXISTS sale_terms(id INTEGER PRIMARY KEY AUTOINCREMENT, sale_id INTEGER, kind TEXT, name TEXT,
  days INTEGER, due TEXT, done INTEGER DEFAULT 0, done_at TEXT);
CREATE INDEX IF NOT EXISTS idx_terms_due ON sale_terms(due);
CREATE TABLE IF NOT EXISTS policy_files(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, agency_id INTEGER, name TEXT,
  path TEXT, created TEXT);
CREATE TABLE IF NOT EXISTS model_meta(model TEXT PRIMARY KEY, qty INTEGER, hidden INTEGER DEFAULT 0, fav INTEGER DEFAULT 0);
DELETE FROM ai_cache WHERE data LIKE '[[], %';
CREATE TABLE IF NOT EXISTS rule_sets(
  id INTEGER PRIMARY KEY AUTOINCREMENT, agency_id INTEGER NOT NULL, valid_from TEXT NOT NULL, created_at TEXT);
CREATE TABLE IF NOT EXISTS rules(
  id INTEGER PRIMARY KEY AUTOINCREMENT, set_id INTEGER NOT NULL, name TEXT, kind TEXT,
  amt_yes INTEGER DEFAULT 0, amt_no INTEGER DEFAULT 0, grp TEXT DEFAULT '', model_kw TEXT DEFAULT '',
  model_ex TEXT DEFAULT '', joins TEXT DEFAULT '', fee_min INTEGER, fee_max INTEGER, memo TEXT DEFAULT '');
"""

STYLE = """
QWidget { font-size: 10pt; }
QPushButton { padding: 6px 12px; border: 1px solid #c3ccd8; border-radius: 5px; background: #ffffff; color: #222; }
QPushButton:hover { background: #eaf1fb; }
QPushButton:disabled { color: #999; background: #f2f2f2; }
QPushButton[primary="true"] { background: #1f5fbf; color: white; border: none; font-weight: bold; }
QPushButton[primary="true"]:hover { background: #174a96; }
QPushButton[big="true"] { font-size: 12pt; padding: 10px 22px; }
QTabBar::tab { padding: 9px 14px; }
QTabBar::tab:selected { font-weight: bold; }
QHeaderView::section { background: #eef1f5; padding: 4px 6px; border: none;
    border-right: 1px solid #d9dee5; border-bottom: 1px solid #d9dee5; font-weight: bold; }
QGroupBox { font-weight: bold; border: 1px solid #d9dee5; border-radius: 6px; margin-top: 12px; padding-top: 10px; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
QTableWidget { gridline-color: #e3e7ec; }
QLabel#banner { background: #eef5ff; border: 1px solid #cfe0f7; border-radius: 6px; padding: 8px 12px; color: #1d3b66; }
"""


# ============================================================
# 공통 유틸
# ============================================================
def to_num(v):
    if v is None:
        return 0.0
    s = str(v).replace(",", "").replace("원", "").replace(" ", "").strip()
    if s in ("", "-", "None"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_int(v):
    return int(round(to_num(v)))


def money_auto(v):
    """1,000 미만 숫자는 만원 단위로 보고 원으로 환산"""
    x = to_num(v)
    if 0 < abs(x) < 1000:
        x *= 10000
    return int(round(x))


def won(n):
    return f"{int(n or 0):,}"


def signed(n):
    return f"{int(n):+,}" if n else "0"


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today():
    return datetime.date.today().isoformat()


def norm_category(s):
    s = (s or "").strip().upper()
    if not s:
        return ""
    if any(k in s for k in ("유선", "인터넷", "TV", "IPTV", "홈", "집전화", "WIRED")):
        return "유선"
    return "무선"


def norm_carrier(s):
    s = (s or "").strip().upper().replace(" ", "")
    if not s:
        return ""
    if s.startswith("SK") or s.startswith("에스케이"):
        return "SKT"
    if s.startswith("KT") or s.startswith("케이티"):
        return "KT"
    if s.startswith("LG") or s.startswith("U+") or "유플" in s or "엘지" in s:
        return "LGU+"
    return ""


def norm_join(s, category="무선"):
    s = (s or "").strip().replace(" ", "")
    if not s:
        return ""
    if s.startswith("신규") or s.startswith("010") or s == "10":
        return "신규"
    if category == "유선":
        if "재약정" in s or "재가입" in s or "갱신" in s:
            return "재약정"
        if "전환" in s or "타사" in s or "이동" in s or "번이" in s:
            return "전환"
        return ""
    if "번호" in s or "번이" in s or "MNP" in s.upper() or "이동" in s:
        return "번호이동"
    if "기변" in s or "기기" in s or "변경" in s or "재가입" in s or "재약정" in s:
        return "기기변경"
    return ""


def norm_date(v):
    s = str(v or "").strip()
    if not s:
        return ""
    s = s.split(" ")[0].split("T")[0]
    digits = re.sub(r"[^0-9]", "", s)
    if len(digits) == 8:
        return f"{digits[:4]}-{digits[4:6]}-{digits[6:]}"
    parts = re.split(r"[-./]", s)
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        y, m, d = parts
        if len(y) == 2:
            y = "20" + y
        try:
            return datetime.date(int(y), int(m), int(d)).isoformat()
        except ValueError:
            return ""
    return ""


def last4(v):
    d = re.sub(r"[^0-9]", "", str(v or ""))
    return d[-4:] if len(d) >= 4 else d


def order_key(p):
    cat, c, j = p["category"], p["carrier"], p["join_type"]
    return (CATEGORIES.index(cat) if cat in CATEGORIES else 9,
            CARRIERS.index(c) if c in CARRIERS else 9, c or "", p["model"], p["plan"] or "",
            ALL_JOINS.index(j) if j in ALL_JOINS else 9)


def commission(per_unit, rate, margin):
    return int((per_unit or 0) + max(0, margin or 0) * (rate or 0) / 100)


def pay_status(expected, paid):
    if paid is None:
        return "미입금"
    if paid == expected:
        return "완료"
    return "차감/부족" if paid < expected else "초과입금"


def fmt_cell(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v.is_integer() or abs(v) >= 100_000:
            return str(int(round(v)))
        return f"{round(v, 4):.4f}".rstrip("0").rstrip(".")
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).replace("\r", " ").replace("\n", " ").strip()


def read_sheets(path):
    """엑셀(xlsx/xls)·CSV → [(시트명, 2차원 문자열 리스트)]  (병합 셀은 값으로 채움)"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        for enc in ("utf-8-sig", "cp949"):
            try:
                with open(path, newline="", encoding=enc) as f:
                    return [("CSV", [[fmt_cell(x) for x in row] for row in csv.reader(f)])]
            except UnicodeDecodeError:
                continue
        raise RuntimeError("CSV 파일 글자를 읽을 수 없습니다.")
    if ext == ".xls":
        import xlrd
        wb = xlrd.open_workbook(path)
        out = []
        for sh in wb.sheets():
            rows = []
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            row.append(fmt_cell(xlrd.xldate_as_datetime(cell.value, wb.datemode)))
                            continue
                        except Exception:
                            pass
                    row.append(fmt_cell(cell.value))
                rows.append(row)
            out.append((sh.name, rows))
        return out
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        merged = {}
        for rng in ws.merged_cells.ranges:
            v = ws.cell(rng.min_row, rng.min_col).value
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    merged[(r, c)] = v
        rows = []
        for r_i, row in enumerate(ws.iter_rows(), 1):
            rows.append([fmt_cell(merged.get((r_i, c_i), cell.value)) for c_i, cell in enumerate(row, 1)])
        out.append((ws.title, rows))
    wb.close()
    return out


def sheets_to_text(path, limit=150_000):
    parts = []
    for name, rows in read_sheets(path):
        lines = []
        prev_line = None
        for i, row in enumerate(rows, 1):
            row = [("〃" if j and len(v) > 12 and v == row[j - 1] else v) for j, v in enumerate(row)]
            while row and row[-1] in ("", "〃"):
                row = row[:-1]
            if not any(row):
                continue
            body = " | ".join(row)
            if body == prev_line:      # 병합으로 똑같은 줄이 반복되면 생략
                continue
            prev_line = body
            lines.append(f"행{i}: " + body)
        if lines:
            parts.append(f"=== 시트: {name} ===\n" + "\n".join(lines))
    text = "\n\n".join(parts)
    if not text.strip():
        raise RuntimeError("엑셀에서 읽을 내용이 없습니다.")
    if len(text) > limit:
        text = text[:limit] + "\n...(너무 길어서 뒷부분 생략)"
    return text


def map_headers(data, aliases, required=0):
    best = (None, {})
    for i, row in enumerate(data[:15]):
        mapping = {}
        for j, cell in enumerate(row):
            h = str(cell).replace(" ", "").lower()
            if not h:
                continue
            for c, keys in enumerate(aliases):
                if c in mapping:
                    continue
                if any(k.lower() in h for k in keys):
                    mapping[c] = j
                    break
        if required in mapping and len(mapping) > len(best[1]):
            best = (i, mapping)
    return best


# ============================================================
# AI 정책 읽기 (Claude API)
# ============================================================
DEFAULT_EXCLUDE = "3G, 2G, WCDMA, 선불"


def exclude_list(text):
    return [k.strip() for k in re.split(r"[,/\n]", text or "") if k.strip()]


def exclude_prompt(ex):
    if not ex:
        return ""
    return (f"\n[읽지 않을 것 — 매우 중요] 다음 말이 들어간 표·구간·요금제·모델·시트는 통째로 건너뛴다 (block·cols·줄을 만들지 말 것): "
            f"{', '.join(ex)}\n  예) '3G' 요금제 열, '3G 단말' 구간의 모델들은 전부 무시. 5G·LTE 휴대폰 정책만 읽는다(제외 목록에 없으면).\n")


def has_word(text, kw):
    """'2G'가 '512G' 안에서 걸리지 않도록 앞뒤가 영문·숫자가 아닐 때만 일치"""
    return re.search(r"(?<![0-9A-Za-z])" + re.escape(kw) + r"(?![0-9A-Za-z])", text or "", re.I) is not None


def drop_excluded(rows, ex):
    """AI가 놓친 것까지 한 번 더 걸러냄 (모델·요금제·그룹에 제외 말이 들어가면 버림)"""
    if not ex:
        return rows
    out = []
    for r in rows:
        blob = " ".join(str(x) for x in (r[2], r[3], r[10] if len(r) > 10 else ""))
        if not any(has_word(blob, k) for k in ex):
            out.append(r)
    return out


def ai_system(carrier, cat_hint):
    return f"""너는 한국 휴대폰 판매점이 대리점에서 받은 '정책표(단가표)'를 표 데이터로 옮기는 도우미다.
첨부된 정책표(엑셀 내용, 사진, PDF)를 읽고 판매점이 받는 정책금액을 빠짐없이 추출하라.

[이 대리점 정보] 통신사: {carrier or '표에서 판단'} / 구분 힌트: {cat_hint}

[출력 형식 — 반드시 지킬 것]
- 머리말·설명·코드블록 없이, 한 줄에 JSON 배열 하나씩만 출력한다. 생각은 짧게 하고 바로 출력을 시작한다.
- 한 줄 = 모델 하나 × 요금제 구간 하나. 가입유형 3가지를 한 줄에 같이 적는다:
  ["구분","통신사","모델","요금제/약정",출고가,공시_A,공시_B,공시_C,리베_A,리베_B,리베_C,"조건/메모","그룹"]
  그룹 = 정책표가 모델을 나눈 이름(예: "프리미엄","중저가","세컨"). 없으면 "".
  무선: A=신규, B=번호이동(MNP), C=기기변경
  유선: A=신규, B=전환(타사), C=재약정
- 그 가입유형 칸이 표에 비어 있으면 null. 한 칸이 여러 유형 공통이면(예: '신규/기변') 해당 자리에 같은 값을 넣는다.
- 구분: "무선"(휴대폰·태블릿·워치) 또는 "유선"(인터넷·TV·집전화 등)
- 통신사: "SKT","KT","LGU+" 중 하나 (표에 없으면 위 대리점 통신사)
- 금액은 원 단위 정수. 표가 만원 단위(예: 45, 45.5, -9)면 450000, 455000, -90000 으로 바꾼다. 모르면 0.
- 공시: 통신사 공시지원금(이통사 지원금). 가입유형별로 다르면 각각, 하나뿐이면 셋 다 같은 값.
- 리베: 판매점이 받는 정책 금액(판매정책, 단가, 리베이트, 수수료 등). 음수도 표에 적힌 그대로.
  조건 없는 추가 정책은 합산하고 조건/메모에 "기본45+추가10" 식으로. 조건부 추가/차감(부가 미유치 차감 등)은 합산하지 말고 조건/메모에 짧게.
- 모델: 펫네임 + 용량 (예: "갤럭시 S26 256GB"). 모델코드가 있으면 뒤에 괄호로 (예: "갤럭시 S26 256GB (SM-S942N256)"). 펫네임이 없거나 #N/A면 같은 계열 다른 행의 펫네임을 참고하고, 그래도 모르면 모델코드만.
- 요금제/약정: 구간명을 짧게 쓰고 뒤에 그 구간에 해당하는 요금제 월정액(천원)을 대괄호로 붙인다.
  예: "115군 [115,130]", "61/69군 [61,69]", "5GX 프라임 [89]". 월정액을 모르면 대괄호 생략. 유선은 약정·결합 조건.
- 워치·태블릿처럼 '단독/번들' 등 구분이 있으면 요금제/약정 칸에 그 구분을 적는다.
- 표에 없는 값을 지어내지 말 것. 차감표·환수표·부당영업 안내 같은 벌칙 표는 줄로 만들지 말고 '#' 요약에만 넣는다.
- 조건/메모에는 아래 규칙으로 옮기지 못한 것만 짧게 적는다.

[부가·차감 규칙 — 정책 줄을 다 쓴 뒤 이어서]
정책표에 있는 부가서비스·보험·요금제·기변 조건 등의 가감을 계산 가능한 규칙으로, 한 줄에 JSON 객체 하나씩:
  {{"규칙":"짧은 이름","조건":"조건키","해당":금액,"미해당":금액,"그룹":"","모델포함":"","모델제외":"","가입유형":"","요금최소":null,"요금최대":null,"메모":""}}
- 조건키: 손님 상황에 따라 달라지면 {', '.join(FLAG_KEYS)} 중 하나.
  손님과 상관없이 대상에 해당하면 항상 적용되는 가감은 "자동".
  금액으로 계산할 수 없는 벌칙·환수·유지기간·개통 후 일어나는 일(해지 환수, 미로그인 등)은 "주의"(금액 0).
  부가=부가서비스, 필링=컬러링·음악 등 콘텐츠 부가, 보험=폰보험, 지원금소액=유통망지원금 10만원 미만 입력,
  단기기변=이전폰 13개월 미만 기변, 초단기기변=183일 미만 기변(단기기변과 중복되면 추가분만), 시니어=시니어·청소년·키즈 요금제.
- 해당/미해당: 그 조건에 해당할 때(가입·유치했을 때)/안 할 때 정책표 금액에 더하거나 빼는 원 단위 금액.
  예) '미유치 4차감' → 해당 0, 미해당 -40000 / '유치시 4추가' → 해당 40000, 미해당 0 /
      '미유치 3차감 유치시 1추가' → 해당 10000, 미해당 -30000 / '자동' 규칙은 해당에 금액, 미해당 0.
- 그룹: 특정 모델 그룹(정책 줄의 그룹 이름)에만 적용되면 그 이름(여러 개면 쉼표). 전체면 "".
- 모델포함/모델제외: 특정 모델만/제외(예: '무너폰 적용제외' → 모델제외 "무너"). 쉼표로 여러 개.
- 가입유형: 특정 유형만이면 "기기변경"처럼(쉼표 여러 개). 전체면 "".
- 요금최소/요금최대: 요금제 월정액(천원) 조건. 예) '115군 기변시 1차감' → 조건 "자동", 해당 -10000, 가입유형 "기기변경", 요금최소 115.
  '33군 미만 10차감' → 조건 "자동", 해당 -100000, 요금최대 32.
- "주의" 규칙은 핵심만 5개 이내로(환수 금액·기간 위주).
- 맨 마지막에 그 밖의 공통 안내가 있으면 '#'으로 시작하는 줄로 3줄 이내 요약한다.
"""


AI_JOIN_ORDER = {"무선": ["신규", "번호이동", "기기변경"], "유선": ["신규", "전환", "재약정"]}


def image_parts(img):
    """QImage → base64 JPEG 리스트 (세로로 긴 캡처는 여러 조각으로 나눔)"""
    if img.isNull():
        raise RuntimeError("이미지를 읽을 수 없습니다.")
    if img.hasAlphaChannel():
        img = img.convertToFormat(QImage.Format.Format_RGB32)
    w, h = img.width(), img.height()
    pieces = []
    if h > w * 2.2:
        step = int(w * 1.5)
        overlap = int(step * 0.06)
        y = 0
        while y < h:
            ph = min(step, h - y)
            pieces.append(img.copy(0, y, w, ph))
            if y + ph >= h:
                break
            y += step - overlap
            if len(pieces) >= 12:
                break
    else:
        pieces = [img]
    out = []
    for p in pieces:
        if max(p.width(), p.height()) > 2000:
            p = p.scaled(2000, 2000, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        ba = QByteArray()
        buf = QBuffer(ba)
        buf.open(QIODevice.OpenModeFlag.WriteOnly)
        p.save(buf, "JPEG", 90)
        buf.close()
        out.append(base64.b64encode(bytes(ba)).decode("ascii"))
    return out


def try_template_rows(path):
    """프로그램 입력양식 파일이면 AI 없이 (정책 줄들, 규칙들) 반환. 아니면 None"""
    if os.path.splitext(path)[1].lower() not in SHEET_EXTS:
        return None
    for _name, rows in read_sheets(path):
        for i, row in enumerate(rows[:5]):
            hdr = [str(c).strip() for c in row]
            if sum(1 for c in POLICY_COLS if c in hdr) >= 7:
                idx = [hdr.index(c) if c in hdr else None for c in POLICY_COLS]
                out, rules, rhdr = [], [], None
                in_rules = False
                for r in rows[i + 1:]:
                    first = str(r[0]).strip() if r else ""
                    if first.startswith("###"):
                        in_rules = True
                        continue
                    if in_rules:
                        if rhdr is None:
                            rhdr = [str(c).strip() for c in r]
                            continue
                        if not any(str(v).strip() for v in r):
                            continue
                        d = {rhdr[j]: r[j] for j in range(min(len(r), len(rhdr))) if rhdr[j]}
                        rules.append(norm_rule(d))
                        continue
                    vals = [(r[j] if j is not None and j < len(r) else "") for j in idx]
                    if any(str(v).strip() for v in vals):
                        out.append(vals)
                return out, rules
    return None


# ------------------------------------------------------------
# 만능 읽기: 엑셀은 AI가 '표 구조(읽는 법)'만 파악하고, 숫자는 프로그램이 원본 셀에서 그대로 읽음
#   → AI가 숫자를 옮겨 적다 틀리는 일이 없고, 표가 커도 빠름
# ------------------------------------------------------------
def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letter):
    n = 0
    for ch in str(letter or "").strip().upper():
        if "A" <= ch <= "Z":
            n = n * 26 + (ord(ch) - 64)
    return n


def first_number(v):
    m = re.search(r"-?\d[\d,]*(?:\.\d+)?", str(v or "").replace(" ", ""))
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", ""))
    except ValueError:
        return None


def grid_numeric_count(grid):
    return sum(1 for row in grid for v in row if re.fullmatch(r"-?[\d,]+(\.\d+)?", str(v).strip() or "x"))


def _is_num(v):
    return re.fullmatch(r"-?[\d,]+(\.\d+)?", v) is not None


def grid_text(grid, limit=140_000, compact=None):
    """AI에게 보여줄 시트 내용: 'R7: B=SM-S942N256 | C=1254000 | …' (빈칸 생략)
    큰 시트는 '압축': 숫자 줄이 계속되면 처음 몇 줄만 전체를 보여주고 나머지는 앞 몇 칸(모델 이름)만 → AI가 훨씬 빨리 읽음.
    (숫자는 어차피 프로그램이 원본 칸에서 직접 읽으므로 AI는 표 구조와 행 범위만 알면 됨)"""
    def rows_iter():
        for r, row in enumerate(grid, 1):
            cells = []
            for c, v in enumerate(row, 1):
                v = str(v).strip()
                if not v:
                    continue
                if c > 1 and len(v) > 12 and v == str(row[c - 2]).strip():
                    continue            # 가로 병합으로 반복되는 긴 글 생략
                cells.append((c, v))
            if cells:
                yield r, cells

    full = []
    prev = None
    for r, cells in rows_iter():
        line = " | ".join(f"{col_letter(c)}={v[:120]}" for c, v in cells)
        if line != prev:
            full.append(f"R{r}: {line}")
        prev = line
    text = "\n".join(full)
    if compact is None:
        compact = len(text) > 30_000
    if not compact:
        return text[:limit] + ("\n...(너무 길어서 뒷부분 생략)" if len(text) > limit else "")
    out, run, prev_sig = [], 0, None
    for r, cells in rows_iter():
        nums = sum(1 for _, v in cells if _is_num(v))
        data_row = nums >= 5
        if not data_row:
            run = 0
            out.append(f"R{r}: " + " | ".join(f"{col_letter(c)}={v[:80]}" for c, v in cells))
            continue
        sig = tuple(v for _, v in cells[2:])          # 모델코드·이름 빼고 숫자가 앞줄과 똑같으면(색상만 다른 모델 등) 생략
        if sig == prev_sig:
            continue
        prev_sig = sig
        run += 1
        if run <= 4:
            out.append(f"R{r}: " + " | ".join(f"{col_letter(c)}={v[:40]}" for c, v in cells))
        else:
            out.append(f"R{r}: " + " | ".join(f"{col_letter(c)}={v[:40]}" for c, v in cells[:4]) + " | …")
    text = "\n".join(out)
    return text[:limit] + ("\n...(너무 길어서 뒷부분 생략)" if len(text) > limit else "")


_CAP_RE = re.compile(r"(128|256|512|1TB|1T)$", re.I)


def model_label(code, name, base_names):
    code, name = str(code or "").strip(), str(name or "").strip()
    bad = lambda x: (not x) or x.upper() in ("#N/A", "N/A", "#REF!", "-", "0")
    m = _CAP_RE.search(code.replace("-", ""))
    cap = ""
    if m:
        cap = "1TB" if m.group(1).upper().startswith("1T") else m.group(1) + "GB"
    base = _CAP_RE.sub("", code.rstrip("-"))
    if not bad(name):
        base_names[base] = name
    else:
        name = base_names.get(base, "")
    if not name or name == code:
        return code
    if cap and not re.search(r"\d+\s*(GB|TB|G\b|T\b)", name, re.I):
        name = f"{name} {cap}"
    return f"{name} ({code})"


def ai_spec_system(carrier, cat_hint):
    return f"""너는 한국 휴대폰 판매점이 대리점에서 받은 '정책표' 엑셀 시트의 구조를 분석해서,
프로그램이 셀 값을 정확히 읽어갈 수 있도록 '읽는 법'을 JSON으로 만드는 도우미다. 숫자를 옮겨 적지 말고 위치만 알려줘라.

[이 대리점] 통신사: {carrier or '표에서 판단'} / 구분 힌트: {cat_hint}
[시트 표기] 각 줄은 'R행번호: 열문자=값 | 열문자=값' (빈칸 생략). 병합된 칸은 같은 값이 여러 칸에 들어 있다.

[출력] 설명·코드블록 없이 JSON 객체 하나만:
{{"blocks":[{{"group":"프리미엄","category":"무선","row_from":7,"row_to":26,
   "model_col":"B","name_col":"D","release_col":"C","unit":10000,
   "subsidy":{{"신규":"F","번호이동":"E","기기변경":"F"}},
   "cols":[{{"col":"G","plan":"115군 [115,130]","join":"신규"}},{{"col":"H","plan":"115군 [115,130]","join":"번호이동"}}]}}],
 "rules":[{{"규칙":"…","조건":"…","해당":0,"미해당":0,"그룹":"","모델포함":"","모델제외":"","가입유형":"","요금최소":null,"요금최대":null,"메모":""}}],
 "notes":["…"]}}

[blocks 규칙]
- block = 모델(또는 상품)이 세로로 나열되고 같은 열 구조를 가진 표 하나. 표가 여러 개면(프리미엄/중저가/워치/태블릿/유선 등) 각각 block.
- row_from/row_to: 그 표에서 모델 데이터가 있는 첫 행과 마지막 행 번호(머리글 행 제외). 중간에 빈 행이 있어도 된다.
- model_col: 모델코드 또는 모델명이 있는 열. name_col: 펫네임(한글/영문 이름) 열, 없으면 "". release_col: 출고가 열, 없으면 "".
- subsidy: 가입유형별 공시지원금(이통사 지원금) 열. 하나뿐이면 세 유형 모두 그 열. 없으면 {{}}.
- cols: 판매점이 받는 정책 금액(리베이트/단가/정책)이 있는 열마다 하나씩. 공시지원금·출고가 열은 넣지 않는다.
  plan = 그 열의 요금제 구간 이름 + 대괄호 안에 그 구간에 들어가는 요금제 월정액(천원) 중 '가장 낮은 금액'.
  예: "115군 [115]", "P_119 베스트109 [109]", "110K 이상 [110]", "61/69군 [61]", "5GX 프라임 [89]". 모르면 대괄호 생략.
  워치·태블릿처럼 요금 구간이 아니라 '단독/번들' 같은 구분이면 plan에 그 이름.
  join = 무선 "신규"/"번호이동"/"기기변경", 유선 "신규"/"전환"/"재약정". 한 열이 여러 유형 공통이면(예: '신규/기변') 같은 col로 유형마다 하나씩 적는다.
  어떤 열의 숫자가 다른 열 금액에 '추가'로 더해지는 구조면 {{"col":"J","plus_col":"N","plan":"번들+추가","join":"번호이동"}} 처럼 plus_col 사용.
- unit: 정책 금액 숫자의 단위. 45, -9 같은 만원 단위면 10000, 400·110 같은 천원 단위면 1000, 450000 같은 원 단위면 1.
  (표 머리글에 '단위: 천원/만원'이 있으면 그대로 따른다. 출고가·공시지원금 단위는 프로그램이 알아서 맞춤)
- 할인 방식: 공시지원금(지원금약정) 정책과 선택약정 정책이 따로 있으면 둘 다 읽는다.
  · 열로 나뉘어 있으면 cols 항목에 "discount":"공시" 또는 "discount":"선약"
  · 행으로 나뉘어 있으면(예: 한 열에 '이통사/선약/공통') 그 열을 block의 "discount_col"로
  · 구분이 없으면 비워둔다.
- 공시지원금이 요금 구간마다 다른 열에 있으면 cols 항목에 "subsidy_col":"H" 처럼 그 열을 적는다.
- 같은 모델이 색상만 다르게 여러 줄(SM-F766N, SM-F766NB, SM-F766NR…)이어도 row 범위에 다 포함하면 된다.
- '3G 전환', '공통지원금표', '요금제 목록' 같은 참고용 표는 block으로 만들지 않는다.
- 셀에 '추가 4', '전유형 21'처럼 글자가 섞여 있어도 프로그램이 숫자만 읽으니 그 열을 그대로 지정하면 된다.
- 벌칙·환수·부당영업 안내 표는 block으로 만들지 않는다.

[rules 규칙 — 부가서비스·보험·요금제·기변 조건의 가감]
- 조건: 손님 상황에 따라 달라지면 {', '.join(FLAG_KEYS)} 중 하나. 대상이면 항상 적용은 "자동".
  금액으로 계산 못 하는 벌칙·환수·유지기간·개통 후 일(해지 환수, 미로그인 등)은 "주의"(금액 0, 5개 이내, 메모에 핵심).
  부가=부가서비스, 필링=컬러링·음악 등 콘텐츠 부가, 보험=폰보험, 지원금소액=유통망지원금 10만원 미만 입력,
  단기기변=이전폰 13개월 미만 기변, 초단기기변=183일 미만 기변(중복이면 추가분만), 시니어=시니어·청소년·키즈 요금제.
- 해당/미해당: 그 조건에 해당(가입·유치)할 때 / 안 할 때 정책표 금액에 더하거나 빼는 원 단위 금액.
  예) '미유치 4차감' → 해당 0, 미해당 -40000 / '유치시 4추가' → 해당 40000, 미해당 0 / '미유치 3차감 유치시 1추가' → 해당 10000, 미해당 -30000.
- 그룹: 특정 block의 group에만 적용되면 그 이름(여러 개면 쉼표), 전체면 "".
- 모델제외: '무너폰 적용제외' → "무너". 가입유형: 특정 유형만이면 쉼표로. 요금최소/요금최대: 요금 조건(천원).
  예) '115군 기변시 1차감' → 조건 "자동", 해당 -10000, 가입유형 "기기변경", 요금최소 115 / '33군 미만 10차감' → 조건 "자동", 해당 -100000, 요금최대 32.
[notes] 그 밖의 공통 안내 3줄 이내.
"""


def parse_json_obj(text):
    a, b = text.find("{"), text.rfind("}")
    if a < 0 or b <= a:
        return None
    try:
        return json.loads(text[a:b + 1])
    except ValueError:
        return None


TIER_MODES = [("top", "제일 높은 요금 구간만 (빠름)"), ("top2", "높은 구간 2개"), ("all", "모든 구간 (느림)")]


def tier_prompt(mode):
    if mode == "top":
        return ("\n[요금 구간 — 중요] cols에는 요금 구간 중 '가장 비싼 구간 하나'만 넣는다(신규·번이·기변, 공시·선약). "
                "나머지 구간 열은 넣지 않는다. 워치·태블릿의 단독/번들 구분은 그대로.\n")
    if mode == "top2":
        return "\n[요금 구간 — 중요] cols에는 요금 구간 중 '가장 비싼 구간 2개'만 넣는다. 나머지 구간 열은 넣지 않는다.\n"
    return ""


def limit_tiers(rows, mode):
    """요금 구간을 위쪽 1~2개만 남김 (요금 숫자가 없는 구분 — 단독/번들 등 — 은 그대로)"""
    if mode not in ("top", "top2"):
        return rows
    keep_n = 1 if mode == "top" else 2
    levels = {}
    for r in rows:
        f = fee_numbers(r[3])
        if f:
            levels.setdefault((r[1], r[10] if len(r) > 10 else ""), set()).add(min(f))
    top = {k: sorted(v, reverse=True)[:keep_n] for k, v in levels.items()}
    out = []
    for r in rows:
        f = fee_numbers(r[3])
        if not f or min(f) in top.get((r[1], r[10] if len(r) > 10 else ""), []):
            out.append(r)
    return out


def apply_spec(grid, spec, carrier, cat_hint, tier_mode="all"):
    """AI가 알려준 '읽는 법'대로 원본 셀에서 숫자를 그대로 읽어 정책 줄 생성"""
    def cell(r, letter):
        c = col_index(letter)
        if not c or r < 1 or r > len(grid) or c > len(grid[r - 1]):
            return ""
        return str(grid[r - 1][c - 1]).strip()

    rows, base_names = [], {}
    for b in spec.get("blocks", []) or []:
        try:
            r0 = int(b.get("row_from"))
            r1 = int(b.get("row_to") or len(grid))
        except (TypeError, ValueError):
            continue
        unit = to_num(b.get("unit") or 1) or 1
        cat = norm_category(str(b.get("category", ""))) or (cat_hint if cat_hint in CATEGORIES else "무선")
        grp = str(b.get("group", "") or "").strip()
        sub = {norm_join(k, cat) or k: v for k, v in (b.get("subsidy") or {}).items()}
        disc_col = b.get("discount_col")
        prev_code, prev_sig = None, None
        for r in range(max(1, r0), min(r1, len(grid)) + 1):
            code = cell(r, b.get("model_col"))
            sig = (cell(r, b.get("release_col")) if b.get("release_col") else "",
                   cell(r, b.get("discount_col")) if b.get("discount_col") else "") + \
                tuple(cell(r, c.get("col")) for c in (b.get("cols") or []))
            if prev_code and code and sig == prev_sig and \
                    len(os.path.commonprefix([code, prev_code])) >= max(6, len(prev_code) - 3):
                continue            # 색상만 다른 같은 모델(SM-F766N → SM-F766NB, NR, NSO…)은 한 번만
            if code:
                prev_code, prev_sig = code, sig
            if not code or code in ("모델", "모델명", "모델코드", "기종", "펫네임") or first_number(code) == to_num(code) != 0 \
                    or len(code) > 60:
                continue
            name = cell(r, b.get("name_col")) if b.get("name_col") else ""
            model = model_label(code, name, base_names)
            rp = first_number(cell(r, b.get("release_col"))) if b.get("release_col") else None
            if rp and 0 < rp < 100_000:
                rp *= 1000                                   # 천원 단위 출고가
            row_disc = ""
            if disc_col:
                t = cell(r, disc_col)
                row_disc = "선약" if ("선약" in t or "선택" in t) else ("" if ("공통" in t or not t) else "공시")
            for c in b.get("cols", []) or []:
                x = first_number(cell(r, c.get("col")))
                if x is None:
                    continue
                if c.get("plus_col"):
                    p = first_number(cell(r, c.get("plus_col")))
                    if p:
                        x += p
                join = norm_join(str(c.get("join", "")), cat) or str(c.get("join", "")).strip()
                sc = c.get("subsidy_col") or sub.get(join)
                sv = first_number(cell(r, sc)) if sc else None
                if sv and 0 < sv < 20_000:
                    sv *= 1000                               # 천원 단위 공시지원금
                disc = str(c.get("discount", "") or "").strip() or row_disc
                if disc and "선" in disc and row_disc == "공시":
                    continue                                 # 행과 열의 할인 방식이 안 맞으면 건너뜀
                if disc and ("공시" in disc or "지원금" in disc) and row_disc == "선약":
                    continue
                plan = str(c.get("plan", "")).strip()
                if "선" in disc and "(선약)" not in plan:
                    plan += " (선약)"
                rows.append([cat, carrier, model, plan, join, int(rp or 0),
                             int(sv or 0), int(round(x * unit)), 0, "", grp])
    return limit_tiers(rows, tier_mode)


def is_select_plan(plan):
    return "(선약)" in (plan or "")


def by_discount(cands, mode):
    """할인 방식에 맞는 정책만: 공시지원 → '(선약)' 아닌 것 / 선택약정 → '(선약)' 우선, 없으면 공통"""
    if mode == "선약":
        sel = [p for p in cands if is_select_plan(p["plan"])]
        return sel or [p for p in cands if not is_select_plan(p["plan"])]
    return [p for p in cands if not is_select_plan(p["plan"])]


def discount_combo():
    cb = QComboBox()
    cb.addItem("공시지원금", "공시")
    cb.addItem("선택약정", "선약")
    cb.setToolTip("손님이 공시지원금을 받는지, 선택약정(요금 25% 할인)인지 — 대리점 정책이 다를 수 있음")
    return cb


def xlsx_images(path):
    """엑셀 안에 붙여넣은 정책 '사진' 꺼내기"""
    out = []
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.startswith("xl/media/") and n.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp")):
                    img = QImage()
                    if img.loadFromData(z.read(n)) and img.width() > 300 and img.height() > 200:
                        out.append((os.path.basename(n), img))
    except Exception:
        pass
    return out


_SHEET_BAD = ("차감", "환수", "공지", "안내", "재고", "매입", "서식", "양식", "신청서", "가이드", "유의", "패널티",
              "페널티", "정산", "부당", "개통표", "실적", "목록", "리스트", "사은품표", "주소", "연락처", "변경내역", "이력")
_SHEET_GOOD = ("모델", "펫네임", "신규", "mnp", "번이", "번호이동", "기변", "기기변경", "요금", "정책", "단가", "리베이트", "출고가")


_CURRENT_EXCLUDE = [DEFAULT_EXCLUDE]      # 화면에서 바꾸면 여기에 반영

SHEET_MODES = [("policy", "📊 정책표로 읽기"), ("rules", "📝 부가·차감 조건만 읽기"), ("skip", "🚫 안 읽기")]
_SKIP_NAMES = ("지원금", "공시", "요금제", "테이블", "계산기", "카드", "전환", "재고", "주소", "연락처", "양식", "서식",
               "신청서", "실적", "이력", "변경내역", "매입")
_RULE_NAMES = ("공지", "부가", "추가정책", "적용기준", "페널티", "패널티", "차감", "안내", "유의", "환수", "기준")


def grid_char_count(grid):
    return sum(len(str(v)) for row in grid for v in row if str(v).strip())


def classify_sheet(name, grid):
    """시트 이름·내용으로 기본 읽기 방법 추천: policy(정책표) / rules(조건만) / skip(안 읽음)"""
    n = name.replace(" ", "")
    if any(has_word(name, k) for k in exclude_list(_CURRENT_EXCLUDE[0])):
        return "skip"
    nums = grid_numeric_count(grid)
    if any(k in n for k in _SKIP_NAMES) and not any(k in n for k in ("단가", "가격표", "정책표")):
        return "skip"
    if any(k in n for k in _RULE_NAMES):
        return "rules" if grid_char_count(grid) > 150 else "skip"
    if nums >= 30:
        return "policy"
    return "rules" if grid_char_count(grid) > 300 else "skip"


def rules_system(carrier):
    return f"""너는 한국 휴대폰 판매점이 대리점에서 받은 정책 공지·추가정책 문서에서 '부가·차감 조건'만 뽑는 도우미다.
[이 대리점] 통신사: {carrier or '문서에서 판단'}
문서에 있는 부가서비스·보험·컬러링·요금제 유지·기변·카드·결합 등의 가감 조건을, 한 줄에 JSON 객체 하나씩만 출력한다(설명 금지):
  {{"규칙":"짧은 이름","조건":"조건키","해당":금액,"미해당":금액,"그룹":"","모델포함":"","모델제외":"","가입유형":"","요금최소":null,"요금최대":null,"메모":""}}
- 조건키: 손님 상황에 따라 달라지면 {', '.join(FLAG_KEYS)} 중 하나. 대상이면 항상 적용은 "자동".
  금액으로 계산 못 하는 벌칙·환수·유지기간·개통 후 일은 "주의"(금액 0, 최대 6개, 메모에 핵심).
- 해당/미해당: 그 조건에 해당(가입·유치)할 때/안 할 때 정책 금액에 더하거나 빼는 원 단위 금액.
  예) '미유치 4차감' → 해당 0, 미해당 -40000 / '유치시 4추가' → 해당 40000, 미해당 0 / '미가입시 -2' → 해당 0, 미해당 -20000.
  금액 숫자가 만원 단위(4, -2 등)면 원으로 바꿔 적는다.
- 모델제외·가입유형·요금최소/요금최대(월 요금 천원) 조건이 있으면 채운다.
- 개통 방법·서류·연락처·일반 안내처럼 금액 가감이 아닌 내용은 무시한다."""


def sheet_score(name, grid):
    """정책표일 가능성 점수 (높을수록 정책표). 음수면 기본으로 빼둠"""
    nums = grid_numeric_count(grid)
    head = " ".join(str(v) for row in grid[:40] for v in row if str(v).strip()).lower()
    good = sum(1 for k in _SHEET_GOOD if k in head)
    bad = any(k in name for k in _SHEET_BAD) or any(has_word(name, k) for k in exclude_list(_CURRENT_EXCLUDE[0]))
    score = min(nums, 400) / 40 + good * 2 - (8 if bad else 0)
    if nums < 20:
        score -= 10
    return score


def load_source(path):
    name = os.path.basename(path)
    ext = os.path.splitext(path)[1].lower()
    if ext in SHEET_EXTS:
        sheets = [(sn, g) for sn, g in read_sheets(path) if any(any(str(v).strip() for v in row) for row in g)]
        if not sheets:
            raise RuntimeError("엑셀에서 읽을 내용이 없습니다.")
        src = dict(name=name, kind="sheet", data=sheets, images=[])
        if ext in (".xlsx", ".xlsm") and sum(grid_numeric_count(g) for _, g in sheets) < 30:
            src["images"] = [(n, image_parts(img)) for n, img in xlsx_images(path)]
        return src
    if ext in IMG_EXTS:
        return dict(name=name, kind="image", data=image_parts(QImage(path)))
    if ext == ".pdf":
        with open(path, "rb") as f:
            raw = f.read()
        if len(raw) > 25 * 1024 * 1024:
            raise RuntimeError("PDF가 너무 큽니다 (25MB 이하).")
        return dict(name=name, kind="pdf", data=base64.b64encode(raw).decode("ascii"))
    if ext in (".heic", ".heif"):
        raise RuntimeError("아이폰 사진(HEIC)은 바로 못 읽습니다. 사진을 화면에 띄워 캡처한 뒤 '붙여넣기' 해주세요.")
    raise RuntimeError("지원하지 않는 파일입니다. (엑셀, CSV, 사진, PDF만 가능)")


def build_content(src):
    ask_txt = "위 정책표에서 규칙대로 모든 모델·요금제 구간을 빠짐없이 추출해줘."
    if src["kind"] == "text":
        return [{"type": "text", "text": f"[파일: {src['name']}]\n{src['data']}\n\n{ask_txt}"}]
    if src["kind"] == "sheet":
        body = "\n\n".join(f"=== 시트: {sn} ===\n{grid_text(g)}" for sn, g in src["data"])
        return [{"type": "text", "text": f"[파일: {src['name']}]\n{body}\n\n{ask_txt}"}]
    if src["kind"] == "pdf":
        return [{"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": src["data"]}},
                {"type": "text", "text": f"[파일: {src['name']}] {ask_txt}"}]
    blocks = [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": p}}
              for p in src["data"]]
    note = ""
    if len(src["data"]) > 1:
        note = " 이미지 여러 장은 세로로 긴 한 장의 표를 위에서부터 나눈 조각이다. 겹치는 부분은 한 번만 추출하라."
    blocks.append({"type": "text", "text": f"[파일: {src['name']}]{note} {ask_txt}"})
    return blocks


class Cancelled(Exception):
    pass


def _claude_request(key, model, system, content, max_tokens=32000, stream=True, on_event=None,
                should_stop=None, timeout=180):
    """on_event(kind, info): kind = 'connected' | 'thinking' | 'text'(줄수) | 'ping'"""
    body = {"model": model, "max_tokens": max_tokens, "system": system,
            "messages": [{"role": "user", "content": content}]}
    if stream:
        body["stream"] = True
    req = urllib.request.Request(
        API_URL, data=json.dumps(body).encode("utf-8"), method="POST",
        headers={"x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        if not stream:
            d = json.loads(resp.read().decode("utf-8"))
            text = "".join(b.get("text", "") for b in d.get("content", []) if b.get("type") == "text")
            return text, d.get("stop_reason")
        parts, stop, lines = [], None, 0
        while True:
            if should_stop and should_stop():
                resp.close()
                raise Cancelled()
            raw = resp.readline()
            if not raw:
                break
            line = raw.decode("utf-8", "ignore").strip()
            if not line.startswith("data:"):
                continue
            try:
                ev = json.loads(line[5:].strip())
            except ValueError:
                continue
            t = ev.get("type")
            if t == "message_start":
                on_event and on_event("connected", 0)
            elif t == "content_block_delta":
                d = ev.get("delta", {})
                if d.get("type") == "text_delta":
                    chunk = d.get("text", "")
                    parts.append(chunk)
                    lines += chunk.count("\n")
                    on_event and on_event("text", lines)
                else:
                    on_event and on_event("thinking", 0)
            elif t == "ping":
                on_event and on_event("ping", 0)
            elif t == "message_delta":
                stop = ev.get("delta", {}).get("stop_reason") or stop
            elif t == "error":
                raise RuntimeError(ev.get("error", {}).get("message", "AI 서버 오류"))
        return "".join(parts), stop


def _to_openai(content):
    parts = []
    for b in content:
        if b["type"] == "text":
            parts.append({"type": "text", "text": b["text"]})
        elif b["type"] == "image":
            parts.append({"type": "image_url", "image_url": {
                "url": f"data:{b['source']['media_type']};base64,{b['source']['data']}", "detail": "high"}})
        elif b["type"] == "document":
            parts.append({"type": "file", "file": {"filename": "policy.pdf",
                                                  "file_data": f"data:application/pdf;base64,{b['source']['data']}"}})
    return parts


def _to_gemini(content):
    parts = []
    for b in content:
        if b["type"] == "text":
            parts.append({"text": b["text"]})
        elif b["type"] in ("image", "document"):
            parts.append({"inline_data": {"mime_type": b["source"]["media_type"], "data": b["source"]["data"]}})
    return parts


def _post_json(url, body, headers, timeout):
    h = {"content-type": "application/json", "User-Agent": "PhonePolicyManager"}
    h.update(headers)
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), method="POST", headers=h)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def api_request(key, model, system, content, max_tokens=32000, stream=True, on_event=None,
                should_stop=None, timeout=180):
    """키 종류를 보고 Claude / OpenAI / Gemini 중 맞는 곳으로 요청"""
    prov = ai_provider(key)
    if not prov:
        raise RuntimeError("AI 키 모양을 알 수 없습니다.\nClaude(sk-ant-…), OpenAI GPT(sk-…), Gemini(AIza…) 키 중 하나를 넣어주세요.")
    model = pick_model(key, model)
    if prov == "claude":
        return _claude_request(key, model, system, content, max_tokens, stream, on_event, should_stop, timeout)
    if on_event:
        on_event("waiting", 0)
    long_to = max(timeout, 400) if max_tokens > 100 else timeout
    if prov == "openai":
        body = {"model": model, "max_completion_tokens": max(max_tokens, 2000),
                "messages": [{"role": "system", "content": system}, {"role": "user", "content": _to_openai(content)}]}
        if model.startswith(("gpt-5", "o1", "o3", "o4")):
            body["reasoning_effort"] = "low"          # 생각은 짧게 → 훨씬 빠름
        d = _post_json("https://api.openai.com/v1/chat/completions", body, {"Authorization": f"Bearer {key}"}, long_to)
        ch = (d.get("choices") or [{}])[0]
        text = (ch.get("message") or {}).get("content") or ""
        return text, ("max_tokens" if ch.get("finish_reason") == "length" else ch.get("finish_reason"))
    body = {"system_instruction": {"parts": [{"text": system}]},
            "contents": [{"role": "user", "parts": _to_gemini(content)}],
            "generationConfig": {"maxOutputTokens": max(max_tokens, 2000)}}
    if "2.5" in model:
        body["generationConfig"]["thinkingConfig"] = {"thinkingBudget": 1024}   # 생각은 짧게 → 빠름
    d = _post_json(f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                   body, {"x-goog-api-key": key}, long_to)
    cand = (d.get("candidates") or [{}])[0]
    text = "".join(p.get("text", "") for p in (cand.get("content") or {}).get("parts", []) if not p.get("thought"))
    return text, ("max_tokens" if cand.get("finishReason") == "MAX_TOKENS" else cand.get("finishReason"))


def friendly_error(ex):
    if isinstance(ex, urllib.error.HTTPError):
        try:
            body = ex.read().decode("utf-8", "ignore")
            msg = json.loads(body).get("error", {}).get("message", body[:300])
        except Exception:
            msg = str(ex)
        low = msg.lower()
        if ex.code == 401 or "api key not valid" in low or "invalid api key" in low or "incorrect api key" in low:
            return "AI 키가 올바르지 않습니다.\n설정 탭에서 키를 다시 붙여넣고 '연결 테스트'를 눌러보세요."
        if "credit" in low or "billing" in low or "quota" in low:
            return ("AI 충전금(크레딧)이 부족하거나 사용 한도를 넘었습니다.\n"
                    "Claude: console.anthropic.com / GPT: platform.openai.com / Gemini: aistudio.google.com 에서 확인해 주세요.")
        if ex.code == 404 or ("model" in low and ex.code in (400, 403)):
            return f"AI 모델 이름이 맞지 않습니다. 설정 탭의 '모델'을 확인하세요.\n({msg})"
        if ex.code in (429, 500, 503, 529):
            return "AI 서버가 바쁩니다. 1~2분 뒤 다시 눌러주세요."
        if ex.code == 413 or "too large" in low or "too long" in low:
            return "파일이 너무 큽니다. 시트를 나누거나 사진을 몇 장으로 나눠서 올려주세요."
        return f"AI 요청 오류 ({ex.code})\n{msg}"
    if isinstance(ex, (socket.timeout, TimeoutError)):
        return ("AI 서버에서 3분 동안 아무 응답이 없습니다.\n"
                "인터넷(회사 방화벽·백신 프로그램)이 막고 있는지 확인하고, 설정 탭 [연결 테스트]를 먼저 눌러보세요.")
    if isinstance(ex, urllib.error.URLError):
        reason = str(ex.reason)
        if "CERTIFICATE" in reason.upper():
            return ("보안 인증서 문제로 AI 서버에 연결을 못 했습니다.\n"
                    "cmd 창에서  pip install pip-system-certs  입력 후 프로그램을 다시 켜주세요.")
        return f"인터넷 연결을 확인해 주세요.\n({reason})"
    return str(ex)


def parse_ai_lines(text):
    rows, notes = [], []
    for line in text.splitlines():
        s = line.strip().rstrip(",")
        if not s or s.startswith("```"):
            continue
        if s.startswith("#"):
            notes.append(s.lstrip("#").strip())
            continue
        if s.startswith("["):
            try:
                arr = json.loads(s)
            except ValueError:
                continue
            if isinstance(arr, list) and len(arr) >= 5:
                rows.append(arr)
        elif s.startswith("{"):
            try:
                obj = json.loads(s)
            except ValueError:
                continue
            if isinstance(obj, dict):
                rows.append(obj)
    return rows, notes


def _blank(v):
    return v is None or str(v).strip() in ("", "-", "null", "None")


def expand_ai_row(arr, carrier, cat_hint):
    """AI 한 줄(모델×요금제, 가입유형 3개) → 프로그램 표 줄 여러 개"""
    a = list(arr)
    cat = norm_category(str(a[0])) or (cat_hint if cat_hint in CATEGORIES else "무선")
    car = norm_carrier(str(a[1])) or carrier or str(a[1]).strip()
    model, plan = str(a[2]).strip(), str(a[3]).strip()
    if len(a) >= 12:                      # 새 형식
        rp, subs, rebs, cond = a[4], a[5:8], a[8:11], str(a[11] or "").strip()
        grp = str(a[12] or "").strip() if len(a) > 12 else ""
        out = []
        for i, jt in enumerate(AI_JOIN_ORDER[cat]):
            if _blank(rebs[i]):
                continue
            out.append([cat, car, model, plan, jt, money_auto(rp), money_auto(subs[i]),
                        money_auto(rebs[i]), 0, cond, grp])
        return out
    a = (a + [""] * 10)[:10]              # 예전 형식(가입유형 1개)
    join = norm_join(str(a[4]), cat) or str(a[4]).strip()
    return [[cat, car, model, plan, join] + [money_auto(x) for x in a[5:9]] + [str(a[9] or "").strip(), ""]]


def save_ai_log(name, text):
    try:
        d = os.path.join(BASE_DIR, "ai_log")
        os.makedirs(d, exist_ok=True)
        safe = re.sub(r'[\\/:*?"<>|]', "_", name)[:40]
        with open(os.path.join(d, f"{datetime.datetime.now():%Y%m%d_%H%M%S}_{safe}.txt"), "w",
                  encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass


class AIWorker(QThread):
    progress = Signal(str)
    done = Signal(object, object, object, bool)
    failed = Signal(str)

    def __init__(self, key, model, sources, carrier, cat_hint):
        super().__init__()
        self.key, self.model, self.sources = key, model, sources
        self.carrier, self.cat_hint = carrier, cat_hint
        self.exclude = []
        self.tier_mode = "top"
        self.diag = []
        self._stop = False

    def cancel(self):
        self._stop = True

    def _call(self, system, content, tag, max_tokens=32000):
        def ev(kind, k):
            if kind == "connected":
                self.progress.emit(f"🤖 {tag} — 연결됨. AI가 표를 살펴보는 중…")
            elif kind == "thinking":
                self.progress.emit(f"🤖 {tag} — AI가 표 구조를 파악하는 중…")
            elif kind == "text":
                self.progress.emit(f"🤖 {tag} — 받아 적는 중… {k}줄")
            elif kind == "waiting":
                self.progress.emit(f"🤖 {tag} — 연결됨. AI가 읽는 중… (다 읽으면 한 번에 도착합니다)")
        import time
        for attempt in range(4):         # 동시에 여러 개 보내다 '잠깐 쉬어라(429)'가 오면 기다렸다 다시
            try:
                text, stop = api_request(self.key, self.model, system, content, max_tokens=max_tokens,
                                         on_event=ev, should_stop=lambda: self._stop)
                break
            except urllib.error.HTTPError as ex:
                if ex.code not in (429, 500, 503, 529) or attempt == 3 or self._stop:
                    raise
                self.progress.emit(f"⏳ {tag} — AI 서버가 바빠서 {15 * (attempt + 1)}초 뒤 다시 시도…")
                time.sleep(15 * (attempt + 1))
        save_ai_log(tag, text)
        return text, stop

    def _transcribe(self, src, tag, car=None):
        """사진·PDF·카톡 글·구조 분석 실패한 표: AI가 직접 옮겨 적기"""
        car = car or self.carrier
        text, stop = self._call(ai_system(car, self.cat_hint) + exclude_prompt(self.exclude) + tier_prompt(self.tier_mode),
                                build_content(src), tag)
        r, nt = parse_ai_lines(text)
        rows, rules = [], []
        for x in r:
            if isinstance(x, dict):
                rules.append(norm_rule(x))
            else:
                rows += expand_ai_row(x, car, self.cat_hint)
        return limit_tiers(rows, self.tier_mode), rules, nt, stop == "max_tokens"

    def _canon(self, row_lists):
        """여러 대리점 정책표의 모델 이름을 같은 기종끼리 같은 이름으로 (SK·KT·LG 비교가 되도록)"""
        names = sorted({r[2] for rows in row_lists for r in rows if r and r[2]})
        if not names:
            return
        self.progress.emit(f"🔤 모델 이름 {len(names)}개를 통신사끼리 맞추는 중…")
        mapping = {}
        existing = list(getattr(self, "existing_models", []))[:400]
        for i in range(0, len(names), 250):
            chunk = names[i:i + 250]
            content = [{"type": "text", "text": "[기존 표준 이름 목록]\n" + "\n".join(existing) +
                        "\n\n[정리할 모델 이름]\n" + "\n".join(chunk)}]
            try:
                text, _ = self._call(CANON_SYSTEM, content, "모델이름", 16000)
                d = parse_json_obj(text) or {}
                for k, v in d.items():
                    if isinstance(v, str) and v.strip():
                        mapping[k] = [v.strip()]
                    elif isinstance(v, list):
                        vs = [str(x).strip() for x in v if str(x).strip()]
                        if vs:
                            mapping[k] = vs
            except Cancelled:
                raise
            except Exception:
                pass
        for rows in row_lists:
            new = []
            for r in rows:
                if r and r[2] in mapping:
                    for nm in mapping[r[2]]:
                        x = list(r)
                        x[2] = nm
                        new.append(x)
                else:
                    new.append(r)
            rows[:] = new

    def _sheet_task(self, src, sn, grid, tag, car=None):
        car = car or self.carrier
        stag = f"{tag} [{sn}]"
        content = [{"type": "text", "text": f"[시트: {sn}]\n{grid_text(grid)}\n\n위 시트의 읽는 법(JSON)을 만들어줘."}]
        text, _ = self._call(ai_spec_system(car, self.cat_hint) + exclude_prompt(self.exclude) + tier_prompt(self.tier_mode),
                             content, stag, 16000)
        spec = parse_json_obj(text) or {}
        got = apply_spec(grid, spec, car, self.cat_hint, self.tier_mode)
        if not spec:
            self.diag.append(f"{stag}: AI 응답 해석 실패 — {text.strip()[:150] or '(빈 응답)'}")
        elif len(got) < 3:
            self.diag.append(f"{stag}: 표 구조 응답은 왔는데 금액을 못 찾음 (구역 {len(spec.get('blocks') or [])}개)")
        if len(got) >= 3:
            return (got, [norm_rule(x) for x in (spec.get("rules") or []) if isinstance(x, dict)],
                    [f"[{sn}] {x}" for x in (spec.get("notes") or [])], False, len(got))
        # 구조 분석이 안 되는 특이한 표 → AI가 직접 옮겨 적기
        self.progress.emit(f"🤖 {stag} — 특이한 표라서 AI가 직접 옮겨 적는 중…")
        r, ru, nt, tr = self._transcribe(dict(name=f"{src['name']} [{sn}]", kind="sheet", data=[(sn, grid)]), stag, car)
        return r, ru, nt, tr, 0

    def _rules_task(self, src, sn, grid, tag, car=None):
        car = car or self.carrier
        lines = [" | ".join(str(v).strip() for v in row if str(v).strip()) for row in grid]
        text = "\n".join(l for l in lines if l)[:15000]
        t, _ = self._call(rules_system(car), [{"type": "text", "text": f"[{sn}]\n{text}"}], f"{tag} [{sn}] 조건", 16000)
        r, nt = parse_ai_lines(t)
        return [], [norm_rule(x) for x in r if isinstance(x, dict)], [], False, 0

    def _other_task(self, src, tag, car=None):
        r, ru, nt, tr = self._transcribe(src, tag, car)
        return r, ru, [f"[{src['name']}] {x}" for x in nt], tr, 0

    def run(self):
        try:
            tasks = []
            n = len(self.sources)
            for i, src in enumerate(self.sources, 1):
                tag = f"({i}/{n}) {src['name']}"
                if src["kind"] == "sheet":
                    for sn, grid in src["data"]:
                        if grid_numeric_count(grid) >= 5:
                            tasks.append((self._sheet_task, (src, sn, grid, tag)))
                    for sn, grid in src.get("rule_sheets", []):
                        tasks.append((self._rules_task, (src, sn, grid, tag)))
                    for iname, parts in src.get("images", []):
                        tasks.append((self._other_task, (dict(name=f"{src['name']} 속 사진 {iname}", kind="image",
                                                               data=parts), f"{tag} 사진")))
                else:
                    tasks.append((self._other_task, (src, tag)))
            if not tasks:
                self.done.emit([], [], [], False)
                return
            self.progress.emit(f"🤖 AI 서버에 연결하는 중… (시트·사진 {len(tasks)}개를 동시에 읽습니다)")
            results = [None] * len(tasks)
            done_n = [0]

            def run_one(i):
                fn, args = tasks[i]
                results[i] = fn(*args)
                done_n[0] += 1
                self.progress.emit(f"📐 {done_n[0]}/{len(tasks)}개 다 읽음…")

            with ThreadPoolExecutor(max_workers=min(4, len(tasks))) as ex:
                futs = [ex.submit(run_one, i) for i in range(len(tasks))]
                for f in futs:
                    f.result()          # 오류가 있으면 여기서 올라옴
            results = [(drop_excluded(r[0], self.exclude),) + tuple(r[1:]) for r in results]
            self._canon([r[0] for r in results])
            rows, rules, notes, truncated = [], [], [], False
            self.exact = 0
            for r, ru, nt, tr, ex_n in results:
                rows += r
                rules += ru
                notes += nt
                truncated |= tr
                self.exact += ex_n
            self.done.emit(rows, rules, notes, truncated)
        except Cancelled:
            pass
        except Exception as ex:
            save_ai_log("오류", repr(ex))
            self.failed.emit(friendly_error(ex))


CANON_SYSTEM = """너는 한국 휴대폰 모델명 정리 도우미다. 아래 모델 이름들은 SK·KT·LG 여러 대리점 정책표에서 모은 것이라
같은 기종이 서로 다르게 적혀 있다. 같은 기종·같은 용량이면 반드시 같은 '표준 이름'이 되게 바꿔라.
- 표준 이름 형식: 브랜드 + 모델 + 용량. 예) "갤럭시 S26 256GB", "갤럭시 S26 울트라 512GB", "갤럭시 Z 플립7 256GB",
  "아이폰 17 Pro Max 256GB", "아이폰 17 256GB", "갤럭시 A17", "갤럭시 워치8 44mm", "스타일폴더2"
- 모델코드만 있으면(SM-S942N256, UIP17PM-256, AT-M140L 등) 코드로 기종을 알아내서 이름으로 바꾼다.
- 통신사 전용 표시(모델코드 끝 K/L/S 등)나 색상은 빼고, 용량은 살린다.
  용량이 안 적힌 기본 모델은 그 기종의 기본(가장 작은) 용량을 붙인다. 예) "갤럭시 Z 폴드7" → "갤럭시 Z 폴드7 256GB",
  "아이폰 17 Pro" → "아이폰 17 Pro 256GB". 용량이 없는 기종(워치·키즈폰·폴더폰 등)은 붙이지 않는다.
- [기존 표준 이름 목록]에 같은 기종·용량이 있으면 반드시 그 이름을 그대로 쓴다.
- 확실하지 않으면 원래 이름을 최대한 살린다. 서로 다른 기종을 같은 이름으로 합치지 말 것.
- 한 이름에 여러 기종·용량이 함께 적혀 있으면(예: "S26 일반 256/512 S26 울트라 256/512", "SM-S942(8)NK 갤럭시 S26(Ultra)",
  "AIP17(P,PM) iPhone 17 (Pro,Max)") 해당하는 표준 이름을 모두 배열로 준다.
  예) "S26 일반 256/512" → ["갤럭시 S26 256GB","갤럭시 S26 512GB"]
출력: 설명 없이 JSON 객체 하나 {"원래 이름": "표준 이름" 또는 ["표준 이름", …], ...} — [정리할 모델 이름]을 모두 포함."""


class BatchWorker(AIWorker):
    """여러 대리점 파일을 한꺼번에: 전부 동시에 읽고, 모델 이름을 통신사끼리 맞춘 뒤 파일별 결과로 돌려줌"""
    batch_done = Signal(object)      # dict(int→결과)는 Qt 신호로 넘기면 키가 망가져서 object로 그대로 전달

    def __init__(self, key, model, items, cached, existing_models):
        super().__init__(key, model, [], "", "자동 판단")
        self.items = items              # [(idx, src, carrier)]
        self.cached = cached            # {idx: (rows, rules, notes, trunc, exact)}
        self.existing_models = existing_models

    def run(self):
        try:
            tasks = []
            for idx, src, car in self.items:
                tag = src["name"]
                if src["kind"] == "sheet":
                    for sn, grid in src["data"]:
                        if grid_numeric_count(grid) >= 5:
                            tasks.append((idx, self._sheet_task, (src, sn, grid, tag, car)))
                    for sn, grid in src.get("rule_sheets", []):
                        tasks.append((idx, self._rules_task, (src, sn, grid, tag, car)))
                    for iname, parts in src.get("images", []):
                        tasks.append((idx, self._other_task, (dict(name=f"{tag} 속 사진 {iname}", kind="image",
                                                                    data=parts), f"{tag} 사진", car)))
                else:
                    tasks.append((idx, self._other_task, (src, tag, car)))
            results = [None] * len(tasks)
            cnt = [0]
            if tasks:
                self.progress.emit(f"🤖 AI가 {len(tasks)}개 시트·사진을 동시에 읽는 중…")

                def run_one(i):
                    _, fn, args = tasks[i]
                    results[i] = fn(*args)
                    cnt[0] += 1
                    self.progress.emit(f"📐 {cnt[0]}/{len(tasks)}개 다 읽음…")

                with ThreadPoolExecutor(max_workers=min(4, len(tasks))) as ex:
                    for f in [ex.submit(run_one, i) for i in range(len(tasks))]:
                        f.result()
            per = {k: [list(v[0]), list(v[1]), list(v[2]), v[3], v[4]] for k, v in self.cached.items()}
            for (idx, _, _), res in zip(tasks, results):
                p = per.setdefault(idx, [[], [], [], False, 0])
                p[0] += res[0]
                p[1] += res[1]
                p[2] += res[2]
                p[3] = p[3] or res[3]
                p[4] += res[4]
            for idx, src_, _ in self.items:
                if not per.get(idx, [[]])[0]:
                    mine = [d for d in self.diag if d.startswith(src_["name"])]
                    per.setdefault(idx, [[], [], [], False, 0])[2].extend(mine or ["AI가 이 파일에서 정책 금액을 못 찾음"])
            for p in per.values():
                p[0] = drop_excluded(p[0], self.exclude)
            self._canon([p[0] for p in per.values()])
            self.batch_done.emit(per)
        except Cancelled:
            pass
        except Exception as ex:
            save_ai_log("오류", repr(ex))
            self.failed.emit(friendly_error(ex))


def save_policy_set(db, aid, vf, rows, rules=None, end_missing=True):
    """정책 줄들 + 규칙을 한 대리점에 저장 → (신규·변경, 동일, 종료, 문제줄수)"""
    items, bad = {}, 0
    for v in rows:
        v = [("" if x is None else str(x)) for x in (list(v) + [""] * 11)[:11]]
        cat = norm_category(v[0]) or "무선"
        car = norm_carrier(v[1]) or v[1].strip()
        model, plan = v[2].strip(), v[3].strip()
        join = norm_join(v[4], cat) or (v[4].strip() if cat == "유선" else "")
        if not (car and model and join):
            bad += 1
            continue
        items[(cat, car, model, plan, join)] = dict(release_price=to_int(v[5]), public_subsidy=to_int(v[6]),
                                                    rebate=to_int(v[7]), deduction=to_int(v[8]),
                                                    conditions=v[9].strip(), grp=v[10].strip())
    ins = same = ended = 0
    for k, d in items.items():
        cur = db.find_policy(vf, aid, *k)
        if cur and all(cur[f] == d[f] for f in MONEY_FIELDS) and (cur["conditions"] or "") == d["conditions"] \
                and (cur["grp"] or "") == d["grp"]:
            same += 1
            continue
        db.x("INSERT INTO policies(agency_id,category,carrier,model,plan,join_type,release_price,public_subsidy,"
             "rebate,deduction,conditions,grp,valid_from,deleted,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,0,?)",
             (aid, *k, d["release_price"], d["public_subsidy"], d["rebate"], d["deduction"], d["conditions"],
              d["grp"], vf, now()), commit=False)
        ins += 1
    if end_missing and items:
        cats = {k[0] for k in items}
        for p in db.effective_policies(vf, aid):
            k = (p["category"], p["carrier"], p["model"], p["plan"], p["join_type"])
            if k not in items and k[0] in cats:
                db.x("INSERT INTO policies(agency_id,category,carrier,model,plan,join_type,valid_from,deleted,"
                     "created_at) VALUES(?,?,?,?,?,?,?,1,?)", (aid, *k, vf, now()), commit=False)
                ended += 1
    db.commit()
    if rules:
        db.save_rules(aid, vf, rules)
    db.unify_models()
    return ins, same, ended, bad


# ============================================================
# 최신 모델 먼저 · 정책 변동 · 이미지 만들기
# ============================================================
_CAP_TOKEN = re.compile(r"(?<![0-9])(\d{2,4})\s?(GB|TB|G|T)(?![A-Za-z])", re.I)


def _cap_value(name):
    m = _CAP_TOKEN.search(name or "")
    if not m:
        return None
    v = int(m.group(1))
    return v * 1024 if m.group(2).upper().startswith("T") else v


def unify_capacity_names(names):
    """'갤럭시 Z 폴드7' 처럼 용량 없는 이름을, 같은 모델의 가장 작은 용량 이름('… 256GB')으로 합치기"""
    names = set(n for n in names if n)
    with_cap = {}
    for n in names:
        v = _cap_value(n)
        if v:
            base = _CAP_TOKEN.sub("", n).strip()
            base = re.sub(r"\s+", " ", base)
            with_cap.setdefault(base, []).append((v, n))
    mp = {}
    for n in names:
        if _cap_value(n) is None:
            b = re.sub(r"\s+", " ", n).strip()
            if b in with_cap:
                mp[n] = min(with_cap[b])[1]
    return mp


def model_year(name):
    """모델 이름으로 출시 연도 추정 (최신 모델을 위로 올리기 위해)"""
    n = (name or "").lower()
    for x, y in (("갤럭시", " "), ("galaxy", " "), ("아이폰", "iphone "), ("플립", "flip"), ("폴드", "fold"),
                 ("워치", "watch"), ("탭", "tab "), ("울트라", " ultra"), ("플러스", "+")):
        n = n.replace(x, y)
    for pat, base, fam in ((r"tab\s*s?\s*(\d{1,2})(?!\d)", 2014, 2),                  # 탭 S10 → 2024
                           (r"watch\s*(\d{1,2})(?!\d)", 2017, 2),                       # 워치8 → 2025
                           (r"(?:flip|fold)\s*(\d)(?!\d)", 2018, 0),                    # 플립7 → 2025
                           (r"iphone\s*(\d{2})(?!\d)", 2008, 0),                        # 아이폰17 → 2025
                           (r"(?<![a-z0-9])s\s?(\d{2})(?!\d)", 2000, 0),                # 갤럭시 S26 → 2026
                           (r"(?<![a-z0-9])a\s?(\d{2})(?!\d)", 2019, 1)):               # 갤럭시 A36·A16 → 2025 (끝자리)
        m = re.search(pat, n)
        if m:
            v = int(m.group(1))
            return (base + v % 10 if fam == 1 else base + v), fam
    return 0, 3


def model_sort_key(name, release=0):
    y, fam = model_year(name)
    return (-y, fam, -(release or 0), name)


def policy_changes(db, aid, vf):
    """이번에 저장한 정책이 직전 정책보다 오른/내린 것 → (오른 것, 내린 것, 새 모델 수)"""
    rows = db.q("""SELECT p.model, p.plan, p.join_type, p.rebate - p.deduction AS cur,
        (SELECT q.rebate - q.deduction FROM policies q WHERE q.agency_id=p.agency_id AND q.category=p.category
           AND q.carrier=p.carrier AND q.model=p.model AND q.plan=p.plan AND q.join_type=p.join_type
           AND q.valid_from<p.valid_from AND q.deleted=0 ORDER BY q.valid_from DESC, q.id DESC LIMIT 1) AS prev
        FROM policies p WHERE p.agency_id=? AND p.valid_from=? AND p.deleted=0""", (aid, vf))
    ups, downs, new = [], [], 0
    hidden = {m for m, r in db.model_meta().items() if r["hidden"]}
    for r in rows:
        if r["model"] in hidden:
            continue
        if r["prev"] is None:
            new += 1
        elif r["cur"] != r["prev"]:
            (ups if r["cur"] > r["prev"] else downs).append(
                (r["model"], re.sub(r"\s*\[.*?\]", "", r["plan"]), r["join_type"], r["prev"], r["cur"]))
    ups.sort(key=lambda x: -(x[4] - x[3]))
    downs.sort(key=lambda x: x[4] - x[3])
    return ups, downs, new


def change_text(ups, downs, n=5):
    short = {"신규": "신규", "번호이동": "번이", "기기변경": "기변"}
    f = lambda x: f"{x[0]} {short.get(x[2], x[2])} {x[3] / 10000:g}→{x[4] / 10000:g}만"
    return (("📈 " + ", ".join(f(x) for x in ups[:n])) if ups else "") + \
           ("   " if ups and downs else "") + (("📉 " + ", ".join(f(x) for x in downs[:n])) if downs else "")


def html_to_png(html_text, width, path):
    """HTML 표를 카톡에 올릴 이미지로"""
    doc = QTextDocument()
    doc.setDefaultFont(QFont("Malgun Gothic", 11))
    doc.setHtml(html_text)
    doc.setTextWidth(width)
    h = int(doc.size().height()) + 4
    img = QImage(int(width), max(h, 50), QImage.Format.Format_ARGB32)
    img.fill(QColor("white"))
    p = QPainter(img)
    doc.drawContents(p)
    p.end()
    img.save(path)
    return path


def desktop_dir():
    d = os.path.join(os.path.expanduser("~"), "Desktop")
    try:
        out = _ps("[Environment]::GetFolderPath('Desktop')").stdout.strip()
        if out and os.path.isdir(out):
            d = out
    except Exception:
        pass
    return d


# ============================================================
# 유지기간(환수 방지) · 부가 해지 가능일 · 할부 만료 재방문 · 자동 백업 · 정책 원본 보관
# ============================================================
TERM_KINDS = {"부가": "부가서비스 해지 가능", "요금제": "요금제 변경 가능", "회선": "회선 유지 끝(환수 없음)",
              "재방문": "할부 만료 → 재방문 안내"}


def _days_in(text):
    m = re.search(r"(\d{2,3})\s*일", text or "")
    return int(m.group(1)) if m else None


def derive_terms(db, s):
    """개통 한 건 → 지켜야 할 유지기간들 [(종류, 이름, 일수)]. 대리점 규칙 메모의 '93일' 같은 숫자를 우선 사용"""
    if (s["category"] or "무선") != "무선":
        return []
    d_add = to_int(db.get("keep_addon_days", "93")) or 93
    d_plan = to_int(db.get("keep_plan_days", "93")) or 93
    d_line = to_int(db.get("keep_line_days", "183")) or 183
    rules = db.effective_rules(s["sale_date"], s["agency_id"]) if s["agency_id"] else []
    flags = set(_split(s["flags"] or ""))
    terms, names, covered = [], set(), set()
    for r in rules:
        if r["kind"] in ADDON_MAIN and r["kind"] in flags:
            nm = clean_addon_name(r["name"])
            covered.add(r["kind"])
            if nm not in names:
                names.add(nm)
                terms.append(("부가", nm, _days_in(r["memo"]) or _days_in(r["name"]) or d_add))
    for k in flags:
        if k in ADDON_MAIN and k not in covered:          # 대리점 규칙엔 없지만 손님이 가입한 것
            terms.append(("부가", dict(FLAGS).get(k, k).replace(" 가입", ""), d_add))
    plan_days, line_days = d_plan, d_line
    for r in rules:
        if r["kind"] != "주의":
            continue
        txt = f"{r['name']} {r['memo']}"
        dd = _days_in(txt)
        if dd and "요금" in txt:
            plan_days = dd
        if dd and ("해지" in txt or "환수" in txt) and "요금" not in txt:
            line_days = max(line_days if line_days != d_line else 0, dd)
    terms.append(("요금제", "요금제 유지", plan_days))
    terms.append(("회선", "회선 유지(해지 시 환수)", line_days or d_line))
    months = s["months"] if "months" in s.keys() and s["months"] is not None else 24
    if months:
        terms.append(("재방문", f"할부 {months}개월 만료", int(months * 30.4)))
    return terms


def sync_terms(db, sale_id):
    s = db.one("SELECT * FROM sales WHERE id=?", (sale_id,))
    if not s:
        return
    old = {(r["kind"], r["name"]): r for r in db.q("SELECT * FROM sale_terms WHERE sale_id=?", (sale_id,))}
    db.con.execute("DELETE FROM sale_terms WHERE sale_id=?", (sale_id,))
    base = datetime.date.fromisoformat(s["sale_date"])
    for kind, name, days in derive_terms(db, s):
        o = old.get((kind, name))
        db.con.execute("INSERT INTO sale_terms(sale_id, kind, name, days, due, done, done_at) VALUES(?,?,?,?,?,?,?)",
                       (sale_id, kind, name, days, (base + datetime.timedelta(days=days)).isoformat(),
                        o["done"] if o else 0, o["done_at"] if o else None))
    db.commit()


def backfill_terms(db, limit=3000):
    ids = [r["id"] for r in db.q("SELECT s.id FROM sales s LEFT JOIN sale_terms t ON t.sale_id=s.id "
                                  "WHERE t.id IS NULL AND COALESCE(s.category,'무선')='무선' LIMIT ?", (limit,))]
    for i in ids:
        try:
            sync_terms(db, i)
        except Exception:
            pass
    return len(ids)


def auto_backup(db, force=False):
    """하루 한 번 데이터 백업 (최근 30개 보관) + 설정한 백업 폴더(구글 드라이브 등)에도 복사"""
    d = os.path.join(BASE_DIR, "backup")
    os.makedirs(d, exist_ok=True)
    fn = f"phone_policy_{datetime.date.today():%Y%m%d}.db"
    path = os.path.join(d, fn)
    if force or not os.path.exists(path):
        dst = sqlite3.connect(path + ".tmp")
        db.con.backup(dst)
        dst.close()
        os.replace(path + ".tmp", path)
        for f in sorted(glob.glob(os.path.join(d, "phone_policy_*.db")))[:-30]:
            try:
                os.remove(f)
            except Exception:
                pass
        ext = db.get("backup_dir", "").strip()
        if ext and os.path.isdir(ext):
            pc = re.sub(r"[^0-9A-Za-z가-힣_-]", "", db.get("pc_name", "") or socket.gethostname()) or "PC"
            shutil.copyfile(path, os.path.join(ext, f"폰정책_{pc}_{fn[13:]}"))
        db.set("last_backup", now())
    return path


def archive_policy_file(db, path, aid, vf):
    """대리점이 보낸 정책 원본 파일을 날짜·대리점별로 보관 (나중에 정산 다툼 때 증거)"""
    try:
        if not path or not os.path.isfile(path):
            return
        ag = db.one("SELECT name FROM agencies WHERE id=?", (aid,))
        folder = os.path.join(BASE_DIR, "정책원본보관", vf, re.sub(r'[\\/:*?"<>|]', "_", ag["name"] if ag else "대리점"))
        os.makedirs(folder, exist_ok=True)
        dst = os.path.join(folder, os.path.basename(path))
        if os.path.abspath(dst) != os.path.abspath(path):
            shutil.copyfile(path, dst)
        if not db.one("SELECT id FROM policy_files WHERE path=?", (dst,)):
            db.x("INSERT INTO policy_files(date, agency_id, name, path, created) VALUES(?,?,?,?,?)",
                 (vf, aid, os.path.basename(path), dst, now()))
    except Exception:
        pass


class ArchiveDialog(QDialog):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.setWindowTitle("🗂 정책 원본 보관함")
        self.resize(820, 520)
        lay = QVBoxLayout(self)
        lay.addWidget(banner("대리점이 보낸 정책 원본(엑셀·사진)이 <b>날짜·대리점별로</b> 자동 보관됩니다. "
                             "정산 금액이 안 맞을 때 그날 받은 정책표를 바로 열어 보여줄 수 있어요. 더블클릭 = 열기"))
        self.table = make_table(["정책 날짜", "대리점", "파일", "보관 시각"])
        rows = db.q("SELECT f.*, a.name AS agency FROM policy_files f LEFT JOIN agencies a ON a.id=f.agency_id "
                    "ORDER BY f.date DESC, f.id DESC")
        self.paths = [r["path"] for r in rows]
        fill_table(self.table, [[r["date"], r["agency"] or "", r["name"], r["created"]] for r in rows],
                   ids=list(range(len(rows))))
        self.table.cellDoubleClicked.connect(lambda r, _c: self.open_file(r))
        lay.addWidget(self.table, 1)
        b = QHBoxLayout()
        b.addWidget(btn("📂 보관 폴더 열기", lambda: QDesktopServices.openUrl(
            QUrl.fromLocalFile(os.path.join(BASE_DIR, "정책원본보관")))))
        b.addStretch()
        b.addWidget(btn("닫기", self.accept))
        lay.addLayout(b)

    def open_file(self, r):
        i = row_id(self.table, r)
        if i is not None and os.path.isfile(self.paths[i]):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.paths[i]))
        else:
            warn(self, "파일을 찾을 수 없습니다.")


# ============================================================
# 부가·차감 규칙 계산 / 요금구간 · 모델 이름 맞추기
# ============================================================
def fee_numbers(plan):
    """요금제 구간 글자에서 월정액(천원) 숫자들 추출. '115군 [115,130]' → [115,130], '61/69군' → [61,69]"""
    plan = plan or ""
    m = re.search(r"\[([0-9,\s/]+)\]", plan)
    if m:
        return [int(x) for x in re.findall(r"\d+", m.group(1))]
    nums = []
    for mm in re.finditer(r"(\d{2,3})(?![0-9.]*\s*(?:GB|G\b|MB|M\b|기가|%|개월|년|일|원|mm|인치))", plan, re.I):
        n = int(mm.group(1))
        if 10 <= n <= 300:
            nums.append(n)
    return nums


def plan_fee(plan):
    nums = fee_numbers(plan)
    return min(nums) if nums else None


_MODEL_REPL = [("갤럭시", ""), ("galaxy", ""), ("아이폰", "iphone"), ("프로맥스", "promax"), ("프로", "pro"),
               ("맥스", "max"), ("울트라", "ultra"), ("플러스", "+"), ("plus", "+"), ("플립", "flip"),
               ("폴드", "fold"), ("엣지", "edge"), ("에어", "air"), ("1tb", "1024"), ("테라", "1024")]


def norm_model(name):
    s = re.sub(r"\(.*?\)", "", (name or "")).lower()
    for a, b in _MODEL_REPL:
        s = s.replace(a, b)
    s = re.sub(r"(\d+)\s*(gb|g)\b", r"\1", s)
    return re.sub(r"[^0-9a-z가-힣+]", "", s)


def _split(v):
    return [x.strip() for x in re.split(r"[,/]", v or "") if x.strip()]


def rule_applies(rule, pol, join, fee):
    grps = _split(rule["grp"])
    pg = (pol["grp"] if "grp" in pol.keys() else "") or ""
    if grps and pg and pg not in grps:
        return False
    low = (pol["model"] or "").lower()
    kws = _split(rule["model_kw"])
    if kws and not any(k.lower() in low for k in kws):
        return False
    if any(k.lower() in low for k in _split(rule["model_ex"])):
        return False
    joins = _split(rule["joins"])
    if joins and join not in [norm_join(j, pol["category"]) or j for j in joins]:
        return False
    if fee is not None:
        if rule["fee_min"] is not None and fee < rule["fee_min"]:
            return False
        if rule["fee_max"] is not None and fee > rule["fee_max"]:
            return False
    elif rule["fee_min"] is not None or rule["fee_max"] is not None:
        return False
    return True


def man(n):
    """금액을 '만' 단위 짧은 표기로: -40000 → -4만"""
    v = n / 10000
    return f"{v:+g}만"


def apply_rules(rules, pol, join, fee, flags):
    """→ (가감 합계, 적용 설명 리스트, 주의 리스트)"""
    total, applied, warns = 0, [], []
    for r in rules:
        if not rule_applies(r, pol, join, fee):
            continue
        if r["kind"] == "주의":
            warns.append(r["name"] + (f"({r['memo']})" if r["memo"] else ""))
            continue
        amt = r["amt_yes"] if (r["kind"] == "자동" or flags.get(r["kind"])) else r["amt_no"]
        if amt:
            total += amt
            applied.append(f"{r['name']} {man(amt)}")
    return total, applied, warns


def norm_rule(d):
    """AI/파일에서 온 규칙 dict → 저장용 dict"""
    g = lambda *ks: next((d[k] for k in ks if k in d and d[k] is not None), "")
    kind = str(g("조건", "kind")).strip()
    if kind not in RULE_KINDS:
        kind = next((k for k in FLAG_KEYS if k in kind), "주의")

    def fee(v):
        return None if _blank(v) else to_int(v)
    return dict(name=str(g("규칙", "규칙 이름", "name")).strip() or "(이름 없음)", kind=kind,
                amt_yes=money_auto(g("해당", "해당 시", "amt_yes")), amt_no=money_auto(g("미해당", "미해당 시", "amt_no")),
                grp=str(g("그룹", "적용 그룹", "grp")).strip(), model_kw=str(g("모델포함", "모델 포함", "model_kw")).strip(),
                model_ex=str(g("모델제외", "모델 제외", "model_ex")).strip(), joins=str(g("가입유형", "joins")).strip(),
                fee_min=fee(g("요금최소", "요금 최소(천원)", "fee_min")), fee_max=fee(g("요금최대", "요금 최대(천원)", "fee_max")),
                memo=str(g("메모", "memo")).strip())


def pick_by_fee(cands, fee):
    """같은 대리점·모델·가입유형 정책들 중 손님 요금에 맞는 구간 고르기
    (요금 숫자가 없는 구분 — 예: 워치 '단독/번들' — 은 항상 포함)"""
    with_fee = [(p, fee_numbers(p["plan"])) for p in cands]
    nonum = [p for p, n in with_fee if not n]
    num = [(p, n) for p, n in with_fee if n]
    if fee is None or not num:
        return cands
    exact = [p for p, n in num if fee in n]
    if exact:
        return exact + nonum
    below = [(p, min(n)) for p, n in num if min(n) <= fee]
    if below:
        top = max(m for _, m in below)
        return [p for p, m in below if m == top] + nonum
    lowest = min(min(n) for _, n in num)      # 손님 요금이 제일 낮은 구간보다도 낮으면 제일 낮은 구간
    return [p for p, n in num if min(n) == lowest] + nonum


# ============================================================
# DB
# ============================================================
class DB:
    def __init__(self, path):
        self.path = path
        self.con = sqlite3.connect(path)
        self.con.row_factory = sqlite3.Row
        self.con.executescript(SCHEMA)
        self._migrate()
        self.con.commit()
        self.unify_models()

    def _migrate(self):
        def cols(t):
            return {r["name"] for r in self.q(f"PRAGMA table_info({t})")}
        if "phone" not in cols("sales"):
            self.con.execute("ALTER TABLE sales ADD COLUMN phone TEXT DEFAULT ''")
            self.con.execute("ALTER TABLE sales ADD COLUMN months INTEGER DEFAULT 24")
        if "source" not in cols("sales"):
            self.con.execute("ALTER TABLE sales ADD COLUMN source TEXT DEFAULT ''")
            self.con.execute("ALTER TABLE sales ADD COLUMN remote_id INTEGER")
        if "category" not in cols("policies"):
            self.con.execute("ALTER TABLE policies ADD COLUMN category TEXT NOT NULL DEFAULT '무선'")
        if "category" not in cols("sales"):
            self.con.execute("ALTER TABLE sales ADD COLUMN category TEXT NOT NULL DEFAULT '무선'")
        if "carrier" not in cols("agencies"):
            self.con.execute("ALTER TABLE agencies ADD COLUMN carrier TEXT DEFAULT ''")
        if "grp" not in cols("policies"):
            self.con.execute("ALTER TABLE policies ADD COLUMN grp TEXT DEFAULT ''")
        if "adjust" not in cols("sales"):
            self.con.execute("ALTER TABLE sales ADD COLUMN adjust INTEGER DEFAULT 0")
        if "flags" not in cols("sales"):
            self.con.execute("ALTER TABLE sales ADD COLUMN flags TEXT DEFAULT ''")
        self.con.execute("CREATE INDEX IF NOT EXISTS idx_pol_key2 ON policies"
                         "(agency_id, category, carrier, model, plan, join_type, valid_from)")
        self.con.execute("CREATE INDEX IF NOT EXISTS idx_sales_date ON sales(sale_date)")
        if not self.q("SELECT id FROM agencies LIMIT 1"):
            for car, short in (("SKT", "SK"), ("KT", "KT"), ("LGU+", "LG")):
                for i in range(1, 4):
                    self.con.execute("INSERT INTO agencies(name, carrier) VALUES(?,?)", (f"{short} 대리점{i}", car))

    def q(self, sql, p=()):
        return self.con.execute(sql, p).fetchall()

    def one(self, sql, p=()):
        return self.con.execute(sql, p).fetchone()

    def x(self, sql, p=(), commit=True):
        cur = self.con.execute(sql, p)
        if commit:
            self.con.commit()
        return cur.lastrowid

    def commit(self):
        self.con.commit()

    after_change = None      # 공유 기능이 연결하는 콜백 (데이터가 바뀌면 호출)

    def notify(self, what="data"):
        if self.after_change:
            try:
                self.after_change(what)
            except Exception:
                pass

    def is_staff_pc(self):
        return (self.get("role", "") or DEFAULT_ROLE) == "직원"

    def get(self, key, default=""):
        r = self.one("SELECT value FROM settings WHERE key=?", (key,))
        return r["value"] if r and r["value"] is not None else default

    def set(self, key, val):
        self.x("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, str(val)))

    def agencies(self, carrier=None):
        rows = self.q("SELECT * FROM agencies ORDER BY name")
        rows = sorted(rows, key=lambda a: (CARRIERS.index(a["carrier"]) if a["carrier"] in CARRIERS else 9, a["name"]))
        return [a for a in rows if not carrier or a["carrier"] == carrier]

    def agency_items(self, carrier=None):
        return [(a["id"], f"{a['name']}  ({a['carrier'] or '?'})") for a in self.agencies(carrier)]

    def agency_carrier(self, aid):
        r = self.one("SELECT carrier FROM agencies WHERE id=?", (aid,)) if aid else None
        return r["carrier"] if r else ""

    def carriers(self):
        extra = sorted({a["carrier"] for a in self.agencies() if a["carrier"] and a["carrier"] not in CARRIERS})
        return CARRIERS + extra

    def staff(self, active_only=False):
        sql = "SELECT * FROM staff" + (" WHERE active=1" if active_only else "") + " ORDER BY active DESC, name"
        return self.q(sql)

    def effective_policies(self, date, agency_id=None):
        sql = ("SELECT p.*, a.name AS agency FROM policies p JOIN agencies a ON a.id=p.agency_id "
               "WHERE p.valid_from<=?")
        params = [date]
        if agency_id:
            sql += " AND p.agency_id=?"
            params.append(agency_id)
        sql += " ORDER BY p.valid_from DESC, p.id DESC"
        seen, out = set(), []
        for r in self.q(sql, params):
            k = (r["agency_id"], r["category"], r["carrier"], r["model"], r["plan"], r["join_type"])
            if k in seen:
                continue
            seen.add(k)
            if not r["deleted"]:
                out.append(r)
        return out

    def effective_rules(self, date, agency_id):
        st = self.one("SELECT id, valid_from FROM rule_sets WHERE agency_id=? AND valid_from<=? "
                      "ORDER BY valid_from DESC, id DESC LIMIT 1", (agency_id, date))
        return self.q("SELECT * FROM rules WHERE set_id=? ORDER BY id", (st["id"],)) if st else []

    def save_rules(self, agency_id, valid_from, rules):
        for st in self.q("SELECT id FROM rule_sets WHERE agency_id=? AND valid_from=?", (agency_id, valid_from)):
            self.con.execute("DELETE FROM rules WHERE set_id=?", (st["id"],))
            self.con.execute("DELETE FROM rule_sets WHERE id=?", (st["id"],))
        sid = self.x("INSERT INTO rule_sets(agency_id, valid_from, created_at) VALUES(?,?,?)",
                     (agency_id, valid_from, now()), commit=False)
        for r in rules:
            self.con.execute("INSERT INTO rules(set_id,name,kind,amt_yes,amt_no,grp,model_kw,model_ex,joins,fee_min,"
                             "fee_max,memo) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                             (sid, r["name"], r["kind"], r["amt_yes"], r["amt_no"], r["grp"], r["model_kw"],
                              r["model_ex"], r["joins"], r["fee_min"], r["fee_max"], r["memo"]))
        self.commit()

    def unify_models(self):
        """저장된 정책의 모델 이름 중 용량 없는 것 → 같은 모델 기본 용량 이름으로 (SK·KT·LG 줄이 따로 노는 것 방지)"""
        try:
            names = [r["model"] for r in self.q("SELECT DISTINCT model FROM policies WHERE category='무선'")]
            mp = unify_capacity_names(names)
            for a, b in mp.items():
                self.con.execute("UPDATE policies SET model=? WHERE model=?", (b, a))
                self.con.execute("UPDATE sales SET model=? WHERE model=?", (b, a))
                self.con.execute("DELETE FROM model_meta WHERE model=?", (a,))
            if mp:
                self.con.commit()
            return len(mp)
        except Exception:
            return 0

    def model_meta(self):
        return {r["model"]: r for r in self.q("SELECT * FROM model_meta")}

    def set_meta(self, model, **kw):
        cur = self.one("SELECT * FROM model_meta WHERE model=?", (model,))
        d = dict(qty=None, hidden=0, fav=0) if not cur else dict(qty=cur["qty"], hidden=cur["hidden"], fav=cur["fav"])
        d.update(kw)
        self.x("INSERT OR REPLACE INTO model_meta(model, qty, hidden, fav) VALUES(?,?,?,?)",
               (model, d["qty"], int(bool(d["hidden"])), int(bool(d["fav"]))), commit=False)

    def visible_policies(self, date, agency_id=None):
        """화면에 보여줄 정책: 숨긴 모델 빼고, '재고 있는 모델만' 켜져 있으면 재고 0인 모델도 뺌"""
        meta = self.model_meta()
        stock_only = self.get("stock_only", "0") == "1"
        out = []
        for p in self.effective_policies(date, agency_id):
            m = meta.get(p["model"])
            if m is not None and m["hidden"]:
                continue
            if stock_only and p["category"] == "무선" and not (m is not None and (m["qty"] or 0) > 0):
                continue
            out.append(p)
        return out

    def find_policy(self, date, agency_id, category, carrier, model, plan, join_type):
        r = self.one(
            "SELECT * FROM policies WHERE agency_id=? AND category=? AND carrier=? AND model=? AND plan=? "
            "AND join_type=? AND valid_from<=? ORDER BY valid_from DESC, id DESC LIMIT 1",
            (agency_id, category, carrier, model, plan or "", join_type, date))
        return None if (r is None or r["deleted"]) else r


# ============================================================
# 직원 PC 공유 · 자동 업데이트 (공유 폴더 방식)
#   공유 폴더 안에 생기는 파일
#     phone_policy_manager.py : 최신 프로그램 (사장님 PC가 올림)
#     정책공유.json            : 대리점·정책·규칙·직원 목록 (사장님 PC가 올림)
#     개통_<PC이름>.json        : 각 직원 PC의 개통 내역 (직원 PC가 올림 → 사장님 PC가 모음)
# ============================================================
def ver_tuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", v or "0"))


def this_program():
    if FROZEN:
        return APP_PY
    return os.path.abspath(sys.argv[0] if sys.argv and sys.argv[0].endswith(".py") else __file__)


def this_program_script():
    return os.path.abspath(sys.argv[0] if sys.argv and sys.argv[0].endswith(".py") else __file__)


def read_version(path):
    try:
        with open(path, encoding="utf-8") as f:
            m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', f.read(6000), re.M)
        return m.group(1) if m else None
    except Exception:
        return None


def _atomic_write(path, text):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(text)
    os.replace(tmp, path)


class Share:
    SHARED_TABLES = ("agencies", "staff", "policies", "rule_sets", "rules")

    def __init__(self, db):
        self.db = db

    def folder(self):
        d = self.db.get("share_dir", "").strip()
        return d if d and os.path.isdir(d) else ""

    def pc(self):
        n = re.sub(r"[^0-9A-Za-z가-힣_-]", "", self.db.get("pc_name", "") or socket.gethostname())
        return n or "PC"

    def staff(self):
        return self.db.is_staff_pc()

    # --- 프로그램 버전 ---
    def remote_program(self):
        d = self.folder()
        if not d:
            return None, None
        path = os.path.join(d, "phone_policy_manager.py")
        if not os.path.isfile(path):
            return path, None
        try:
            with open(path, encoding="utf-8") as f:
                head = f.read(6000)
            m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', head, re.M)
            return path, (m.group(1) if m else None)
        except Exception:
            return path, None

    def publish_program(self):
        d = self.folder()
        if not d:
            return False
        dst = os.path.join(d, "phone_policy_manager.py")
        me = this_program()
        if os.path.exists(dst) and os.path.samefile(me, dst):
            return True
        shutil.copyfile(me, dst + ".tmp")
        os.replace(dst + ".tmp", dst)
        return True

    def install_update(self, path):
        """공유 폴더의 새 프로그램으로 내 파일 교체 (고장난 파일이면 설치 안 함)"""
        with open(path, encoding="utf-8") as f:
            src = f.read()
        compile(src, "update_check", "exec")
        if "APP_VERSION" not in src or "class MainWindow" not in src:
            raise RuntimeError("공유 폴더의 프로그램 파일이 올바르지 않습니다.")
        me = this_program()
        shutil.copyfile(me, me + ".bak")
        _atomic_write(me, src)
        return me

    # --- 정책 공유 (사장님 → 직원) ---
    def publish_policies(self):
        d = self.folder()
        if not d or self.staff():
            return False
        pack = {"stamp": datetime.datetime.now().isoformat(timespec="microseconds"), "from": self.pc(), "version": APP_VERSION,
                "settings": {k: self.db.get(k, "") for k in ("store_name", "target_margin", "round_unit")}}
        for t in self.SHARED_TABLES:
            pack[t] = [dict(r) for r in self.db.q(f"SELECT * FROM {t}")]
        _atomic_write(os.path.join(d, SHARE_POLICY_FILE), json.dumps(pack, ensure_ascii=False))
        return True

    def import_policies(self, force=False):
        """→ 새로 받았으면 받은 시각 문자열, 아니면 None"""
        d = self.folder()
        if not d or not self.staff():
            return None
        path = os.path.join(d, SHARE_POLICY_FILE)
        if not os.path.isfile(path):
            return None
        with open(path, encoding="utf-8") as f:
            pack = json.load(f)
        stamp = pack.get("stamp", "")
        if not force and stamp <= self.db.get("policy_stamp", ""):
            return None
        con = self.db.con
        cols = lambda t: [r["name"] for r in self.db.q(f"PRAGMA table_info({t})")]
        keep_ag = {r["agency_id"] for r in self.db.q("SELECT DISTINCT agency_id FROM sales")}
        pk_ag = {a["id"] for a in pack.get("agencies", [])}
        for aid in [a["id"] for a in self.db.q("SELECT id FROM agencies")]:
            if aid not in pk_ag and aid not in keep_ag:
                con.execute("DELETE FROM agencies WHERE id=?", (aid,))
        pk_st = {x["id"] for x in pack.get("staff", [])}
        con.execute(f"DELETE FROM staff WHERE id NOT IN ({','.join('?' * len(pk_st)) or 'NULL'})", tuple(pk_st))
        for t in ("policies", "rule_sets", "rules"):
            con.execute(f"DELETE FROM {t}")
        for t in self.SHARED_TABLES:
            tc = cols(t)
            for row in pack.get(t, []):
                keys = [k for k in row if k in tc]
                con.execute(f"INSERT OR REPLACE INTO {t}({','.join(keys)}) VALUES({','.join('?' * len(keys))})",
                            tuple(row[k] for k in keys))
        for k, v in pack.get("settings", {}).items():
            if v != "":
                con.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (k, str(v)))
        con.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('policy_stamp',?)", (stamp,))
        con.commit()
        return stamp

    # --- 개통 내역 (직원 → 사장님) ---
    def publish_sales(self):
        d = self.folder()
        if not d or not self.staff():
            return False
        rows = [dict(r) for r in self.db.q("SELECT * FROM sales WHERE COALESCE(source,'')=''")]
        _atomic_write(os.path.join(d, f"{SHARE_SALES_PREFIX}{self.pc()}.json"),
                      json.dumps({"pc": self.pc(), "stamp": now(), "sales": rows}, ensure_ascii=False))
        return True

    def collect_sales(self):
        """사장님 PC: 직원 PC들의 개통 내역을 모아서 합치기 (입금 처리 내용은 사장님 것 유지) → 반영 건수"""
        d = self.folder()
        if not d or self.staff():
            return 0
        tc = [r["name"] for r in self.db.q("PRAGMA table_info(sales)")]
        skip = {"id", "paid", "paid_date", "source", "remote_id"}
        n = 0
        for path in glob.glob(os.path.join(d, f"{SHARE_SALES_PREFIX}*.json")):
            try:
                with open(path, encoding="utf-8") as f:
                    pack = json.load(f)
            except Exception:
                continue
            src = pack.get("pc") or os.path.basename(path)
            if src == self.pc():
                continue
            ids = set()
            for s in pack.get("sales", []):
                ids.add(s["id"])
                keys = [k for k in s if k in tc and k not in skip]
                cur = self.db.one("SELECT id FROM sales WHERE source=? AND remote_id=?", (src, s["id"]))
                if cur:
                    self.db.con.execute(f"UPDATE sales SET {','.join(k + '=?' for k in keys)} WHERE id=?",
                                        (*[s[k] for k in keys], cur["id"]))
                else:
                    self.db.con.execute(
                        f"INSERT INTO sales({','.join(keys)},source,remote_id) VALUES({','.join('?' * len(keys))},?,?)",
                        (*[s[k] for k in keys], src, s["id"]))
                    n += 1
            old = [r["id"] for r in self.db.q("SELECT id, remote_id FROM sales WHERE source=?", (src,))
                   if r["remote_id"] not in ids]
            for i in old:
                self.db.con.execute("DELETE FROM sales WHERE id=?", (i,))
        self.db.commit()
        return n


# ============================================================
# 인터넷 자동 업데이트 (GitHub) — 카카오톡처럼 새 버전을 인터넷에서 받아 설치
#   저장소에 올라가는 것: phone_policy_manager.py (프로그램), version.json (버전 정보)
#   ※ 정책·개통·API키 같은 매장 데이터는 절대 올라가지 않습니다
# ============================================================
GH_API = "https://api.github.com"


def gh_repo(db):
    r = (db.get("gh_repo", "") or UPDATE_REPO or "").strip().strip("/")
    r = re.sub(r"^https?://github\.com/", "", r)
    return r if re.fullmatch(r"[\w.-]+/[\w.-]+", r) else ""


def _gh(url, token=None, method="GET", body=None, raw=False, timeout=15):
    h = {"User-Agent": "PhonePolicyManager", "Accept": "application/vnd.github.raw" if raw
         else "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"}
    if token:
        h["Authorization"] = f"Bearer {token}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    if data:
        h["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, method=method, headers=h)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        b = r.read()
    return b if raw else (json.loads(b.decode("utf-8")) if b else {})


def gh_get_file(repo, path, timeout=15):
    """공개 저장소 파일 받기 (API → 실패하면 raw 주소)"""
    try:
        return _gh(f"{GH_API}/repos/{repo}/contents/{path}", raw=True, timeout=timeout)
    except Exception:
        url = f"https://raw.githubusercontent.com/{repo}/HEAD/{path}?t={int(datetime.datetime.now().timestamp())}"
        with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "PhonePolicyManager"}),
                                    timeout=timeout) as r:
            return r.read()


def gh_put_file(repo, token, path, data, message):
    url = f"{GH_API}/repos/{repo}/contents/{path}"
    sha = None
    try:
        sha = _gh(url, token).get("sha")
    except urllib.error.HTTPError as ex:
        if ex.code != 404:
            raise
    body = {"message": message, "content": base64.b64encode(data).decode("ascii")}
    if sha:
        body["sha"] = sha
    return _gh(url, token, "PUT", body, timeout=60)


def gh_error(ex):
    if isinstance(ex, urllib.error.HTTPError):
        if ex.code == 401:
            return "GitHub 토큰이 틀렸거나 만료됐습니다. 설정 탭에서 토큰을 다시 넣어주세요."
        if ex.code == 403:
            return ("GitHub 권한이 없습니다. 토큰 만들 때 'Contents' 권한을 'Read and write'로 했는지,\n"
                    "그 저장소를 선택했는지 확인해 주세요.")
        if ex.code == 404:
            return "GitHub 저장소를 찾을 수 없습니다. 아이디/저장소 이름과, 저장소가 Public인지 확인해 주세요."
        return f"GitHub 오류 ({ex.code})"
    if isinstance(ex, urllib.error.URLError):
        return f"인터넷 연결을 확인해 주세요. ({ex.reason})"
    return str(ex)


def source_with_repo(src, repo):
    """프로그램 파일 안에 업데이트 저장소 주소를 박아 넣기 (직원 PC가 어디서 받을지 알도록)"""
    return re.sub(r'^UPDATE_REPO = ".*?"', f'UPDATE_REPO = "{repo}"', src, count=1, flags=re.M)


def source_for_staff(src):
    """직원에게 나가는 프로그램: 처음 켜면 '직원 PC' 화면 (토큰·배포 버튼 안 보임)"""
    return re.sub(r'^DEFAULT_ROLE = ".*?"', 'DEFAULT_ROLE = "직원"', src, count=1, flags=re.M)


def publish_source(db):
    """배포·설치파일에 쓰는 프로그램 본문 (저장소 주소 포함, 직원용)"""
    with open(this_program(), encoding="utf-8") as f:
        src = f.read()
    repo = gh_repo(db)
    if repo:
        src = source_with_repo(src, repo)
    if not db.is_staff_pc():
        db.set("role", "관리자")          # 사장님 PC는 역할을 확실히 기록 (직원용 배포본의 기본값에 안 휘둘리게)
    return source_for_staff(src)


def staff_block(parent, db, what="이 작업"):
    """직원 PC에서 막아야 하는 작업이면 안내하고 True"""
    if db.is_staff_pc():
        warn(parent, f"직원 PC에서는 {what}을(를) 할 수 없어요.\n사장님 PC에서 하면 직원 PC에 자동으로 전달됩니다.")
        return True
    return False


# ============================================================
# 위젯 헬퍼
# ============================================================
def info(parent, msg):
    QMessageBox.information(parent, APP_NAME, msg)


def warn(parent, msg):
    QMessageBox.warning(parent, APP_NAME, msg)


def ask(parent, msg):
    return QMessageBox.question(parent, APP_NAME, msg) == QMessageBox.StandardButton.Yes


def btn(text, fn, primary=False, big=False, tip=""):
    b = QPushButton(text)
    if primary:
        b.setProperty("primary", True)
    if big:
        b.setProperty("big", True)
    if tip:
        b.setToolTip(tip)
    b.clicked.connect(lambda _=False: fn())
    return b


def banner(text):
    lb = QLabel(text)
    lb.setObjectName("banner")
    lb.setWordWrap(True)
    return lb


def hint(text):
    lb = QLabel(text)
    lb.setStyleSheet("color:#667")
    lb.setWordWrap(True)
    return lb


def date_edit(fmt="yyyy-MM-dd", d=None):
    e = QDateEdit()
    e.setCalendarPopup(True)
    e.setDisplayFormat(fmt)
    e.setDate(d or QDate.currentDate())
    return e


def dstr(e):
    return e.date().toString("yyyy-MM-dd")


def mstr(e):
    return e.date().toString("yyyy-MM")


def money_spin(minimum=-50_000_000, maximum=50_000_000, step=10000):
    s = QSpinBox()
    s.setRange(minimum, maximum)
    s.setSingleStep(step)
    s.setGroupSeparatorShown(True)
    s.setSuffix(" 원")
    s.setAlignment(Qt.AlignmentFlag.AlignRight)
    s.setMinimumWidth(120)
    return s


def fill_combo(combo, items, all_label=None):
    cur = combo.currentData()
    combo.blockSignals(True)
    combo.clear()
    if all_label:
        combo.addItem(all_label, None)
    for id_, name in items:
        combo.addItem(name, id_)
    idx = combo.findData(cur) if cur is not None else -1
    combo.setCurrentIndex(idx if idx >= 0 else 0)
    combo.blockSignals(False)


def fill_text_combo(combo, items):
    cur = combo.currentText()
    combo.blockSignals(True)
    combo.clear()
    combo.addItems(items)
    i = combo.findText(cur)
    combo.setCurrentIndex(i if i >= 0 else 0)
    combo.blockSignals(False)


class NumItem(QTableWidgetItem):
    def __init__(self, n, text=None, editable=False):
        n = int(n or 0)
        super().__init__(text if text is not None else won(n))
        self.setData(Qt.ItemDataRole.UserRole, n)
        self.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        if not editable:
            self.setFlags(self.flags() & ~Qt.ItemFlag.ItemIsEditable)

    def __lt__(self, other):
        a = self.data(Qt.ItemDataRole.UserRole)
        b = other.data(Qt.ItemDataRole.UserRole)
        if isinstance(a, int) and isinstance(b, int):
            return a < b
        return super().__lt__(other)


def txt_item(s, editable=False):
    it = QTableWidgetItem("" if s is None else str(s))
    if not editable:
        it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
    return it


ID_ROLE = Qt.ItemDataRole.UserRole + 1


def make_table(headers, editable=False):
    t = QTableWidget(0, len(headers))
    t.setHorizontalHeaderLabels(headers)
    t.setAlternatingRowColors(True)
    t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    if not editable:
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    t.verticalHeader().setVisible(False)
    t.horizontalHeader().setStretchLastSection(True)
    return t


def fit_columns(t):
    """열 너비 정리: 금액 칸은 같은 너비로, 남는 공간은 글자 칸(비고·메모 등)만 가져가게"""
    t.resizeColumnsToContents()
    n, rows = t.columnCount(), min(t.rowCount(), 30)
    if not n:
        return
    numeric = []
    for c in range(n):
        items = [t.item(r, c) for r in range(rows)]
        items = [it for it in items if it is not None and it.text() not in ("", "-")]
        numeric.append(bool(items) and all(isinstance(it, NumItem) for it in items))
    for c in range(n):
        if numeric[c]:
            t.setColumnWidth(c, max(min(t.columnWidth(c), 110), 64))
        elif t.columnWidth(c) > 360:
            t.setColumnWidth(c, 360)
    t.horizontalHeader().setStretchLastSection(not numeric[-1])


def fill_table(t, rows, ids=None, colors=None, cell_colors=None, bold_rows=()):
    sort_col = t.horizontalHeader().sortIndicatorSection()
    sort_ord = t.horizontalHeader().sortIndicatorOrder()
    was_sorted = getattr(t, "_user_sorted", False)
    t.setSortingEnabled(False)
    t.clearContents()
    t.setRowCount(len(rows))
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            if isinstance(v, QTableWidgetItem):
                it = v
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                it = NumItem(v)
            else:
                it = txt_item(v)
            if c == 0 and ids is not None:
                it.setData(ID_ROLE, ids[r])
            if colors and colors[r] is not None:
                it.setBackground(QBrush(colors[r]))
            if cell_colors and (r, c) in cell_colors:
                it.setBackground(QBrush(cell_colors[(r, c)]))
            if r in bold_rows:
                f = it.font()
                f.setBold(True)
                it.setFont(f)
            t.setItem(r, c, it)
    fit_columns(t)
    if not getattr(t, "_sort_hooked", False):
        t._sort_hooked = True
        t.horizontalHeader().sectionClicked.connect(lambda _c, tt=t: setattr(tt, "_user_sorted", True))
    if was_sorted:
        t.setSortingEnabled(True)
        t.sortByColumn(sort_col, sort_ord)
    else:
        t.horizontalHeader().setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
        t.setSortingEnabled(True)


def row_id(t, r):
    it = t.item(r, 0)
    return it.data(ID_ROLE) if it else None


def selected_rows(t):
    return sorted({i.row() for i in t.selectedIndexes()})


def export_csv(parent, table, default_name):
    path, _ = QFileDialog.getSaveFileName(parent, "엑셀용 CSV 저장", default_name, "CSV (*.csv)")
    if not path:
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        cols = [c for c in range(table.columnCount()) if not table.isColumnHidden(c)]
        w.writerow([table.horizontalHeaderItem(c).text() for c in cols])
        for r in range(table.rowCount()):
            row = []
            for c in cols:
                it = table.item(r, c)
                if it is None:
                    row.append("")
                    continue
                v = it.data(Qt.ItemDataRole.UserRole)
                row.append(v if isinstance(v, int) and not isinstance(v, bool) else it.text())
            w.writerow(row)
    info(parent, f"저장했습니다. 엑셀로 열 수 있어요.\n{path}")


# ============================================================
# 가격 계산
# ============================================================
def compute_prices(db, date, agency_id, category, carrier, query, margin, unit):
    best = {}
    for p in db.visible_policies(date, agency_id):
        if category and p["category"] != category:
            continue
        if carrier and p["carrier"] != carrier:
            continue
        if query and query not in p["model"].lower() and query not in (p["plan"] or "").lower():
            continue
        k = (p["category"], p["carrier"], p["model"], p["plan"], p["join_type"])
        net = p["rebate"] - p["deduction"]
        if k not in best or net > best[k][0]:
            best[k] = (net, p)
    out = []
    for (cat, car, model, plan, join), (net, p) in best.items():
        support = max(0, net - margin)
        if unit > 1:
            support = support // unit * unit
        rp, ps = p["release_price"], p["public_subsidy"]
        out.append(dict(
            category=cat, carrier=car, model=model, plan=plan, join_type=join, agency=p["agency"],
            release_price=rp, public_subsidy=ps, rebate=p["rebate"], deduction=p["deduction"],
            net=net, support=support, hal_public=max(0, rp - ps - support),
            hal_select=max(0, rp - support), margin=net - support, conditions=p["conditions"] or ""))
    out.sort(key=order_key)
    return out


def build_price_html(store, date, rows):
    e = html.escape
    css = """
    body{font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;margin:28px;color:#222}
    h1{margin:0 0 4px;font-size:26px} .sub{color:#666;margin-bottom:10px}
    h2{background:#1f3a5f;color:#fff;padding:7px 14px;border-radius:6px;margin:26px 0 8px;font-size:18px;
       -webkit-print-color-adjust:exact;print-color-adjust:exact}
    h2.w{background:#2e6b3f}
    table{border-collapse:collapse;width:100%} th,td{border:1px solid #cfd5dc;padding:7px 9px;font-size:14px}
    th{background:#f0f3f7;-webkit-print-color-adjust:exact;print-color-adjust:exact}
    td.n{text-align:right;white-space:nowrap} td.m{font-weight:bold;background:#fafbfc}
    td.hl{color:#d62828;font-weight:bold} tr.gap td{border-top:2px solid #9aa6b2}
    .note{color:#777;font-size:12px;margin-top:22px;line-height:1.6}
    @media print{body{margin:10mm} h2{break-after:avoid} tr{break-inside:avoid}}
    """
    parts = [f"<!doctype html><html lang='ko'><head><meta charset='utf-8'><title>{e(store)} 가격표</title>"
             f"<style>{css}</style></head><body>",
             f"<h1>{e(store)} 가격표</h1><div class='sub'>{e(date)} 기준</div>"]
    carriers = CARRIERS + sorted({r["carrier"] for r in rows} - set(CARRIERS))
    for car in carriers:
        crow = [r for r in rows if r["carrier"] == car and r["category"] == "무선"]
        if crow:
            parts.append(f"<h2>{e(car)} 휴대폰</h2><table><tr><th>모델</th><th>요금제</th><th>가입유형</th>"
                         "<th>출고가</th><th>공시지원 선택 시<br>할부원금</th><th>선택약정 선택 시<br>할부원금</th></tr>")
            prev = None
            for r in crow:
                first = r["model"] != prev
                prev = r["model"]
                parts.append(
                    f"<tr class='{'gap' if first else ''}'><td class='m'>{e(r['model']) if first else ''}</td>"
                    f"<td>{e(r['plan'])}</td><td>{e(r['join_type'])}</td><td class='n'>{won(r['release_price'])}</td>"
                    f"<td class='n hl'>{won(r['hal_public'])}</td><td class='n'>{won(r['hal_select'])}</td></tr>")
            parts.append("</table>")
    for car in carriers:
        wrow = [r for r in rows if r["carrier"] == car and r["category"] == "유선"]
        if wrow:
            parts.append(f"<h2 class='w'>{e(car)} 인터넷·TV</h2><table><tr><th>상품</th><th>약정/조건</th>"
                         "<th>가입유형</th><th>사은품(최대)</th></tr>")
            prev = None
            for r in wrow:
                first = r["model"] != prev
                prev = r["model"]
                parts.append(
                    f"<tr class='{'gap' if first else ''}'><td class='m'>{e(r['model']) if first else ''}</td>"
                    f"<td>{e(r['plan'])}</td><td>{e(r['join_type'])}</td><td class='n hl'>{won(r['support'])}</td></tr>")
            parts.append("</table>")
    parts.append("<div class='note'>※ 휴대폰 금액은 할부원금이며 요금제·부가서비스 유지 등 조건이 있을 수 있습니다.<br>"
                 "※ 선택약정은 월 요금 25% 할인이 별도로 적용됩니다. 정책은 수시로 바뀌니 매장에 문의해 주세요.</div>")
    parts.append("</body></html>")
    return "".join(parts)


# ============================================================
# 탭: 정책 입력 (AI 자동읽기 포함)
# ============================================================
class DropArea(QLabel):
    def __init__(self, on_files, on_image, on_click):
        super().__init__("📂  대리점에서 받은 정책 파일(엑셀·사진·PDF)을 여기로 끌어다 놓으세요\n"
                         "(또는 여기를 눌러서 파일 선택)")
        self.on_files, self.on_image, self.on_click = on_files, on_image, on_click
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setAcceptDrops(True)
        self.setMinimumHeight(78)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._style(False)

    def _style(self, active):
        bg = "#dcebff" if active else "#f4f8fd"
        self.setStyleSheet(f"QLabel{{border:2px dashed #7a93b8;border-radius:10px;background:{bg};"
                           "color:#35507a;font-size:12pt;padding:6px}")

    def mousePressEvent(self, e):
        self.on_click()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls() or e.mimeData().hasImage():
            self._style(True)
            e.acceptProposedAction()

    def dragLeaveEvent(self, e):
        self._style(False)

    def dropEvent(self, e):
        self._style(False)
        md = e.mimeData()
        if md.hasUrls():
            self.on_files([u.toLocalFile() for u in md.urls() if u.isLocalFile()])
        elif md.hasImage():
            img = md.imageData()
            self.on_image(img if isinstance(img, QImage) else QImage(img))


class SheetPickDialog(QDialog):
    """엑셀 시트마다 읽는 방법 고르기: 정책표 / 부가·차감 조건만 / 안 읽기 (필요 없는 건 빼서 빠르게)"""
    def __init__(self, parent, fname, sheets, saved_modes):
        super().__init__(parent)
        self.setWindowTitle("시트 읽는 방법 고르기")
        self.resize(760, 520)
        lay = QVBoxLayout(self)
        lay.addWidget(banner(f"<b>{html.escape(fname)}</b> — 시트 {len(sheets)}개<br>"
                             "📊 <b>정책표</b> = 모델별 금액표 · 📝 <b>조건만</b> = 부가서비스·차감 공지 · 🚫 <b>안 읽기</b> = 공시지원금표·요금제표·3G 등<br>"
                             "프로그램이 추천해 둔 것이고, 이 대리점은 다음부터 같은 선택을 기억합니다. 🚫가 많을수록 빨라져요."))
        self.rows = []
        box = QWidget()
        g = QGridLayout(box)
        for i, (sn, grid) in enumerate(sheets):
            mode = (saved_modes or {}).get(sn) or classify_sheet(sn, grid)
            cb = QComboBox()
            for k, lab in SHEET_MODES:
                cb.addItem(lab, k)
            cb.setCurrentIndex(max(0, cb.findData(mode)))
            g.addWidget(QLabel(f"<b>{html.escape(sn)}</b>"), i, 0)
            g.addWidget(QLabel(f"숫자 {grid_numeric_count(grid)}개 · 글자 {grid_char_count(grid) // 1000}천자"), i, 1)
            g.addWidget(cb, i, 2)
            self.rows.append((sn, cb))
        g.setRowStretch(len(sheets), 1)
        from PySide6.QtWidgets import QScrollArea
        sa = QScrollArea()
        sa.setWidgetResizable(True)
        sa.setWidget(box)
        lay.addWidget(sa, 1)
        b = QHBoxLayout()
        b.addStretch()
        b.addWidget(btn("취소", self.reject))
        b.addWidget(btn("✅ 이대로 읽기", self.accept, primary=True))
        lay.addLayout(b)

    def modes(self):
        return {sn: cb.currentData() for sn, cb in self.rows}


def apply_sheet_modes(src, all_sheets, modes):
    src["data"] = [(sn, g) for sn, g in all_sheets if modes.get(sn) == "policy"]
    src["rule_sheets"] = [(sn, g) for sn, g in all_sheets if modes.get(sn) == "rules"]


def default_modes(db, aid, all_sheets):
    saved = db.get(f"sheet_modes_{aid}", "") if aid else ""
    try:
        sm = json.loads(saved) if saved else {}
    except ValueError:
        sm = {}
    return {sn: sm.get(sn) or classify_sheet(sn, g) for sn, g in all_sheets}


class PolicyEditTab(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.last_paths = set()
        self.sources = []
        self.worker = None
        self.loaded_agency = None   # '현재 정책 불러오기' 한 대리점
        self.from_file = False      # 파일/AI로 새 정책표를 채운 경우
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>대리점 정책이 오면 여기서 올리세요.</b>  ① 대리점·날짜 고르기 → ② 파일/사진 올리고 "
                             "<b>[AI로 읽기]</b> → ③ 표 확인 후 <b>[저장]</b>. 금액이 틀린 칸은 더블클릭해서 고치면 됩니다."))

        # ① 대리점 / 날짜
        s1 = QHBoxLayout()
        self.agency = QComboBox()
        self.agency.setMinimumWidth(220)
        self.date = date_edit()
        self.cat_hint = QComboBox()
        self.cat_hint.addItems(["자동 판단", "무선", "유선"])
        for w in [QLabel("<b>①</b> 대리점"), self.agency, QLabel("  정책 적용 날짜"), self.date,
                  QLabel("  종류"), self.cat_hint]:
            s1.addWidget(w)
        s1.addStretch()
        lay.addLayout(s1)

        # ② 파일 올리기
        box = QGroupBox("② 정책 파일 올리기 → AI가 자동으로 표로 옮김")
        bl = QHBoxLayout(box)
        left = QVBoxLayout()
        self.drop = DropArea(self.add_files, self.add_image, self.pick_files)
        left.addWidget(self.drop)
        row = QHBoxLayout()
        row.addWidget(btn("📁 파일 선택", self.pick_files))
        row.addWidget(btn("📋 캡처 붙여넣기 (Ctrl+V)", self.paste,
                          tip="카톡 사진을 화면에 띄우고 Win+Shift+S 로 캡처한 뒤 누르세요.\n엑셀 칸을 복사해서 붙여넣어도 됩니다."))
        row.addWidget(btn("목록 비우기", self.clear_sources))
        row.addStretch()
        left.addLayout(row)
        bl.addLayout(left, 3)
        right = QVBoxLayout()
        self.src_list = QListWidget()
        self.src_list.setMaximumHeight(90)
        right.addWidget(QLabel("올린 파일"))
        right.addWidget(self.src_list)
        self.b_ai = btn("🤖 AI로 읽기", self.run_ai, primary=True, big=True)
        right.addWidget(self.b_ai)
        self.b_cancel = btn("■ 읽기 중단", self.cancel_ai)
        self.b_cancel.hide()
        right.addWidget(self.b_cancel)
        self.ai_msg = ""
        self.ai_sec = 0
        self.ai_timer = QTimer(self)
        self.ai_timer.timeout.connect(self._tick)
        bl.addLayout(right, 2)
        lay.addWidget(box)
        self.status = QLabel("")
        self.status.setStyleSheet("color:#1f5fbf;font-weight:bold")
        self.status.setWordWrap(True)
        lay.addWidget(self.status)

        # ③ 표
        bar = QHBoxLayout()
        bar.addWidget(QLabel("<b>③</b> 확인·수정"))
        for text, fn in [("행 추가", self.add_row), ("선택 행 복제", self.dup_rows), ("선택 행 삭제", self.del_rows),
                         ("표 비우기", self.clear_grid), ("현재 저장된 정책 불러오기", self.load_current),
                         ("빈 입력양식 받기", self.export_template)]:
            bar.addWidget(btn(text, fn))
        bar.addStretch()
        bar.addWidget(btn("💾 저장", self.save, primary=True, big=True))
        lay.addLayout(bar)

        self.inner = QTabWidget()
        self.grid = QTableWidget(0, len(POLICY_COLS))
        self.grid.setHorizontalHeaderLabels(POLICY_COLS)
        self.grid.horizontalHeader().setStretchLastSection(True)
        self.grid.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.grid.setAlternatingRowColors(True)
        for c, w in enumerate([50, 55, 200, 150, 80, 95, 95, 95, 85, 260]):
            self.grid.setColumnWidth(c, w)
        self.inner.addTab(self.grid, "💰 정책 금액")

        rw = QWidget()
        rl = QVBoxLayout(rw)
        rl.setContentsMargins(0, 4, 0, 0)
        rl.addWidget(hint("대리점마다 다른 <b>부가서비스·보험·요금제·기변 조건</b>을 규칙으로 저장합니다. "
                          "'조건'이 부가/보험 등이면 손님이 가입했을 때(해당 시)와 안 했을 때(미해당 시) 금액이 달라지고, "
                          "'자동'은 대상이면 항상 적용, '주의'는 계산 없이 경고로만 보여줍니다. "
                          "금액은 원 단위(만원 단위 숫자는 자동 환산). 적용 그룹은 정책 금액 표의 '그룹' 이름과 맞춰 쓰면 됩니다."))
        rb = QHBoxLayout()
        rb.addWidget(btn("규칙 추가", self.add_rule))
        rb.addWidget(btn("선택 규칙 삭제", self.del_rules))
        rb.addStretch()
        rl.addLayout(rb)
        self.rgrid = QTableWidget(0, len(RULE_COLS))
        self.rgrid.setHorizontalHeaderLabels(RULE_COLS)
        self.rgrid.horizontalHeader().setStretchLastSection(True)
        self.rgrid.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.rgrid.setAlternatingRowColors(True)
        for c, w in enumerate([200, 95, 80, 80, 110, 90, 90, 90, 90, 90]):
            self.rgrid.setColumnWidth(c, w)
        self.rgrid.itemChanged.connect(lambda *_: setattr(self, "rules_dirty", True))
        rl.addWidget(self.rgrid)
        self.inner.addTab(rw, "📐 부가·차감 규칙 (0)")
        self.rules_dirty = False
        lay.addWidget(self.inner, 1)
        self.count_lbl = hint("")
        lay.addWidget(self.count_lbl)
        self.grid.model().rowsInserted.connect(lambda *_: self._count())
        self.grid.model().rowsRemoved.connect(lambda *_: self._count())

        sc = QShortcut(QKeySequence(QKeySequence.StandardKey.Paste), self)
        sc.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        sc.activated.connect(self.paste)

    # --- 규칙 표 ---
    def set_rules(self, rules, append=False):
        if not append:
            self.rgrid.setRowCount(0)
        for r in rules:
            self._rule_row(r)
        self.rules_dirty = True
        self._count()

    def _rule_row(self, r):
        i = self.rgrid.rowCount()
        self.rgrid.blockSignals(True)
        self.rgrid.insertRow(i)
        vals = [r["name"], None, won(r["amt_yes"]), won(r["amt_no"]), r["grp"], r["model_kw"], r["model_ex"],
                r["joins"], "" if r["fee_min"] is None else str(r["fee_min"]),
                "" if r["fee_max"] is None else str(r["fee_max"]), r["memo"]]
        for c, v in enumerate(vals):
            if c == 1:
                continue
            it = QTableWidgetItem(str(v or ""))
            if c in (2, 3):
                it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.rgrid.setItem(i, c, it)
        cb = QComboBox()
        for k in RULE_KINDS:
            lab = dict(FLAGS).get(k, {"자동": "항상 적용", "주의": "경고만(계산 X)"}.get(k, ""))
            cb.addItem(f"{k}  · {lab}", k)
        cb.setCurrentIndex(max(0, cb.findData(r["kind"])))
        cb.currentIndexChanged.connect(lambda *_: setattr(self, "rules_dirty", True))
        self.rgrid.setCellWidget(i, 1, cb)
        self.rgrid.blockSignals(False)

    def rules_from_grid(self):
        out = []
        for i in range(self.rgrid.rowCount()):
            t = lambda c: (self.rgrid.item(i, c).text().strip() if self.rgrid.item(i, c) else "")
            if not t(0):
                continue
            out.append(norm_rule({"규칙": t(0), "조건": self.rgrid.cellWidget(i, 1).currentData(),
                                  "해당": t(2), "미해당": t(3), "그룹": t(4), "모델포함": t(5), "모델제외": t(6),
                                  "가입유형": t(7), "요금최소": t(8), "요금최대": t(9), "메모": t(10)}))
        return out

    def add_rule(self):
        self._rule_row(dict(name="새 규칙", kind="부가", amt_yes=0, amt_no=0, grp="", model_kw="", model_ex="",
                            joins="", fee_min=None, fee_max=None, memo=""))
        self.rules_dirty = True
        self.inner.setCurrentIndex(1)
        self.rgrid.editItem(self.rgrid.item(self.rgrid.rowCount() - 1, 0))
        self._count()

    def del_rules(self):
        for r in reversed(selected_rows(self.rgrid)):
            self.rgrid.removeRow(r)
        self.rules_dirty = True
        self._count()

    def _count(self):
        self.inner.setTabText(0, f"💰 정책 금액 ({self.grid.rowCount()})")
        self.inner.setTabText(1, f"📐 부가·차감 규칙 ({self.rgrid.rowCount()})")
        self.count_lbl.setText(f"표에 {self.grid.rowCount()}줄 · 저장해야 반영됩니다. "
                               "금액은 원 단위 (1,000 미만 숫자를 넣으면 만원으로 보고 자동 환산)")

    def refresh(self):
        fill_combo(self.agency, self.db.agency_items())

    # --- 파일 올리기 ---
    def pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "정책 파일 선택", "", "정책 파일 (*.xlsx *.xlsm *.xls *.csv *.png *.jpg *.jpeg *.webp *.bmp *.pdf)")
        if paths:
            self.add_files(paths)

    def add_files(self, paths):
        errs = []
        for p in paths:
            try:
                direct = try_template_rows(p)
                if direct is not None:
                    prow, prules = direct
                    for vals in prow:
                        self.append_row(vals)
                    if prules:
                        self.set_rules(prules, append=self.rgrid.rowCount() > 0 and not self.from_file)
                    self.from_file = True
                    info(self, f"📥 '{os.path.basename(p)}' 은(는) 프로그램 양식이라 AI 없이 바로 넣었습니다.\n"
                               f"정책 {len(prow)}줄 · 부가·차감 규칙 {len(prules)}개\n\n"
                               "① 대리점·날짜가 맞는지 확인하고 [💾 저장]을 누르세요.")
                    continue
                src = load_source(p)
                if src["kind"] == "sheet":
                    allsh = list(src["data"])
                    aid = self.agency.currentData()
                    modes = default_modes(self.db, aid, allsh)
                    if len(allsh) > 1:
                        dlg = SheetPickDialog(self, os.path.basename(p), allsh, modes)
                        if dlg.exec() != QDialog.DialogCode.Accepted:
                            continue
                        modes = dlg.modes()
                        if aid:
                            self.db.set(f"sheet_modes_{aid}", json.dumps(modes, ensure_ascii=False))
                    apply_sheet_modes(src, allsh, modes)
                    if not src["data"] and not src["rule_sheets"]:
                        continue
                    src["name"] = f"{src['name']} (정책표 {len(src['data'])} · 조건 {len(src['rule_sheets'])})"
                self.last_paths.add(p)
                self._add_source(src)
            except Exception as ex:
                errs.append(f"{os.path.basename(p)}: {ex}")
        if errs:
            warn(self, "\n".join(errs))

    def add_image(self, img):
        try:
            self._add_source(dict(name=f"붙여넣은 이미지 {len(self.sources) + 1}", kind="image", data=image_parts(img)))
        except Exception as ex:
            warn(self, str(ex))

    def paste(self):
        md = QApplication.clipboard().mimeData()
        if md.hasImage():
            self.add_image(QApplication.clipboard().image())
        elif md.hasUrls():
            self.add_files([u.toLocalFile() for u in md.urls() if u.isLocalFile()])
        elif md.hasText() and md.text().strip():
            txt = md.text().strip("\r\n")
            lines = txt.splitlines()
            if sum(1 for l in lines if "\t" in l) >= 3:        # 엑셀에서 복사한 표
                grid = [l.split("\t") for l in lines]
                self._add_source(dict(name=f"붙여넣은 표 {len(self.sources) + 1}", kind="sheet",
                                      data=[("붙여넣기", grid)], images=[]))
            else:                                               # 카톡 등 글로 받은 정책
                self._add_source(dict(name=f"붙여넣은 글 {len(self.sources) + 1}", kind="text", data=txt))
        else:
            warn(self, "붙여넣을 내용이 없습니다.\n사진은 Win+Shift+S 로 캡처한 뒤, 엑셀은 칸을 복사한 뒤 눌러주세요.")

    def _add_source(self, src):
        self.sources.append(src)
        kind = {"text": "글", "image": "사진", "pdf": "PDF", "sheet": "엑셀"}[src["kind"]]
        extra = f" ({len(src['data'])}조각)" if src["kind"] == "image" and len(src["data"]) > 1 else ""
        self.src_list.addItem(f"{kind} · {src['name']}{extra}")
        self.status.setText(f"파일 {len(self.sources)}개 준비됨 → [🤖 AI로 읽기]를 누르세요")

    def clear_sources(self):
        self.sources = []
        self.src_list.clear()
        self.status.setText("")

    def run_ai(self):
        if not self.sources:
            return warn(self, "먼저 정책 파일이나 사진을 올려주세요.")
        aid = self.agency.currentData()
        if not aid:
            return warn(self, "① 대리점을 먼저 고르세요.")
        key = self.db.get("api_key", "").strip()
        if not key:
            return warn(self, "AI 키가 없습니다.\n'⚙ 설정' 탭에서 AI 키를 넣어주세요.\n"
                              "Claude · GPT · Gemini 중 아무 키나 넣으면 됩니다. (방법은 '❓ 사용법' 탭 4번)")
        hint_txt = self.cat_hint.currentText()
        self.b_ai.setEnabled(False)
        self.b_cancel.show()
        self.ai_msg = "🤖 AI 서버에 연결하는 중…"
        self.ai_sec = 0
        self._tick()
        self.ai_timer.start(1000)
        self.worker = AIWorker(key, pick_model(key, self.db.get("ai_model", "")), list(self.sources),
                               self.db.agency_carrier(aid), hint_txt if hint_txt in CATEGORIES else "자동 판단")
        self.worker.existing_models = sorted({p["model"] for p in self.db.effective_policies("9999-12-31")})
        self.worker.exclude = exclude_list(self.db.get("exclude_words", DEFAULT_EXCLUDE))
        self.worker.tier_mode = self.db.get("tier_mode", "top") or "top"
        self.worker.progress.connect(self._set_ai_msg)
        self.worker.done.connect(self.ai_done)
        self.worker.failed.connect(self.ai_failed)
        self.worker.start()

    def _set_ai_msg(self, m):
        self.ai_msg = m
        self._tick(advance=False)

    def _tick(self, advance=True):
        if advance and self.ai_timer.isActive():
            self.ai_sec += 1
        t = f"{self.ai_sec // 60}분 {self.ai_sec % 60:02d}초" if self.ai_sec >= 60 else f"{self.ai_sec}초"
        extra = ""
        if self.ai_sec >= 90 and "옮겨 적는" not in self.ai_msg:
            extra = "\n   (표가 크면 몇 분 걸릴 수 있어요. 3분 넘게 그대로면 [읽기 중단] 후 설정 탭 [연결 테스트]를 눌러보세요)"
        self.status.setText(f"{self.ai_msg}   ⏱ {t}{extra}")

    def _ai_stop_ui(self):
        self.ai_timer.stop()
        self.b_ai.setEnabled(True)
        self.b_cancel.hide()

    def cancel_ai(self):
        if self.worker:
            self.worker.cancel()
            for sig in (self.worker.progress, self.worker.done, self.worker.failed):
                try:
                    sig.disconnect()
                except Exception:
                    pass
        self._ai_stop_ui()
        self.status.setText("■ 읽기를 중단했습니다.")

    def ai_done(self, rows, rules, notes, truncated):
        self._ai_stop_ui()
        if not rows:
            self.status.setText("AI가 정책을 찾지 못했습니다.")
            return warn(self, "AI가 표에서 정책을 찾지 못했습니다.\n사진이 흐리면 더 선명하게 캡처하거나 엑셀 원본을 올려주세요.")
        append = False
        if self.grid.rowCount() and not ask(self, "표에 이미 내용이 있습니다. 지우고 새로 채울까요?\n(아니오 = 아래에 이어 붙이기)"):
            append = True
        if not append:
            self.grid.setRowCount(0)
        for r in rows:
            self.append_row(r)
        if rules:
            self.set_rules(rules, append=append)
        self.from_file = True
        self.clear_sources()
        start_row = self.grid.rowCount() - len(rows)
        sus = 0
        for r in range(max(0, start_row), self.grid.rowCount()):
            v = self.row_vals(r)
            reb = to_int(v[7])
            if abs(reb) > 2_000_000 or not v[2].strip() or reb < 0:
                sus += 1
                for c in range(len(POLICY_COLS)):
                    it = self.grid.item(r, c)
                    if it:
                        it.setBackground(QBrush(YELLOW))
        exact = getattr(self.worker, "exact", 0)
        how = (f"\n· 엑셀 {exact}줄은 원본 셀 숫자를 그대로 읽었습니다 (AI가 옮겨 적은 게 아니라 정확)." if exact else "")
        how += (f"\n· {len(rows) - exact}줄은 사진·글이라 AI가 옮겨 적었습니다 → 금액 몇 개만 눈으로 확인해 주세요."
                if len(rows) - exact > 0 else "")
        how += f"\n· 노란 줄 {sus}개: 마이너스·너무 큰 금액 등 확인이 필요한 줄" if sus else ""
        msg = (f"✅ 정책 {len(rows)}줄, 부가·차감 규칙 {len(rules)}개를 읽었습니다.{how}\n\n"
               "'💰 정책 금액'과 '📐 부가·차감 규칙' 두 표를 훑어보고 [💾 저장]을 누르세요.")
        self.status.setText(msg)
        extra = ""
        if truncated:
            extra += "\n\n⚠ 표가 커서 끝부분이 잘렸을 수 있습니다. 시트/사진을 나눠서 다시 읽혀주세요."
        if notes:
            extra += "\n\n📌 정책표 공통 참고사항:\n" + "\n".join(notes[:15])
        info(self, msg + extra)

    def ai_failed(self, msg):
        self._ai_stop_ui()
        self.status.setText("❌ AI 읽기 실패")
        warn(self, msg)

    # --- 표 조작 ---
    def set_row(self, r, vals):
        vals = (list(vals) + [""] * len(POLICY_COLS))[:len(POLICY_COLS)]
        for c, v in enumerate(vals):
            if c in MONEY_COLS:
                it = QTableWidgetItem(won(to_int(v)))
                it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            else:
                it = QTableWidgetItem("" if v is None else str(v))
            self.grid.setItem(r, c, it)

    def append_row(self, vals):
        r = self.grid.rowCount()
        self.grid.insertRow(r)
        self.set_row(r, vals)

    def row_vals(self, r):
        return [(self.grid.item(r, c).text() if self.grid.item(r, c) else "") for c in range(len(POLICY_COLS))]

    def add_row(self):
        n = self.grid.rowCount()
        if n:
            v = self.row_vals(n - 1)
            self.append_row([v[0], v[1], v[2], v[3], v[4], v[5], v[6], 0, 0, "", v[10]])
        else:
            car = self.db.agency_carrier(self.agency.currentData()) or "SKT"
            self.append_row(["무선", car, "", "", "신규", 0, 0, 0, 0, ""])
        self.grid.scrollToBottom()
        self.grid.editItem(self.grid.item(self.grid.rowCount() - 1, 2))

    def dup_rows(self):
        for r in selected_rows(self.grid):
            self.append_row(self.row_vals(r))

    def del_rows(self):
        for r in reversed(selected_rows(self.grid)):
            self.grid.removeRow(r)

    def clear_grid(self):
        if self.grid.rowCount() and not ask(self, "표를 모두 비울까요? (이미 저장된 정책은 그대로입니다)"):
            return
        self.grid.setRowCount(0)
        self.rgrid.setRowCount(0)
        self.rules_dirty = False
        self.loaded_agency = None
        self.from_file = False
        self._count()

    def load_current(self):
        aid = self.agency.currentData()
        if not aid:
            return warn(self, "대리점을 선택하세요.")
        if self.grid.rowCount() and not ask(self, "지금 표를 비우고 저장된 정책을 불러올까요?"):
            return
        rows = sorted(self.db.effective_policies(dstr(self.date), aid), key=order_key)
        self.grid.setRowCount(0)
        for p in rows:
            self.append_row([p["category"], p["carrier"], p["model"], p["plan"], p["join_type"], p["release_price"],
                             p["public_subsidy"], p["rebate"], p["deduction"], p["conditions"], p["grp"]])
        self.set_rules([dict(r) for r in self.db.effective_rules(dstr(self.date), aid)])
        self.rules_dirty = False
        self.loaded_agency = aid
        self.from_file = False
        if not rows:
            info(self, "이 날짜 기준 저장된 정책이 없습니다.")

    def export_template(self):
        path, _ = QFileDialog.getSaveFileName(self, "입력양식 저장", "정책_입력양식.csv", "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(POLICY_COLS)
            w.writerow(["무선", "SKT", "갤럭시 S25 256GB", "5GX 프라임 [89]", "번호이동", 1155000, 500000, 650000, 0, "", "프리미엄"])
            w.writerow(["유선", "SKT", "인터넷 500M+TV", "3년약정", "신규", 0, 0, 450000, 0, "", ""])
            w.writerow([])
            w.writerow([RULE_MARK])
            w.writerow(RULE_COLS)
            w.writerow(["부가서비스 미유치", "부가", 0, -40000, "프리미엄", "", "", "", "", "", ""])
            w.writerow(["폰보험", "보험", 10000, -30000, "", "", "", "", "", "", ""])
            w.writerow(["115군 기변 차감", "자동", -10000, 0, "", "", "", "기기변경", 115, "", ""])
            w.writerow(["93일 요금제 유지", "주의", 0, 0, "", "", "", "", "", "", "미유지 시 -8만 환수"])
        info(self, f"양식을 저장했습니다. 엑셀로 채운 뒤 위에 끌어다 놓으면 됩니다.\n{path}")

    # --- 저장 ---
    def save(self):
        if staff_block(self, self.db, "정책 저장"):
            return
        aid = self.agency.currentData()
        if not aid:
            return warn(self, "대리점을 선택하세요.")
        vf = dstr(self.date)
        items, bad = {}, []
        for r in range(self.grid.rowCount()):
            v = self.row_vals(r)
            if not any(x.strip() for x in v):
                continue
            cat = norm_category(v[0]) or "무선"
            car = norm_carrier(v[1]) or v[1].strip()
            model, plan = v[2].strip(), v[3].strip()
            join = norm_join(v[4], cat) or (v[4].strip() if cat == "유선" else "")
            if not (car and model and join):
                bad.append(str(r + 1))
                continue
            items[(cat, car, model, plan, join)] = dict(
                release_price=to_int(v[5]), public_subsidy=to_int(v[6]),
                rebate=to_int(v[7]), deduction=to_int(v[8]), conditions=v[9].strip(), grp=v[10].strip())
        if bad:
            return warn(self, "아래 줄의 '통신사 / 모델 / 가입유형'을 확인해 주세요.\n"
                              "(무선 가입유형: 신규·번호이동·기기변경 / 유선: 신규·재약정·전환)\n\n"
                              f"문제 있는 줄 번호: {', '.join(bad[:30])}")
        if not items:
            return warn(self, "저장할 내용이 없습니다.")

        ins = same = ended = 0
        for k, d in items.items():
            cur = self.db.find_policy(vf, aid, *k)
            if cur and all(cur[f] == d[f] for f in MONEY_FIELDS) and (cur["conditions"] or "") == d["conditions"] \
                    and (cur["grp"] or "") == d["grp"]:
                same += 1
                continue
            self.db.x("INSERT INTO policies(agency_id,category,carrier,model,plan,join_type,release_price,"
                      "public_subsidy,rebate,deduction,conditions,grp,valid_from,deleted,created_at) "
                      "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,0,?)",
                      (aid, *k, d["release_price"], d["public_subsidy"], d["rebate"], d["deduction"],
                       d["conditions"], d["grp"], vf, now()), commit=False)
            ins += 1
        self.db.commit()

        if self.loaded_agency == aid or self.from_file:
            cats = {k[0] for k in items}
            existing = {(p["category"], p["carrier"], p["model"], p["plan"], p["join_type"])
                        for p in self.db.effective_policies(vf, aid)}
            gone = [k for k in existing if k not in items and k[0] in cats]
            if gone:
                q = (f"새 정책표에 없는 기존 정책이 {len(gone)}개 있습니다.\n"
                     f"{vf}부터 '종료'로 처리할까요?\n\n(예 = 이제 안 파는 조건으로 처리 / 아니오 = 그대로 둠)")
                if ask(self, q):
                    for k in gone:
                        self.db.x("INSERT INTO policies(agency_id,category,carrier,model,plan,join_type,valid_from,"
                                  "deleted,created_at) VALUES(?,?,?,?,?,?,?,1,?)", (aid, *k, vf, now()), commit=False)
                        ended += 1
                    self.db.commit()
        rmsg = "부가·차감 규칙: 변경 없음 (이전 규칙 계속 사용)"
        if self.rules_dirty:
            rules = self.rules_from_grid()
            self.db.save_rules(aid, vf, rules)
            rmsg = f"부가·차감 규칙: {len(rules)}개 저장"
            self.rules_dirty = False
        self.loaded_agency = aid
        self.from_file = False
        self.db.notify("policy")
        for p in getattr(self, "last_paths", set()):
            archive_policy_file(self.db, p, aid, vf)
        self.last_paths = set()
        info(self, f"✅ 저장했습니다. ({self.agency.currentText()} · {vf}부터 적용)\n\n"
                   f"새로 들어가거나 바뀐 정책: {ins}개\n변동 없음: {same}개\n종료 처리: {ended}개\n{rmsg}")


# ============================================================
# 탭: 🔎 정책마진 조회 — 모델 하나 검색하면 SK·KT·LG 대리점 전부의 마진을 한눈에
# ============================================================
CHOICE_FLAGS = ("부가", "필링", "보험", "결합", "카드")      # 손님이 가입하면 매장이 더 받는 것들


def margin_breakdown(rules, pol, join, fee):
    """대리점 규칙 중 이 정책에 걸리는 것들: (부가 다 하면 가감, 부가 없이 가감, 설명 리스트, 주의 리스트)"""
    all_on = {k: True for k in CHOICE_FLAGS}
    up, _, warns = apply_rules(rules, pol, join, fee, all_on)
    down, _, _ = apply_rules(rules, pol, join, fee, {})
    desc = []
    for r in rules:
        if r["kind"] == "주의" or not rule_applies(r, pol, join, fee):
            continue
        if r["kind"] == "자동":
            if r["amt_yes"]:
                desc.append(f"{r['name']} {man(r['amt_yes'])}")
        elif r["kind"] in CHOICE_FLAGS:
            desc.append(f"{r['name']}: 하면 {man(r['amt_yes']) if r['amt_yes'] else '0'} / 안하면 "
                        f"{man(r['amt_no']) if r['amt_no'] else '0'}")
        else:
            desc.append(f"{r['name']}({dict(FLAGS).get(r['kind'], r['kind'])}) {man(r['amt_yes'])}")
    return up, down, desc, warns


# ============================================================
# 탭: ⭐ 한눈에 (첫 화면) — 즐겨찾기 모델 3사 최고 금액 + 최근 정책 변동
# ============================================================
# ============================================================
# 탭: ⏰ 알림 — 부가 해지 가능일 · 유지기간(환수 주의) · 할부 만료 재방문
# ============================================================
class AlertTab(QWidget):
    HEAD = ["기한", "D-day", "할 일", "고객", "연락처", "뒷4", "모델", "통신사", "대리점", "개통일", "담당"]

    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>개통한 손님을 날짜별로 챙겨주는 화면</b>입니다. 개통 등록할 때 체크한 부가서비스와 대리점 조건(93일 유지 등)으로 "
                             "날짜가 자동 계산돼요.<br>① <b>안내할 손님</b>: 부가서비스 해지해도 되는 날·요금제 바꿔도 되는 날이 된 손님 → [📋 안내 문자 복사] "
                             "② <b>유지기간 중</b>: 아직 기간 안 끝난 손님 (이 손님이 해지·변경하러 오면 환수됨!) "
                             "③ <b>재방문</b>: 할부 끝나가는 손님 → 기변 영업"))
        self.summary = QLabel("")
        self.summary.setStyleSheet("font-size:12pt;font-weight:bold;color:#1f5fbf")
        lay.addWidget(self.summary)
        self.inner = QTabWidget()
        self.t_due = make_table(self.HEAD)
        self.t_keep = make_table(self.HEAD)
        self.t_rev = make_table(self.HEAD)
        for t, name in ((self.t_due, "📅 안내할 손님"), (self.t_keep, "⚠ 유지기간 중 (환수 주의)"), (self.t_rev, "🔁 할부 만료 재방문")):
            self.inner.addTab(t, name)
        lay.addWidget(self.inner, 1)
        b = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("고객 이름·뒷4자리 검색 (유지기간 중인지 확인할 때)")
        self.search.returnPressed.connect(self.refresh)
        b.addWidget(self.search)
        b.addWidget(btn("🔎 검색", self.refresh))
        b.addWidget(btn("📋 선택 손님 안내 문자 복사", self.copy_msg, primary=True))
        b.addWidget(btn("✔ 안내 완료 표시", self.mark_done))
        b.addWidget(btn("↩ 완료 취소", lambda: self.mark_done(False)))
        b.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.inner.currentWidget(), f"알림_{today()}.csv")))
        lay.addLayout(b)
        lay.addWidget(hint("유지일 기본값(부가 93일·요금제 93일·회선 183일)은 ⚙ 설정에서 바꿀 수 있고, 대리점 규칙 메모에 '○○일'이 있으면 그걸 씁니다. "
                           "고객 연락처·할부 개월은 📱 개통 등록에서 넣어주세요."))
        self.items = {}

    def refresh(self):
        backfill_terms(self.db)
        td = datetime.date.today()
        q = self.search.text().strip()
        sql = ("SELECT t.*, s.customer, s.phone, s.phone4, s.model, s.carrier, s.sale_date, a.name AS agency, st.name AS staff "
               "FROM sale_terms t JOIN sales s ON s.id=t.sale_id LEFT JOIN agencies a ON a.id=s.agency_id "
               "LEFT JOIN staff st ON st.id=s.staff_id")
        rows = self.db.q(sql + " ORDER BY t.due")
        if q:
            rows = [r for r in rows if q in (r["customer"] or "") or q in (r["phone4"] or "") or q in (r["phone"] or "")]
        due, keep, rev = [], [], []
        for r in rows:
            dd = (datetime.date.fromisoformat(r["due"]) - td).days
            if r["kind"] == "재방문":
                if -30 <= dd <= 60 and not r["done"]:
                    rev.append((r, dd))
            elif r["kind"] in ("부가", "요금제"):
                if r["done"] or dd < -30:          # 한 달 넘게 지난 건 이미 지난 일로 봄
                    continue
                (due if dd <= 3 else keep).append((r, dd))
            elif r["kind"] == "회선" and dd > 0:
                keep.append((r, dd))
        self.items = {}
        for t, lst in ((self.t_due, due), (self.t_keep, keep), (self.t_rev, rev)):
            data, colors, ids = [], [], []
            for r, dd in lst:
                ids.append(r["id"])
                self.items[r["id"]] = r
                label = {"부가": f"'{r['name']}' 해지 가능 안내", "요금제": "요금제 변경 가능 안내",
                         "회선": "회선 유지 중 (해지하면 환수)", "재방문": f"{r['name']} → 재방문 안내"}[r["kind"]]
                if t is self.t_keep and r["kind"] != "회선":
                    label = f"'{r['name']}' 유지 중" if r["kind"] == "부가" else "요금제 유지 중"
                data.append([r["due"], NumItem(dd, f"D{'-' if dd >= 0 else '+'}{abs(dd)}"), label, r["customer"] or "",
                             r["phone"] or "", r["phone4"] or "", r["model"] or "", r["carrier"] or "", r["agency"] or "",
                             r["sale_date"], r["staff"] or ""])
                colors.append(RED if (t is self.t_keep and dd <= 14) else (YELLOW if dd < 0 else (GREEN if dd <= 0 else None)))
            fill_table(t, data, ids=ids, colors=colors)
        self.inner.setTabText(0, f"📅 안내할 손님 ({len(due)})")
        self.inner.setTabText(1, f"⚠ 유지기간 중 ({len(keep)})")
        self.inner.setTabText(2, f"🔁 할부 만료 재방문 ({len(rev)})")
        today_n = sum(1 for _, dd in due if dd <= 0)
        self.summary.setText(f"⏰ 오늘 안내할 손님 {today_n}명 · 3일 안에 {len(due) - today_n}명 · 재방문 대상 {len(rev)}명")
        return today_n, len(rev)

    def _sel(self):
        t = self.inner.currentWidget()
        return [self.items[row_id(t, r)] for r in selected_rows(t) if row_id(t, r) in self.items]

    def copy_msg(self):
        sel = self._sel()
        if not sel:
            return warn(self, "표에서 손님 줄을 먼저 클릭하세요. (여러 명은 Ctrl+클릭)")
        store = self.db.get("store_name", "우리매장")
        msgs = []
        for r in sel:
            name = (r["customer"] or "고객") + "님"
            cc = {"SKT": "114(SK텔레콤)", "KT": "114(KT)", "LGU+": "114(LG유플러스)"}.get(r["carrier"], "114")
            if r["kind"] == "부가":
                body = (f"[{store}] {name}, 안녕하세요! {r['sale_date']}에 개통하신 {r['model']} 관련 안내드립니다.\n"
                        f"가입하셨던 '{r['name']}'은(는) 유지기간이 끝나서 이제 해지하셔도 됩니다. "
                        f"해지는 {cc}나 통신사 앱에서 가능하세요. 늘 감사합니다 😊")
            elif r["kind"] == "요금제":
                body = (f"[{store}] {name}, 안녕하세요! {r['model']} 요금제 유지기간이 끝나서 이제 원하시는 요금제로 바꾸셔도 됩니다. "
                        f"요금제 상담 필요하시면 편하게 연락 주세요 😊")
            elif r["kind"] == "재방문":
                body = (f"[{store}] {name}, 안녕하세요! 쓰고 계신 {r['model']} 할부가 {r['due']}쯤 끝나갑니다. "
                        f"새 폰으로 바꾸실 때 제일 좋은 조건으로 안내해 드릴게요. 편하게 들러주세요 🙌")
            else:
                body = f"(내부 확인용) {name} {r['model']} 회선 유지 {r['due']}까지 — 이전에 해지하면 환수"
            to = r["phone"] or (f"뒷자리 {r['phone4']}" if r["phone4"] else "")
            msgs.append((f"받는 사람: {to}\n" if to else "") + body)
        QApplication.clipboard().setText("\n\n────────\n\n".join(msgs))
        info(self, f"📋 안내 문자 {len(msgs)}개를 복사했습니다. 카톡·문자에 붙여넣기(Ctrl+V) 하세요.\n\n"
                   "보낸 뒤 [✔ 안내 완료 표시]를 누르면 목록에서 빠집니다.")

    def mark_done(self, done=True):
        sel = self._sel()
        if not sel:
            return warn(self, "표에서 손님 줄을 먼저 클릭하세요.")
        for r in sel:
            self.db.x("UPDATE sale_terms SET done=?, done_at=? WHERE id=?", (1 if done else 0, now() if done else None, r["id"]),
                      commit=False)
        self.db.commit()
        self.refresh()


class DashboardTab(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>⭐ 즐겨찾기 모델의 SK·KT·LG 최고 정산금</b>과 <b>최근 정책 변동</b>을 한 장으로 봅니다. "
                             "금액 = 부가·보험 다 했을 때 대리점이 주는 돈(만원), 초록 = 그 가입유형 1등 통신사. "
                             "즐겨찾기는 🔎 정책마진 조회나 📦 모델 관리에서 ⭐를 누르면 됩니다 (없으면 최신 모델 12개)."))
        r1 = QHBoxLayout()
        self.fee = QComboBox()
        self.fee.setEditable(True)
        self.disc = discount_combo()
        for w in [QLabel("요금(천원)"), self.fee, QLabel("  할인"), self.disc]:
            r1.addWidget(w)
        r1.addWidget(btn("🔄 새로고침", self.calc, primary=True))
        r1.addStretch()
        lay.addLayout(r1)
        self.changes = QTextBrowser()
        self.changes.setMaximumHeight(130)
        lay.addWidget(self.changes)
        self.addons = AddonPanel(db)
        lay.addWidget(self.addons)
        self.table = make_table(["모델"])
        lay.addWidget(self.table, 1)
        self.fee.activated.connect(lambda *_: self.calc())
        self.disc.currentIndexChanged.connect(lambda *_: self.calc())

    def refresh(self):
        pols = [p for p in self.db.visible_policies(today()) if p["category"] == "무선" and fee_numbers(p["plan"])]
        fees = sorted({min(fee_numbers(p["plan"])) for p in pols}, reverse=True)
        cur = self.fee.currentText()
        self.fee.blockSignals(True)
        self.fee.clear()
        self.fee.addItems([str(f) for f in fees] or ["115"])
        self.fee.setCurrentText(cur or (str(fees[0]) if fees else "115"))
        self.fee.blockSignals(False)
        self.calc()

    def calc(self):
        fee = to_int(self.fee.currentText()) or 115
        data = build_store_policy(self.db, today(), fee, JOIN_BY_CAT["무선"], {k: True for k in CHOICE_FLAGS}, 0, 1,
                                  disc=self.disc.currentData())
        meta = self.db.model_meta()
        favs = [m for m in data if meta.get(m) is not None and meta[m]["fav"]]
        models = sorted(favs or data, key=model_sort_key)[: (200 if favs else 12)]
        cars = sorted({c for m in data.values() for (c, _j) in m}, key=lambda c: CARRIERS.index(c) if c in CARRIERS else 9)
        joins = JOIN_BY_CAT["무선"]
        short = {"신규": "신규", "번호이동": "번이", "기기변경": "기변"}
        heads = ["모델"] + [f"{c} {short[j]}" for c in cars for j in joins] + ["번이 1등"]
        self.table.setSortingEnabled(False)
        self.table.setColumnCount(len(heads))
        self.table.setHorizontalHeaderLabels(heads)
        rows, cc = [], {}
        for r, m in enumerate(models):
            line = [("⭐ " if m in favs else "") + m]
            for j in joins:
                cand = [c for c in cars if (c, j) in data[m]]
                if cand:
                    b = max(cand, key=lambda c: data[m][(c, j)]["net"])
                    cc[(r, 1 + cars.index(b) * len(joins) + joins.index(j))] = GREEN
            for c in cars:
                for j in joins:
                    d = data[m].get((c, j))
                    if d:
                        it = NumItem(d["net"], f"{d['net'] / 10000:g}")
                        it.setToolTip(f"{d['agency']} · {d['plan']}")
                        line.append(it)
                    else:
                        line.append(txt_item("-"))
            cand = [c for c in cars if (c, "번호이동") in data[m]]
            if cand:
                b = max(cand, key=lambda c: data[m][(c, "번호이동")]["net"])
                line.append(f"{b} {data[m][(b, '번호이동')]['agency']} {data[m][(b, '번호이동')]['net'] / 10000:g}만")
            else:
                line.append("-")
            rows.append(line)
        fill_table(self.table, rows, cell_colors=cc)
        try:
            ch = json.loads(self.db.get("last_changes", "[]") or "[]")
        except ValueError:
            ch = []
        self.addons.fill(today())
        self.changes.setHtml("<b>📈 최근 정책 변동</b> " + (f"({html.escape(self.db.get('last_changes_date', ''))})<br>" if ch else
                             "— 아직 없음 (📥 정책 일괄 등록으로 새 정책을 올리면 여기에 오른/내린 모델이 나옵니다)")
                             + "<br>".join(html.escape(x) for x in ch[:12]))


# ============================================================
# 탭: 📦 모델 관리 — 재고 · 숨기기 · 즐겨찾기 (옛날 모델·재고 없는 모델 정리)
# ============================================================
class ModelsTab(QWidget):
    COLS = ["⭐ 즐겨찾기", "🙈 숨김", "재고", "모델", "출고가", "정책 있는 통신사", "숨길 후보 이유"]

    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>모델 정리</b>: 옛날 모델·재고 없는 모델은 <b>🙈 숨김</b>에 체크하면 모든 조회·판매정책·엑셀에서 빠집니다 "
                             "(새 정책을 올려도 계속 숨겨져 있어요). <b>재고</b> 칸에 수량을 적고 아래 '재고 있는 모델만 보기'를 켜면 "
                             "기계 있는 모델만 나옵니다. 최신 모델이 위에 옵니다."))
        r1 = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("모델 검색")
        self.stock_only = QCheckBox("📦 재고 있는 모델만 조회·판매정책에 보이기")
        r1.addWidget(self.search)
        r1.addWidget(btn("🧹 숨길 후보 자동 체크", self.auto_hide))
        r1.addWidget(btn("📦 재고 엑셀 올리기", self.import_stock))
        r1.addWidget(self.stock_only)
        r1.addStretch()
        r1.addWidget(btn("💾 저장", self.save, primary=True, big=True))
        lay.addLayout(r1)
        self.table = make_table(self.COLS, editable=True)
        lay.addWidget(self.table, 1)
        lay.addWidget(hint("재고 칸은 숫자만 (비워두면 '모름'). 재고 엑셀은 첫 줄에 '모델'과 '수량(재고)' 제목이 있으면 이름을 알아서 맞춥니다."))
        self.search.textChanged.connect(lambda *_: self.render())
        self.info = {}

    def refresh(self):
        self.stock_only.setChecked(self.db.get("stock_only", "0") == "1")
        info = {}
        for p in self.db.effective_policies(today()):
            if p["category"] != "무선":
                continue
            d = info.setdefault(p["model"], dict(release=0, cars=set(), nets=[]))
            d["release"] = max(d["release"], p["release_price"] or 0)
            d["cars"].add(p["carrier"])
            d["nets"].append(p["rebate"] - p["deduction"])
        self.info = info
        self.render()

    def _reasons(self, m, d):
        rs = []
        if not d["release"]:
            rs.append("출고가 없음(#N/A)")
        if d["nets"] and max(d["nets"]) < 0:
            rs.append("모든 정책 마이너스")
        y, _ = model_year(m)
        if y and y <= datetime.date.today().year - 3:
            rs.append(f"{y}년쯤 모델")
        return rs

    def render(self):
        meta = self.db.model_meta()
        q = norm_model(self.search.text())
        models = sorted((m for m in self.info if not q or q in norm_model(m)),
                        key=lambda m: model_sort_key(m, self.info[m]["release"]))
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(models))
        for r, m in enumerate(models):
            d, mt = self.info[m], meta.get(m)
            for c, on in ((0, mt is not None and mt["fav"]), (1, mt is not None and mt["hidden"])):
                it = QTableWidgetItem("")
                it.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                it.setCheckState(Qt.CheckState.Checked if on else Qt.CheckState.Unchecked)
                self.table.setItem(r, c, it)
            qty = "" if (mt is None or mt["qty"] is None) else str(mt["qty"])
            self.table.setItem(r, 2, txt_item(qty, editable=True))
            it = txt_item(m)
            it.setData(ID_ROLE, m)
            self.table.setItem(r, 3, it)
            self.table.setItem(r, 4, NumItem(d["release"]))
            self.table.setItem(r, 5, txt_item(" ".join(sorted(d["cars"], key=lambda c: CARRIERS.index(c) if c in CARRIERS else 9))))
            rs = self._reasons(m, d)
            it = txt_item(" · ".join(rs))
            if rs:
                it.setBackground(QBrush(YELLOW))
            self.table.setItem(r, 6, it)
        self.table.resizeColumnsToContents()

    def auto_hide(self):
        n = 0
        for r in range(self.table.rowCount()):
            if self.table.item(r, 6).text():
                self.table.item(r, 1).setCheckState(Qt.CheckState.Checked)
                n += 1
        info(self, f"숨길 후보 {n}개에 체크했습니다. 확인하고 필요한 건 체크를 풀고 [💾 저장]을 누르세요.")

    def import_stock(self):
        path, _ = QFileDialog.getOpenFileName(self, "재고 엑셀", "", "엑셀/CSV (*.xlsx *.xls *.csv)")
        if not path:
            return
        try:
            sheets = read_sheets(path)
        except Exception as ex:
            return warn(self, str(ex))
        keys = {norm_model(m): m for m in self.info}
        got, miss = 0, []
        for _n, data in sheets:
            hi, mp = map_headers(data, [["모델", "기종", "펫네임", "품명"], ["수량", "재고", "개수"]], required=0)
            if hi is None or 1 not in mp:
                continue
            for row in data[hi + 1:]:
                name = str(row[mp[0]]).strip() if mp[0] < len(row) else ""
                if not name:
                    continue
                qty = to_int(row[mp[1]]) if mp[1] < len(row) else 0
                k = norm_model(name)
                m = keys.get(k) or next((v for kk, v in keys.items() if k and (k in kk or kk in k)), None)
                if m:
                    self.db.set_meta(m, qty=qty)
                    got += 1
                else:
                    miss.append(name)
        self.db.commit()
        self.render()
        info(self, f"재고 {got}개 반영했습니다." + (f"\n\n못 맞춘 이름 {len(miss)}개:\n" + "\n".join(miss[:20]) if miss else ""))

    def save(self):
        for r in range(self.table.rowCount()):
            m = self.table.item(r, 3).data(ID_ROLE)
            qtxt = self.table.item(r, 2).text().strip()
            self.db.set_meta(m, fav=self.table.item(r, 0).checkState() == Qt.CheckState.Checked,
                             hidden=self.table.item(r, 1).checkState() == Qt.CheckState.Checked,
                             qty=(to_int(qtxt) if qtxt else None))
        self.db.set("stock_only", "1" if self.stock_only.isChecked() else "0")
        self.db.commit()
        self.db.notify("policy")
        info(self, "저장했습니다. 모든 조회·판매정책에 바로 반영됩니다.")


# ============================================================
# 탭: 🧾 손님 견적서 — 3사 할부원금·월 납부액 비교, 카톡 이미지
# ============================================================
class QuoteTab(QWidget):
    RATE = 0.059     # 통신사 할부 수수료(연)

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.last_html = ""
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>손님 견적서</b>: 모델·가입유형·요금제를 고르면 SK·KT·LG 각각 제일 좋은 조건으로 "
                             "<b>할부원금 · 월 할부금 · 월 요금 · 월 예상 납부액</b>을 계산합니다. [🖼 카톡용 이미지]로 손님에게 보내세요."))
        r1 = QHBoxLayout()
        self.model = QComboBox()
        self.model.setEditable(True)
        self.model.setMinimumWidth(260)
        self.join = QComboBox()
        self.join.addItems(JOIN_BY_CAT["무선"])
        self.join.setCurrentText("번호이동")
        self.fee = QSpinBox()
        self.fee.setRange(10, 300)
        self.fee.setValue(115)
        self.fee.setSuffix(" 천원")
        self.disc = discount_combo()
        self.months = QComboBox()
        self.months.addItems(["24", "36", "12"])
        self.customer = QLineEdit()
        self.customer.setPlaceholderText("손님 이름(선택)")
        for w in [QLabel("모델"), self.model, QLabel(" 가입"), self.join, QLabel(" 요금제 월"), self.fee,
                  QLabel(" 할인"), self.disc, QLabel(" 할부"), self.months, QLabel("개월"), self.customer]:
            r1.addWidget(w)
        r1.addStretch()
        lay.addLayout(r1)
        fb = QGroupBox("손님이 가입할 부가서비스 (체크한 만큼 지원금에 반영)")
        fl = QVBoxLayout(fb)
        self.flags = FlagBox()
        fl.addWidget(self.flags)
        lay.addWidget(fb)
        r2 = QHBoxLayout()
        r2.addWidget(btn("🧾 견적 계산", self.calc, primary=True, big=True))
        r2.addWidget(btn("🖼 카톡용 이미지", self.image, primary=True))
        r2.addStretch()
        lay.addLayout(r2)
        self.view = QTextBrowser()
        lay.addWidget(self.view, 1)

    def refresh(self):
        names = sorted({p["model"] for p in self.db.visible_policies(today()) if p["category"] == "무선"}, key=model_sort_key)
        cur = self.model.currentText()
        self.model.blockSignals(True)
        self.model.clear()
        self.model.addItems(names)
        if cur:
            self.model.setCurrentText(cur)
        self.model.blockSignals(False)

    def monthly(self, principal, n):
        if principal <= 0:
            return 0
        r = self.RATE / 12
        return int(round(principal * r / (1 - (1 + r) ** -n)))

    def calc(self):
        model = self.model.currentText().strip()
        join, fee, disc, n = self.join.currentText(), self.fee.value(), self.disc.currentData(), int(self.months.currentText())
        data = build_store_policy(self.db, today(), fee, [join], self.flags.get(),
                                  to_int(self.db.get("target_margin", "100000")),
                                  to_int(self.db.get("round_unit", "10000")) or 1, disc=disc)
        key = norm_model(model)
        m = next((k for k in data if norm_model(k) == key), None) or next((k for k in data if key and key in norm_model(k)), None)
        if not m:
            self.view.setHtml("")
            return warn(self, "이 모델의 정책이 없습니다.")
        cars = [c for c in CARRIERS if (c, join) in data[m]] + \
               [c for (c, j) in data[m] if j == join and c not in CARRIERS]
        colc = {"SKT": "#E8412C", "KT": "#222222", "LGU+": "#C4006B"}
        rows = {k: [] for k in ("출고가", "공시지원금", "매장 추가지원금", "할부원금", f"월 할부금({n}개월)", "월 요금", "월 예상 납부액")}
        totals = {}
        for c in cars:
            d = data[m][(c, join)]
            pub = 0 if disc == "선약" else d["public"]
            principal = max(0, d["release"] - pub - d["sup"])
            mfee = int(fee * 1000 * (0.75 if disc == "선약" else 1))
            inst = self.monthly(principal, n)
            totals[c] = inst + mfee
            for k, v in zip(rows, (d["release"], pub, d["sup"], principal, inst, mfee, inst + mfee)):
                rows[k].append(v)
        best = min(totals, key=totals.get) if totals else None
        store = html.escape(self.db.get("store_name", "우리매장"))
        who = html.escape(self.customer.text().strip())
        h = [f"<h2 style='color:#1f3a5f'>{store} 견적서</h2>",
             f"<p><b>{html.escape(m)}</b> · {join} · 요금제 월 {fee:,}천원 · {self.disc.currentText()}"
             f"{' · 선택약정 25% 할인' if disc == '선약' else ''}{' · ' + who + ' 고객님' if who else ''} · {today()}</p>",
             "<table border='1' cellspacing='0' cellpadding='7'><tr><th bgcolor='#eef1f5'></th>"]
        for c in cars:
            h.append(f"<th bgcolor='{colc.get(c, '#555')}'><font color='white'>{c}{' ⭐추천' if c == best else ''}</font></th>")
        h.append("</tr>")
        for k, vals in rows.items():
            tag = k == "월 예상 납부액"
            h.append(f"<tr><td bgcolor='#f7f7f7'><b>{k}</b></td>" + "".join(
                (f"<td align='right' bgcolor='#d8f3df'><b>{v:,}원</b></td>" if tag and c == best else
                 f"<td align='right'>{'<b>' if tag else ''}{v:,}원{'</b>' if tag else ''}</td>")
                for c, v in zip(cars, vals)) + "</tr>")
        h.append("</table><p style='color:#777'>※ 할부 수수료 연 5.9% 기준 예상 금액이며, 요금제·부가서비스 유지 조건이 있을 수 있습니다. "
                 "정확한 금액은 개통 시 안내드립니다.</p>")
        self.last_html = "".join(h)
        self.view.setHtml(self.last_html)

    def image(self):
        if not self.last_html:
            self.calc()
            if not self.last_html:
                return
        path = os.path.join(desktop_dir(), f"견적서_{norm_model(self.model.currentText())}_{today()}.png")
        html_to_png(self.last_html, 760, path)
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        info(self, f"✅ 견적서 이미지를 만들었습니다.\n{path}\n\n카톡으로 손님에게 보내면 됩니다.")


def _selected_models(table, col):
    rows = selected_rows(table)
    return sorted({table.item(r, col).text() for r in rows if table.item(r, col) and table.item(r, col).text()
                   and not table.item(r, col).text().startswith("【")})


def hide_selected(parent, table, col, after=None):
    ms = _selected_models(table, col)
    if not ms:
        return warn(parent, "표에서 숨길 모델 줄을 먼저 클릭하세요. (여러 개는 Ctrl+클릭)")
    if not ask(parent, "이 모델들을 숨길까요? 모든 조회·판매정책·엑셀에서 빠집니다.\n(📦 모델 관리 탭에서 다시 보이게 할 수 있어요)\n\n"
                       + "\n".join(ms[:15])):
        return
    for m in ms:
        parent.db.set_meta(m, hidden=1)
    parent.db.commit()
    if after:
        after()


def fav_selected(parent, table, col):
    ms = _selected_models(table, col)
    if not ms:
        return warn(parent, "표에서 즐겨찾기할 모델 줄을 먼저 클릭하세요.")
    for m in ms:
        parent.db.set_meta(m, fav=1)
    parent.db.commit()
    info(parent, "⭐ 즐겨찾기에 넣었습니다. '⭐ 한눈에' 탭 첫 화면에 나옵니다.\n\n" + "\n".join(ms[:15]))


class MarginTab(QWidget):
    JOINS = ("신규", "번호이동", "기기변경")
    SHORT = {"신규": "신규", "번호이동": "번이", "기기변경": "기변"}

    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>모델만 검색하면 SK·KT·LG 대리점 전부의 정책마진이 한 화면에</b> 나옵니다.<br>"
                             "가입유형마다 <b>정책표 기본</b> · <b>부가·보험 다 하면</b> · <b>부가 없이</b> 3가지 금액과, 대리점별 부가·차감 내역을 같이 보여줘요. "
                             "초록 = 그 모델·가입유형에서 제일 많이 주는 곳."))
        r1 = QHBoxLayout()
        self.search = QComboBox()
        self.search.setEditable(True)
        self.search.setMinimumWidth(320)
        self.search.lineEdit().setPlaceholderText("모델 검색  예: S26 / 플립7 / 아이폰17")
        f = self.search.font()
        f.setPointSize(13)
        self.search.setFont(f)
        self.fee = QComboBox()
        self.fee.setEditable(True)
        self.fee.setToolTip("손님 요금제 월정액(천원). 통신사마다 요금제 이름이 달라도 이 금액 기준으로 맞춤")
        self.date = date_edit()
        self.disc = discount_combo()
        for w in [QLabel("🔎"), self.search, QLabel("  요금(천원)"), self.fee, QLabel("  할인"), self.disc,
                  QLabel("  날짜"), self.date]:
            r1.addWidget(w)
        r1.addWidget(btn("조회", self.calc, primary=True, big=True))
        r1.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"정책마진_{dstr(self.date)}.csv")))
        r1.addWidget(btn("🙈 선택 모델 숨기기", lambda: hide_selected(self, self.table, 0, self.calc)))
        r1.addWidget(btn("⭐ 즐겨찾기", lambda: fav_selected(self, self.table, 0)))
        r1.addStretch()
        lay.addLayout(r1)
        self.summary = QTextBrowser()
        self.summary.setMaximumHeight(120)
        lay.addWidget(self.summary)
        heads = ["모델", "통신사", "대리점", "적용 요금 구간"]
        for j in self.JOINS:
            heads += [f"{self.SHORT[j]} 기본", f"{self.SHORT[j]} 부가다함", f"{self.SHORT[j]} 부가없이"]
        heads += ["부가·차감 내역", "⚠ 주의"]
        self.table = make_table(heads)
        lay.addWidget(self.table, 1)
        lay.addWidget(hint("기본 = 정책표 금액(고정 차감 반영) · 부가다함 = 부가서비스·필링·보험 등 손님이 전부 가입했을 때 · "
                           "부가없이 = 아무것도 안 했을 때 (요금제·기변 조건 같은 '자동' 가감은 둘 다 반영). 금액 단위: 만원. "
                           "단기기변·시니어 같은 특수 조건은 🏆 최적 대리점 탭에서 체크해서 보세요."))
        self.search.lineEdit().returnPressed.connect(self.calc)
        self.search.activated.connect(lambda *_: self.calc())

    def refresh(self):
        pols = [p for p in self.db.visible_policies(dstr(self.date)) if p["category"] == "무선"]
        fees = sorted({min(fee_numbers(p["plan"])) for p in pols if fee_numbers(p["plan"])}, reverse=True)
        cur = self.fee.currentText()
        self.fee.blockSignals(True)
        self.fee.clear()
        self.fee.addItems([str(x) for x in fees] or ["115"])
        self.fee.setCurrentText(cur or (str(fees[0]) if fees else "115"))
        self.fee.blockSignals(False)
        names = sorted({p["model"] for p in pols}, key=model_sort_key)
        txt = self.search.currentText()
        self.search.blockSignals(True)
        self.search.clear()
        self.search.addItems(names)
        self.search.setCurrentIndex(-1)
        self.search.setEditText(txt)
        self.search.blockSignals(False)

    def calc(self):
        q = norm_model(self.search.currentText())
        if not q:
            return warn(self, "모델을 입력하세요. 예: S26")
        fee = to_int(self.fee.currentText()) or 115
        date = dstr(self.date)
        pols = [p for p in self.db.visible_policies(date)
                if p["category"] == "무선" and q in norm_model(p["model"]) and fee_numbers(p["plan"])]
        exact = [p for p in pols if norm_model(p["model"]) == q]
        if exact:
            pols = exact
        if not pols:
            fill_table(self.table, [])
            self.summary.setHtml("")
            return warn(self, "그 모델의 정책이 없습니다. 모델 이름을 다르게 검색해 보세요. (예: S26, 플립7, 아이폰17)")
        rules = {}
        groups = {}
        for p in pols:
            groups.setdefault((p["model"], p["agency_id"]), {}).setdefault(p["join_type"], []).append(p)
        rows = []
        for (model, aid), by_join in groups.items():
            if aid not in rules:
                rules[aid] = self.db.effective_rules(date, aid)
            rec = dict(model=model, aid=aid, vals={}, desc=[], warns=[], plan="", carrier="", agency="")
            for j in self.JOINS:
                best = None
                for p in pick_by_fee(by_discount(by_join.get(j, []), self.disc.currentData()), fee):
                    if not fee_numbers(p["plan"]):
                        continue
                    base = p["rebate"] - p["deduction"]
                    up, down, desc, warns = margin_breakdown(rules[aid], p, j, fee)
                    if best is None or base + up > best[1]:
                        best = (base, base + up, base + down, desc, warns, p)
                if best:
                    rec["vals"][j] = best[:3]
                    p = best[5]
                    rec.update(plan=p["plan"], carrier=p["carrier"], agency=p["agency"])
                    for d in best[3]:
                        if d not in rec["desc"]:
                            rec["desc"].append(d)
                    for w in best[4]:
                        if w not in rec["warns"]:
                            rec["warns"].append(w)
            if rec["vals"]:
                rows.append(rec)
        # 모델별로, 통신사 순, 부가다함 번이 큰 순
        rows.sort(key=lambda r: (model_sort_key(r["model"]), CARRIERS.index(r["carrier"]) if r["carrier"] in CARRIERS else 9,
                                 -(r["vals"].get("번호이동", (0, 0, 0))[1])))
        best_cell = {}
        for r in rows:
            for j, v in r["vals"].items():
                for k in range(3):
                    key = (r["model"], j, k)
                    best_cell[key] = max(best_cell.get(key, -10 ** 12), v[k])
        data, cc = [], {}
        for i, r in enumerate(rows):
            line = [r["model"], r["carrier"], r["agency"], re.sub(r"\s*\[.*?\]", "", r["plan"])]
            for jn, j in enumerate(self.JOINS):
                v = r["vals"].get(j)
                for k in range(3):
                    if not v:
                        line.append(txt_item("-"))
                        continue
                    it = NumItem(v[k], f"{v[k] / 10000:g}")
                    line.append(it)
                    if v[k] == best_cell[(r["model"], j, k)] and len(rows) > 1:
                        cc[(i, 4 + jn * 3 + k)] = GREEN
                    elif v[k] < 0:
                        cc[(i, 4 + jn * 3 + k)] = RED
            line += [" / ".join(r["desc"]) or "-", " / ".join(r["warns"])]
            data.append(line)
        self.table._user_sorted = False
        fill_table(self.table, data, cell_colors=cc)
        # 요약: 모델·가입유형별 1등 (부가다함 / 부가없이)
        html_lines = []
        for model in sorted({r["model"] for r in rows}, key=model_sort_key):
            rs = [r for r in rows if r["model"] == model]
            parts = []
            for j in self.JOINS:
                c = [r for r in rs if j in r["vals"]]
                if not c:
                    continue
                a = max(c, key=lambda r: r["vals"][j][1])
                b = max(c, key=lambda r: r["vals"][j][2])
                parts.append(f"<b>{self.SHORT[j]}</b> 부가다함 {a['carrier']} {html.escape(a['agency'])} "
                             f"<b style='color:#1a7f37'>{a['vals'][j][1] / 10000:g}만</b> · 부가없이 {b['carrier']} "
                             f"{html.escape(b['agency'])} <b>{b['vals'][j][2] / 10000:g}만</b>")
            per_car = {}
            for r in rs:
                v = r["vals"].get("번호이동") or next(iter(r["vals"].values()))
                if r["carrier"] not in per_car or v[1] > per_car[r["carrier"]][0]:
                    per_car[r["carrier"]] = (v[1], r["agency"])
            cars = "  ".join(f"{c} {per_car[c][0] / 10000:g}만({html.escape(per_car[c][1])})"
                             for c in sorted(per_car, key=lambda c: CARRIERS.index(c) if c in CARRIERS else 9))
            html_lines.append(f"📱 <b>{html.escape(model)}</b> ({fee}요금)  —  통신사별 최고(번이·부가다함): {cars}<br>"
                              "&nbsp;&nbsp;&nbsp;&nbsp;" + "  |  ".join(parts))
        self.summary.setHtml("<br>".join(html_lines))


# ============================================================
# 탭: 정책 일괄 등록 (9개 대리점 파일 한꺼번에 → 버튼 한 번)
# ============================================================
def guess_carrier_from_src(src):
    """파일 안 글자로 통신사 추측 (파일 이름에 통신사가 없을 때)"""
    if src.get("kind") != "sheet":
        return None
    txt = " ".join(str(v) for _, g in src.get("data", [])[:3] for row in g[:60] for v in row if str(v).strip())
    score = {"SKT": 0, "KT": 0, "LGU+": 0}
    score["SKT"] += len(re.findall(r"5GX|베스트\d|T우주|티다문|에스케이|SKT|P_119|I_100", txt))
    score["KT"] += len(re.findall(r"초이스|티지밀|요고|\bKT\b|NK\b|NK_|NK\d|케이티|110K", txt))
    score["LGU+"] += len(re.findall(r"플러스 ?플랜|U\+|유플|LGU|엘지|데이터 ?플랜|115군", txt))
    best = max(score, key=score.get)
    return best if score[best] >= 2 else None


def guess_agency(fname, agencies, used, src=None):
    low = fname.lower().replace(" ", "")
    for a in agencies:
        n = a["name"].lower().replace(" ", "")
        if n and n in low:
            return a["id"]
    car = None
    if re.search(r"skt|sk|에스케이|티월드", low):
        car = "SKT"
    elif re.search(r"kt|케이티", low):
        car = "KT"
    elif re.search(r"lg|u\+|유플|엘지", low):
        car = "LGU+"
    if not car and src is not None:
        car = guess_carrier_from_src(src)
    if car:
        for a in agencies:
            if a["carrier"] == car and a["id"] not in used:
                return a["id"]
    return None


def file_hash(path, extra=""):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    h.update(("|" + extra + "|v3").encode("utf-8"))
    return h.hexdigest()


class BatchTab(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.items = []      # dict(path, name, src, sheets_all, hash)
        self.worker = None
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>대리점 정책 파일을 전부 먼저 올리고 → 버튼 한 번</b>이면 한꺼번에 읽어서 <b>대리점별로 자동 저장</b>합니다.<br>"
                             "① 파일 9개를 다 끌어다 놓기 → ② 파일마다 대리점이 맞는지 확인(파일 이름으로 자동 추측) → "
                             "③ <b>[🤖 전부 읽고 저장]</b>. 한 번 읽은 파일은 기억해서 다음엔 바로 끝납니다."))
        top = QHBoxLayout()
        self.date = date_edit()
        top.addWidget(QLabel("정책 적용 날짜"))
        top.addWidget(self.date)
        top.addWidget(QLabel("  📶 요금 구간"))
        self.tier = QComboBox()
        for k, lab in TIER_MODES:
            self.tier.addItem(lab, k)
        self.tier.setToolTip("SK처럼 요금 구간이 많은 정책표는 '제일 높은 구간만' 읽으면 훨씬 빠릅니다.")
        self.tier.currentIndexChanged.connect(lambda *_: self.db.set("tier_mode", self.tier.currentData()))
        top.addWidget(self.tier)
        top.addWidget(QLabel("  🚫 빼고 읽기"))
        self.exclude = QLineEdit()
        self.exclude.setMinimumWidth(260)
        self.exclude.setToolTip("여기 적은 말이 들어간 표·요금제·모델은 안 읽습니다 (빨라짐).\n예: 3G, 2G, 선불, 키즈, 태블릿, 워치\n쉼표로 구분")
        self.exclude.editingFinished.connect(self._save_ex)
        top.addWidget(self.exclude)
        top.addWidget(btn("📁 파일 추가", self.pick))
        top.addWidget(btn("선택 줄 빼기", self.remove_sel))
        top.addWidget(btn("목록 비우기", self.clear))
        top.addStretch()
        lay.addLayout(top)
        self.drop = DropArea(self.add_files, lambda img: None, self.pick)
        self.drop.setText("📂  대리점 정책 엑셀·PDF·사진 파일을 여기로 한꺼번에 끌어다 놓으세요 (여러 개 가능)")
        lay.addWidget(self.drop)
        self.table = make_table(["파일", "대리점 (확인·변경)", "읽을 시트", "상태"], editable=False)
        lay.addWidget(self.table, 1)
        b = QHBoxLayout()
        self.b_run = btn("🤖 전부 읽고 저장", self.run, primary=True, big=True)
        self.b_stop = btn("■ 중단", self.stop)
        self.b_stop.hide()
        b.addWidget(self.b_run)
        b.addWidget(self.b_stop)
        self.status = QLabel("")
        self.status.setStyleSheet("color:#1f5fbf;font-weight:bold")
        self.status.setWordWrap(True)
        b.addWidget(self.status, 1)
        lay.addLayout(b)
        self.result = QTextBrowser()
        self.result.setMaximumHeight(170)
        lay.addWidget(self.result)
        lay.addWidget(btn("🗂 정책 원본 보관함 (대리점이 보낸 파일 날짜별 보관)", lambda: ArchiveDialog(self, self.db).exec()))
        self.sec = 0
        self.msg = ""
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)

    def refresh(self):
        self.tier.blockSignals(True)
        self.tier.setCurrentIndex(max(0, self.tier.findData(self.db.get("tier_mode", "top") or "top")))
        self.tier.blockSignals(False)
        self.exclude.setText(self.db.get("exclude_words", DEFAULT_EXCLUDE))
        _CURRENT_EXCLUDE[0] = self.exclude.text()

    def _save_ex(self):
        self.db.set("exclude_words", self.exclude.text().strip())
        _CURRENT_EXCLUDE[0] = self.exclude.text()
        for it in self.items:
            if it["sheets_all"] is not None:
                self._auto_sheets(it)
        self.render()

    def _tick(self):
        self.sec += 1
        self.status.setText(f"{self.msg}   ⏱ {self.sec // 60}분 {self.sec % 60:02d}초")

    def pick(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "대리점 정책 파일 선택", "",
                                                "정책 파일 (*.xlsx *.xlsm *.xls *.csv *.png *.jpg *.jpeg *.webp *.pdf)")
        if paths:
            self.add_files(paths)

    def add_files(self, paths):
        errs = []
        ags = self.db.agencies()
        for p in paths:
            try:
                src = load_source(p)
            except Exception as ex:
                errs.append(f"{os.path.basename(p)}: {ex}")
                continue
            used = {it["agency"] for it in self.items}
            aid = guess_agency(os.path.basename(p), ags, used, src)
            it = dict(path=p, name=os.path.basename(p), src=src, agency=aid,
                      sheets_all=list(src["data"]) if src["kind"] == "sheet" else None)
            if src["kind"] == "sheet":
                self._auto_sheets(it)
            self.items.append(it)
        self.render()
        if errs:
            warn(self, "\n".join(errs))

    def _auto_sheets(self, it):
        it["modes"] = default_modes(self.db, it["agency"], it["sheets_all"])
        apply_sheet_modes(it["src"], it["sheets_all"], it["modes"])

    def render(self):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(self.items))
        ags = self.db.agencies()
        for r, it in enumerate(self.items):
            self.table.setItem(r, 0, txt_item(it["name"]))
            cb = QComboBox()
            cb.addItem("⚠ 대리점을 골라주세요", None)
            for a in ags:
                cb.addItem(f"{a['name']}  ({a['carrier']})", a["id"])
            cb.setCurrentIndex(max(0, cb.findData(it["agency"])))
            cb.currentIndexChanged.connect(lambda _i, it=it, cb=cb: self._set_agency(it, cb.currentData()))
            self.table.setCellWidget(r, 1, cb)
            if it["sheets_all"] is not None:
                b = QPushButton(f"📊{len(it['src']['data'])} 📝{len(it['src'].get('rule_sheets', []))} "
                                f"🚫{len(it['sheets_all']) - len(it['src']['data']) - len(it['src'].get('rule_sheets', []))} (바꾸기)")
                b.clicked.connect(lambda _=False, it=it: self._choose_sheets(it))
                self.table.setCellWidget(r, 2, b)
            else:
                self.table.setItem(r, 2, txt_item({"image": "사진", "pdf": "PDF", "text": "글"}.get(it["src"]["kind"], "")))
            self.table.setItem(r, 3, txt_item(it.get("state", "대기")))
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(1, max(self.table.columnWidth(1), 230))

    def _set_agency(self, it, aid):
        it["agency"] = aid
        if it["sheets_all"] is not None:
            self._auto_sheets(it)
        self.render()

    def _choose_sheets(self, it):
        dlg = SheetPickDialog(self, it["name"], it["sheets_all"], it.get("modes"))
        if dlg.exec() == QDialog.DialogCode.Accepted:
            it["modes"] = dlg.modes()
            apply_sheet_modes(it["src"], it["sheets_all"], it["modes"])
            if it["agency"]:
                self.db.set(f"sheet_modes_{it['agency']}", json.dumps(it["modes"], ensure_ascii=False))
            self.render()

    def remove_sel(self):
        for r in reversed(selected_rows(self.table)):
            self.items.pop(r)
        self.render()

    def clear(self):
        self.items = []
        self.render()
        self.result.clear()

    def stop(self):
        if self.worker:
            self.worker.cancel()
            for sig in (self.worker.progress, self.worker.batch_done, self.worker.failed):
                try:
                    sig.disconnect()
                except Exception:
                    pass
        self._ui_idle()
        self.status.setText("■ 중단했습니다.")

    def _ui_idle(self):
        self.timer.stop()
        self.b_run.setEnabled(True)
        self.b_stop.hide()

    def _key(self, it):
        car = self.db.agency_carrier(it["agency"])
        extra = (car + "|" + ",".join(sn for sn, _ in it["src"]["data"]) + "|r:" +
                 ",".join(sn for sn, _ in it["src"].get("rule_sheets", []))) if it["src"]["kind"] == "sheet" else car
        extra += "|ex:" + ",".join(exclude_list(self.exclude.text())) + "|t:" + (self.tier.currentData() or "top")
        return file_hash(it["path"], extra)

    def run(self):
        if staff_block(self, self.db, "정책 저장"):
            return
        self.db.set("exclude_words", self.exclude.text().strip())
        _CURRENT_EXCLUDE[0] = self.exclude.text()
        if not self.items:
            return warn(self, "먼저 대리점 정책 파일을 올려주세요.")
        miss = [it["name"] for it in self.items if not it["agency"]]
        if miss:
            return warn(self, "대리점이 정해지지 않은 파일이 있습니다. '대리점' 칸에서 골라주세요.\n\n" + "\n".join(miss))
        key = self.db.get("api_key", "").strip()
        cached, todo = {}, []
        for i, it in enumerate(self.items):
            it["hash"] = self._key(it)
            row = self.db.one("SELECT data FROM ai_cache WHERE hash=?", (it["hash"],))
            if row:
                cached[i] = tuple(json.loads(row["data"]))
                it["state"] = "✅ 전에 읽은 파일 (바로 적용)"
            else:
                todo.append((i, it["src"], self.db.agency_carrier(it["agency"])))
                it["state"] = "⏳ 읽는 중"
        if todo and not key:
            return warn(self, "AI 키가 없습니다. ⚙ 설정에서 Claude·GPT·Gemini 키 중 하나를 넣어주세요.")
        self.render()
        self.b_run.setEnabled(False)
        self.b_stop.show()
        self.sec = 0
        self.msg = f"🤖 {len(todo)}개 파일을 동시에 읽는 중…" if todo else "저장 중…"
        self.timer.start(1000)
        existing = sorted({p["model"] for p in self.db.effective_policies("9999-12-31")})
        self.worker = BatchWorker(key, pick_model(key, self.db.get("ai_model", "")), todo, cached, existing)
        self.worker.exclude = exclude_list(self.exclude.text())
        self.worker.tier_mode = self.tier.currentData() or "top"
        self.worker.progress.connect(lambda m: setattr(self, "msg", m))
        self.worker.batch_done.connect(self._done)
        self.worker.failed.connect(self._failed)
        self.worker.start()

    def _failed(self, msg):
        self._ui_idle()
        self.status.setText("❌ 읽기 실패")
        warn(self, msg)

    def _done(self, per):
        self._ui_idle()
        vf = dstr(self.date)
        lines = []
        all_changes = []
        for i, it in enumerate(self.items):
            rows, rules, notes, tr, exact = per.get(i, per.get(str(i), ([], [], [], False, 0)))
            if rows:
                self.db.x("INSERT OR REPLACE INTO ai_cache(hash, created, data) VALUES(?,?,?)",
                          (it["hash"], now(), json.dumps([rows, rules, notes, tr, exact], ensure_ascii=False)))
            ag = self.db.one("SELECT name, carrier FROM agencies WHERE id=?", (it["agency"],))
            if not rows:
                it["state"] = "❌ 정책을 못 찾음"
                why = " / ".join(html.escape(str(n)) for n in notes if "해석 실패" in str(n) or "응답" in str(n))[:400]
                lines.append(f"<b>{html.escape(ag['name'])}</b> ← {html.escape(it['name'])}: <span style='color:#d62828'>"
                             f"정책을 못 찾았습니다</span> {why or '(시트 선택 확인 · ai_log 폴더에 AI 응답 기록 있음)'}")
                continue
            ins, same, ended, bad = save_policy_set(self.db, it["agency"], vf, rows, rules, end_missing=True)
            archive_policy_file(self.db, it["path"], it["agency"], vf)
            ups, downs, newn = policy_changes(self.db, it["agency"], vf)
            ct = change_text(ups, downs)
            if ct:
                all_changes.append(f"{ag['name']}: {ct}")
            neg = sum(1 for r in rows if to_int(r[7]) < 0)
            it["state"] = f"✅ 저장 {len(rows)}줄"
            lines.append(f"<b>{html.escape(ag['name'])} ({ag['carrier']})</b> ← {html.escape(it['name'])}: "
                         f"정책 {len(rows)}줄 (원본 칸에서 정확히 {exact}줄) · 부가·차감 규칙 {len(rules)}개 · "
                         f"새로/바뀜 {ins} · 종료 {ended}" + (f" · <span style='color:#b8860b'>마이너스 {neg}줄 확인</span>" if neg else "")
                         + (" · ⚠ 표가 커서 일부 잘렸을 수 있음" if tr else "")
                         + (f"<br>&nbsp;&nbsp;&nbsp;{html.escape(ct)}" if ct else "")
                         + (f" · 🆕 새로 생긴 정책 {newn}줄" if newn and not ct else ""))
        if all_changes:
            self.db.set("last_changes", json.dumps(all_changes, ensure_ascii=False))
            self.db.set("last_changes_date", vf)
        self.render()
        self.db.notify("policy")
        self.result.setHtml("<br>".join(lines) + "<br><br>👉 이제 <b>📣 판매정책</b> 탭에서 우리 매장 판매정책을 만들고, "
                            "<b>🏆 최적 대리점 · 📶 구간별 비교</b>에서 SK·KT·LG를 비교하세요. "
                            "잘못 읽힌 대리점은 <b>📋 정책 입력</b>에서 그 대리점 → [현재 저장된 정책 불러오기]로 고칠 수 있습니다.")
        self.status.setText(f"✅ 끝! ({self.sec // 60}분 {self.sec % 60:02d}초)")


# ============================================================
# 탭: 📣 우리 매장 판매정책 만들기 (SK·KT·LG 대리점 9곳 중 제일 좋은 조건으로) → 엑셀로 직원 배포
# ============================================================
SHOW_MODES = [("고객지원금(추가지원)", "sup"), ("할부원금(공시지원)", "hal"), ("할부원금(선택약정)", "sel"),
              ("정산금(리베이트, 사장님용)", "net")]


def clean_addon_name(name):
    n = re.sub(r"\s*(미유치|유치|미가입|가입|미이용|차감|추가|시)\b", "", name or "").strip(" -·/")
    return n or (name or "")


ADDON_MAIN = ("부가", "필링", "보험")          # 손님이 가입하는 상품 (결합·카드는 따로)


def short_name(n, k=24):
    n = re.sub(r"\s+", " ", clean_addon_name(n)).strip()
    return n if len(n) <= k else n[:k - 1] + "…"


def addon_rows(db, date, kinds=ADDON_MAIN):
    """통신사·상품별로 묶은 부가서비스 목록: [dict(carrier, name, memo, kind, ags=[(대리점, 하면, 안하면)])]"""
    groups = {}
    for a in db.agencies():
        for r in db.effective_rules(date, a["id"]):
            if r["kind"] not in kinds:
                continue
            nm = clean_addon_name(r["name"])
            g = groups.setdefault((a["carrier"] or "?", nm), dict(carrier=a["carrier"] or "?", name=nm, kind=r["kind"],
                                                                  memo=r["memo"] or "", ags=[]))
            if not any(x[0] == a["name"] for x in g["ags"]):
                g["ags"].append((a["name"], r["amt_yes"], r["amt_no"]))
    rows = list(groups.values())
    rows.sort(key=lambda g: (CARRIERS.index(g["carrier"]) if g["carrier"] in CARRIERS else 9,
                             -max(abs(y) + abs(n) for _, y, n in g["ags"])))
    return rows


def addon_html(db, date, owner=True, per=8):
    """카톡 이미지용: 통신사 3칸 나란히, 상품 이름만 (많이 영향 주는 순 최대 per개)"""
    rows = addon_rows(db, date)
    if not rows:
        return ""
    cars = [c for c in CARRIERS if any(r["carrier"] == c for r in rows)]
    colc = {"SKT": "#E8412C", "KT": "#222222", "LGU+": "#C4006B"}
    h = ["<table border='1' cellspacing='0' cellpadding='5'><tr><th colspan='%d' bgcolor='#fff4cc'>📌 가입 권유할 부가서비스</th></tr><tr>"
         % len(cars)]
    for c in cars:
        h.append(f"<th bgcolor='{colc.get(c, '#555')}'><font color='white'>{c}</font></th>")
    h.append("</tr><tr>")
    for c in cars:
        items = [r for r in rows if r["carrier"] == c][:per]
        h.append("<td valign='top'>" + "<br>".join(f"• {html.escape(short_name(r['name']))}" for r in items) + "</td>")
    h.append("</tr></table>")
    return "".join(h)


def addon_text_lines(db, date, per=10):
    out = []
    rows = addon_rows(db, date)
    for c in CARRIERS:
        items = [short_name(r["name"]) for r in rows if r["carrier"] == c][:per]
        if items:
            out.append(f"{c}: " + ", ".join(items))
    return out


class AddonPanel(QWidget):
    """통신사별 부가서비스를 보기 좋게: SKT | KT | LGU+ 세 칸 나란히, 상품 한 줄씩 (자세한 조건은 마우스 올리면)"""
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.date = today()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        top = QHBoxLayout()
        self.title = QLabel("<b>📌 통신사별 부가서비스</b> (개통할 때 손님께 권유할 상품 · 영향 큰 순 · 마우스를 올리면 자세한 조건)")
        self.more = QCheckBox("결합·카드 조건도 보기")
        self.fold = QPushButton("접기 ▲")
        self.fold.clicked.connect(self.toggle)
        self.more.toggled.connect(lambda *_: self.fill())
        top.addWidget(self.title)
        top.addStretch()
        top.addWidget(self.more)
        top.addWidget(self.fold)
        lay.addLayout(top)
        self.boxes = QHBoxLayout()
        self.lists = {}
        for c in CARRIERS:
            col = QVBoxLayout()
            head = QLabel(f"<b style='color:white'>&nbsp;{c}&nbsp;</b>")
            head.setStyleSheet(f"background:{ {'SKT': '#E8412C', 'KT': '#222222', 'LGU+': '#C4006B'}[c] };"
                               "border-radius:4px;padding:3px")
            lw = QListWidget()
            lw.setMaximumHeight(150)
            col.addWidget(head)
            col.addWidget(lw)
            self.boxes.addLayout(col)
            self.lists[c] = lw
        self.body = QWidget()
        self.body.setLayout(self.boxes)
        lay.addWidget(self.body)

    def toggle(self):
        v = not self.body.isVisible()
        self.body.setVisible(v)
        self.fold.setText("접기 ▲" if v else "펼치기 ▼")

    def fill(self, date=None):
        if date:
            self.date = date
        owner = not self.db.is_staff_pc()
        kinds = ADDON_MAIN + (("결합", "카드") if self.more.isChecked() else ())
        rows = addon_rows(self.db, self.date, kinds)
        for c, lw in self.lists.items():
            lw.clear()
            items = [r for r in rows if r["carrier"] == c]
            if not items:
                lw.addItem("(없음)")
            for r in items:
                txt = short_name(r["name"], 26)
                if owner:
                    best = max(r["ags"], key=lambda x: abs(x[1]) + abs(x[2]))
                    eff = best[1] - best[2]
                    if eff:
                        txt += f"   ({man(eff)})"
                lw.addItem(txt)
                it = lw.item(lw.count() - 1)
                tip = [r["name"]]
                if r["memo"]:
                    tip.append("조건: " + r["memo"])
                if owner:
                    tip += [f"{ag}: 하면 {man(y) if y else '0'} / 안하면 {man(n) if n else '0'}" for ag, y, n in r["ags"]]
                it.setToolTip("\n".join(tip))


def build_store_policy(db, date, fee, join_list, flags, margin, unit, carriers=None, disc="공시"):
    """→ {model: {(carrier, join): dict(best agency...)}}"""
    pols = [p for p in db.visible_policies(date) if p["category"] == "무선" and fee_numbers(p["plan"])]
    groups = {}
    for p in pols:
        if p["join_type"] not in join_list or (carriers and p["carrier"] not in carriers):
            continue
        groups.setdefault((p["model"], p["carrier"], p["join_type"], p["agency_id"]), []).append(p)
    rules = {}
    out = {}
    rel = {}
    for p in pols:                       # 출고가가 비어 있는 대리점('별도확인' 등)은 다른 대리점 출고가로
        if p["release_price"]:
            rel[p["model"]] = max(rel.get(p["model"], 0), p["release_price"])
    for (model, car, join, aid), cands in groups.items():
        if aid not in rules:
            rules[aid] = db.effective_rules(date, aid)
        for p in pick_by_fee(by_discount(cands, disc), fee):
            if not fee_numbers(p["plan"]):
                continue
            adj, applied, warns = apply_rules(rules[aid], p, join, fee, flags)
            net = p["rebate"] - p["deduction"] + adj
            cell = out.setdefault(model, {}).get((car, join))
            if cell is None or net > cell["net"]:
                sup = max(0, net - margin)
                if unit > 1:
                    sup = sup // unit * unit
                rp = p["release_price"] or rel.get(model, 0)
                addons = []
                for rr in rules[aid]:
                    if rr["kind"] in ADDON_MAIN and flags.get(rr["kind"]) and rule_applies(rr, p, join, fee):
                        nm = clean_addon_name(rr["name"])
                        if nm not in addons:
                            addons.append(nm)
                out[model][(car, join)] = dict(addons=addons,
                    net=net, agency=p["agency"], plan=p["plan"], sup=sup, applied=applied, warns=warns,
                    release=rp, public=p["public_subsidy"],
                    hal=max(0, rp - p["public_subsidy"] - sup), sel=max(0, rp - sup))
    return out


class StorePolicyTab(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.data = {}
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>우리 매장 판매정책</b>: 모델마다 SK·KT·LG 대리점 9곳 중 <b>제일 많이 주는 곳</b>을 골라, "
                             "목표마진을 빼고 <b>손님에게 줄 지원금</b>을 계산합니다. 부가서비스·보험·요금제 차감까지 반영.<br>"
                             "다 되면 <b>[📤 직원용 판매정책 엑셀]</b>로 뽑아서 카톡으로 뿌리세요 (리베이트·마진·대리점은 안 보임)."))
        r1 = QHBoxLayout()
        self.date = date_edit()
        self.fee = QComboBox()
        self.fee.setEditable(True)
        self.margin = money_spin(0, 5_000_000)
        self.unit = QComboBox()
        for t, v in [("절사 없음", 1), ("천원 단위", 1000), ("만원 단위", 10000)]:
            self.unit.addItem(t, v)
        self.mode = QComboBox()
        for t, v in SHOW_MODES:
            self.mode.addItem(t, v)
        self.search = QLineEdit()
        self.search.setPlaceholderText("모델 검색")
        for w in [QLabel("날짜"), self.date, QLabel(" 요금 구간(천원)"), self.fee, QLabel(" 목표마진"), self.margin,
                  QLabel(" 지원금"), self.unit, QLabel(" 표에 보일 금액"), self.mode, self.search]:
            r1.addWidget(w)
        self.disc = discount_combo()
        for w in [QLabel(" 할인"), self.disc]:
            r1.addWidget(w)
        r1.addStretch()
        lay.addLayout(r1)
        fb = QGroupBox("판매 기준 손님 조건 (보통 '부가·보험 가입'을 기준으로 정책을 만듭니다)")
        fl = QVBoxLayout(fb)
        self.flags = FlagBox(on_change=lambda: self.data and self.calc())
        fl.addWidget(self.flags)
        lay.addWidget(fb)
        r2 = QHBoxLayout()
        r2.addWidget(btn("📣 판매정책 만들기", self.calc, primary=True, big=True))
        r2.addWidget(btn("📤 직원용 판매정책 엑셀", lambda: self.export(False), primary=True))
        r2.addWidget(btn("🔒 사장님용 엑셀 (대리점·정산금 포함)", lambda: self.export(True)))
        r2.addWidget(btn("🖼 카톡용 이미지", self.export_image, primary=True))
        r2.addWidget(btn("🙈 선택 모델 숨기기", lambda: hide_selected(self, self.table, 0, self.calc)))
        r2.addStretch()
        lay.addLayout(r2)
        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        self.summary.setStyleSheet("font-weight:bold;color:#1a7f37")
        lay.addWidget(self.summary)
        self.addons = AddonPanel(db)
        lay.addWidget(self.addons)
        self.table = make_table(["모델"])
        lay.addWidget(self.table, 1)
        lay.addWidget(hint("칸 = 그 통신사 대리점 3곳 중 제일 좋은 곳 기준 금액. 초록 = 그 모델·가입유형에서 통신사 1등. "
                           "칸에 마우스를 올리면 어느 대리점·어떤 요금 구간·어떤 조건이 적용됐는지 보입니다. "
                           "'-' = 그 통신사 대리점에 그 모델 정책이 없음."))
        self._init = False
        self.mode.currentIndexChanged.connect(lambda *_: self.data and self.render())
        self.search.returnPressed.connect(lambda: self.data and self.render())

    def refresh(self):
        fees = sorted({min(fee_numbers(p["plan"])) for p in self.db.visible_policies(dstr(self.date))
                       if p["category"] == "무선" and fee_numbers(p["plan"])}, reverse=True)
        cur = self.fee.currentText()
        self.fee.blockSignals(True)
        self.fee.clear()
        self.fee.addItems([str(f) for f in fees] or ["115", "95", "69"])
        self.fee.setCurrentText(cur if cur else (str(fees[0]) if fees else "115"))
        self.fee.blockSignals(False)
        if not self._init:
            self._init = True
            self.margin.setValue(to_int(self.db.get("target_margin", "100000")))
            self.unit.setCurrentIndex(max(0, self.unit.findData(to_int(self.db.get("round_unit", "10000")))))

    def calc(self):
        fee = to_int(self.fee.currentText()) or 115
        self.cur_fee = fee
        self.data = build_store_policy(self.db, dstr(self.date), fee, JOIN_BY_CAT["무선"], self.flags.get(),
                                       self.margin.value(), self.unit.currentData() or 1,
                                       disc=self.disc.currentData())
        if not self.data:
            fill_table(self.table, [])
            return warn(self, "정책이 없습니다. 먼저 📥 정책 일괄 등록에서 대리점 정책을 저장해 주세요.")
        self.render()

    def carriers(self):
        cs = {c for m in self.data.values() for (c, _j) in m}
        return sorted(cs, key=lambda c: CARRIERS.index(c) if c in CARRIERS else 9)

    def render(self):
        mode = self.mode.currentData()
        cars = self.carriers()
        joins = JOIN_BY_CAT["무선"]
        short = {"신규": "신규", "번호이동": "번이", "기기변경": "기변"}
        heads = ["모델"] + [f"{c} {short[j]}" for c in cars for j in joins] + ["비고 (1등 대리점 조건·가입할 부가서비스)"]
        q = norm_model(self.search.text())
        models = sorted((m for m in self.data if not q or q in norm_model(m)), key=model_sort_key)
        self.table.setSortingEnabled(False)
        self.table.setColumnCount(len(heads))
        self.table.setHorizontalHeaderLabels(heads)
        rows, cc = [], {}
        wins = {c: 0 for c in cars}
        for r, m in enumerate(models):
            line = [m]
            for j in joins:
                best_c = max((c for c in cars if (c, j) in self.data[m]), key=lambda c: self.data[m][(c, j)]["net"],
                             default=None)
                if best_c:
                    wins[best_c] += 1
                    cc[(r, 1 + cars.index(best_c) * len(joins) + joins.index(j))] = GREEN
            for c in cars:
                for j in joins:
                    d = self.data[m].get((c, j))
                    if not d:
                        line.append(txt_item("-"))
                        continue
                    v = d[mode]
                    it = NumItem(v, f"{v / 10000:g}만")
                    it.setToolTip(f"{d['agency']} · {d['plan']}\n정산금 {won(d['net'])}원 → 고객지원금 {won(d['sup'])}원\n"
                                  + ("\n".join(d["applied"]) if d["applied"] else "조건 가감 없음")
                                  + (f"\n⚠ {' / '.join(d['warns'])}" if d["warns"] else ""))
                    line.append(it)
            line.append(self.note(m, cars, owner=not self.db.is_staff_pc()))
            rows.append(line)
        fill_table(self.table, rows, cell_colors=cc)
        for c in range(1, len(heads) - 1):
            self.table.setColumnWidth(c, 66)
        self.addons.fill(dstr(self.date))
        self.summary.setText(f"📣 {self.cur_fee}요금 기준 · 모델 {len(models)}개 · 목표마진 {won(self.margin.value())}원   |   "
                             "1등 칸 수: " + "  ".join(f"{c} {wins[c]}" for c in cars))

    def note(self, m, cars, owner=True):
        """비고: 통신사마다 (번이 기준 1등 대리점의) 가입해야 할 부가서비스와 주의사항"""
        parts = []
        for c in cars:
            d = self.data[m].get((c, "번호이동")) or next((self.data[m][(cc, j)] for (cc, j) in self.data[m] if cc == c), None)
            if not d:
                continue
            bits = []
            if d.get("addons"):
                ad = [short_name(x, 16) for x in d["addons"]]
                bits.append("부가: " + ", ".join(ad[:3]) + (f" 외 {len(ad) - 3}" if len(ad) > 3 else ""))
            if owner and d.get("warns"):
                bits.append(f"⚠ 주의 {len(d['warns'])}건")
            if bits:
                parts.append(f"{c}{'(' + d['agency'] + ')' if owner else ''} " + " · ".join(bits))
        return "  |  ".join(parts) or "-"

    def export_image(self):
        if not self.data:
            self.calc()
            if not self.data:
                return
        store = html.escape(self.db.get("store_name", "우리매장"))
        cars = self.carriers()
        joins = JOIN_BY_CAT["무선"]
        short = {"신규": "신규", "번호이동": "번이", "기기변경": "기변"}
        colc = {"SKT": "#E8412C", "KT": "#222222", "LGU+": "#C4006B"}
        flag_txt = ", ".join(lab for k, lab in FLAGS if self.flags.get().get(k)) or "부가 조건 없음"
        h = [f"<h2 style='color:#1f3a5f'>{store} 판매정책</h2>",
             f"<p style='color:#555'>{dstr(self.date)} · 요금 {self.cur_fee}천원 구간 · {self.disc.currentText()} · "
             f"기준: {html.escape(flag_txt)} · 금액: 고객지원금(만원)</p>",
             addon_html(self.db, dstr(self.date), owner=False), "<br>",
             "<table border='1' cellspacing='0' cellpadding='5' style='border-collapse:collapse'>",
             "<tr><th rowspan='2' bgcolor='#1f3a5f'><font color='white'>모델</font></th>"]
        for c in cars:
            h.append(f"<th colspan='3' bgcolor='{colc.get(c, '#555')}'><font color='white'>{c}</font></th>")
        h[-1] = h[-1]      # (모델 열 머리글 유지)
        h.append("<th rowspan='2' bgcolor='#1f3a5f'><font color='white'>비고 (가입할 부가서비스)</font></th>")
        h.append("</tr><tr>" + "".join(f"<th bgcolor='#eef1f5'>{short[j]}</th>" for _ in cars for j in joins) + "</tr>")
        for m in sorted(self.data, key=model_sort_key):
            row = [f"<td><b>{html.escape(m)}</b></td>"]
            best = {}
            for j in joins:
                cand = [c for c in cars if (c, j) in self.data[m]]
                if cand:
                    best[j] = max(cand, key=lambda c: self.data[m][(c, j)]["net"])
            for c in cars:
                for j in joins:
                    d = self.data[m].get((c, j))
                    if not d:
                        row.append("<td align='center'>-</td>")
                    elif best.get(j) == c:
                        row.append(f"<td align='center' bgcolor='#d8f3df'><b><font color='#1a7f37'>{d['sup'] / 10000:g}</font></b></td>")
                    else:
                        row.append(f"<td align='center'>{d['sup'] / 10000:g}</td>")
            row.append(f"<td><font size='2'>{html.escape(self.note(m, cars, owner=False))}</font></td>")
            h.append("<tr>" + "".join(row) + "</tr>")
        h.append("</table><p style='color:#777'>초록 = 그 가입유형 최고 통신사 · 공시지원금·할부원금은 매장 문의</p>")
        path = os.path.join(desktop_dir(), f"판매정책_{dstr(self.date)}_{self.cur_fee}.png")
        html_to_png("".join(h), 560 + 70 * 3 * len(cars), path)
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        info(self, f"✅ 카톡용 이미지를 만들었습니다.\n{path}\n\n카톡 단톡방에 이 사진을 올리면 됩니다.")

    def export(self, owner):
        if not self.data:
            self.calc()
            if not self.data:
                return
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
        except ImportError:
            return warn(self, "엑셀 저장 구성요소(openpyxl)가 없습니다. 설치 bat을 다시 실행해 주세요.")
        fees = sorted({to_int(self.fee.itemText(i)) for i in range(self.fee.count())}, reverse=True)
        choice, ok = QInputDialog.getItem(self, "엑셀 만들기", "어떤 요금 구간으로 만들까요?",
                                          [f"지금 구간만 ({self.cur_fee})", f"모든 구간 ({len(fees)}개 시트)"], 0, False)
        if not ok:
            return
        fee_list = fees if choice.startswith("모든") else [self.cur_fee]
        store = self.db.get("store_name", "우리매장")
        date = dstr(self.date)
        name = f"{store}_판매정책_{date}{'_사장님용' if owner else ''}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "판매정책 엑셀 저장", os.path.join(os.path.expanduser("~"), "Desktop", name),
                                              "엑셀 (*.xlsx)")
        if not path:
            return
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        thin = Side(style="thin", color="C8CED6")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        hfill = {"SKT": "E8412C", "KT": "1F1F1F", "LGU+": "C4006B"}
        cars_all = set()
        for fee in fee_list:
            data = build_store_policy(self.db, date, fee, JOIN_BY_CAT["무선"], self.flags.get(), self.margin.value(),
                                      self.unit.currentData() or 1, disc=self.disc.currentData())
            cars = sorted({c for m in data.values() for (c, _j) in m}, key=lambda c: CARRIERS.index(c) if c in CARRIERS else 9)
            cars_all |= set(cars)
            ws = wb.create_sheet(f"{fee}요금")
            cols_per = 4 if owner else 3
            ws.cell(1, 1, f"{store} 판매정책  ·  {date}  ·  요금 {fee}천원 구간  ·  금액 단위: 만원").font = Font(bold=True, size=14)
            flag_txt = ", ".join(lab for k, lab in FLAGS if self.flags.get().get(k)) or "부가·보험 조건 없음"
            ws.cell(2, 1, f"기준: {flag_txt}  /  고객지원금 = 추가지원금(공시지원금 별도)"
                          + ("  /  ※ 사장님용: 대리점·정산금 포함 — 외부 유출 금지" if owner else "")).font = Font(color="666666")
            ws.cell(3, 1, "모델").font = Font(bold=True, color="FFFFFF")
            ws.cell(3, 1).fill = PatternFill("solid", fgColor="1F3A5F")
            ws.merge_cells(start_row=3, start_column=1, end_row=4, end_column=1)
            col = 2
            for c in cars:
                span = 3 * (2 if owner else 1)
                ws.cell(3, col, c).font = Font(bold=True, color="FFFFFF")
                ws.cell(3, col).fill = PatternFill("solid", fgColor=hfill.get(c, "555555"))
                ws.cell(3, col).alignment = Alignment(horizontal="center")
                ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + span - 1)
                for j in JOIN_BY_CAT["무선"]:
                    labs = [j] if not owner else [f"{j} 지원금", f"{j} 대리점/정산"]
                    for lab in labs:
                        ws.cell(4, col, lab).font = Font(bold=True)
                        ws.cell(4, col).fill = PatternFill("solid", fgColor="EEF1F5")
                        ws.cell(4, col).alignment = Alignment(horizontal="center", wrap_text=True)
                        col += 1
            r = 5
            for m in sorted(data, key=model_sort_key):
                ws.cell(r, 1, m).font = Font(bold=True)
                col = 2
                best = {}
                for j in JOIN_BY_CAT["무선"]:
                    cand = [c for c in cars if (c, j) in data[m]]
                    if cand:
                        best[j] = max(cand, key=lambda c: data[m][(c, j)]["net"])
                for c in cars:
                    for j in JOIN_BY_CAT["무선"]:
                        d = data[m].get((c, j))
                        cell = ws.cell(r, col, round(d["sup"] / 10000, 1) if d else "-")
                        cell.alignment = Alignment(horizontal="center")
                        if d and best.get(j) == c:
                            cell.fill = PatternFill("solid", fgColor="D8F3DF")
                            cell.font = Font(bold=True, color="1A7F37")
                        col += 1
                        if owner:
                            ws.cell(r, col, f"{d['agency']} / {d['net'] / 10000:g}" if d else "-").alignment = \
                                Alignment(horizontal="center")
                            col += 1
                self_data, self.data = self.data, data
                ws.cell(r, col, self.note(m, cars, owner=owner))
                self.data = self_data
                note_col = col
                r += 1
            for row in ws.iter_rows(min_row=3, max_row=r - 1, max_col=col - 1):
                for cell in row:
                    cell.border = border
            ws.column_dimensions["A"].width = 30
            for ci in range(2, col):
                ws.column_dimensions[get_column_letter(ci)].width = 16 if owner else 9
            ws.cell(3, col, "비고 (가입할 부가서비스)").font = Font(bold=True)
            ws.column_dimensions[get_column_letter(col)].width = 60
            ws.freeze_panes = "B5"
            ws.cell(r + 1, 1, "초록 = 그 가입유형에서 가장 좋은 통신사. 공시지원금·할부원금은 매장 문의.").font = Font(color="666666")
            ws.cell(r + 3, 1, "📌 통신사별 부가서비스 안내").font = Font(bold=True, size=12)
            rr = r + 4
            for ln in addon_text_lines(self.db, date):
                ws.cell(rr, 1, ln)
                rr += 1
        wb.save(path)
        extra = ""
        d = self.db.get("share_dir", "")
        if not owner and d and os.path.isdir(d):
            shutil.copyfile(path, os.path.join(d, os.path.basename(path)))
            extra = "\n공유 폴더에도 복사했습니다."
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        info(self, f"✅ 엑셀을 만들었습니다.\n{path}{extra}\n\n" +
             ("이 파일은 대리점·정산금이 들어 있으니 직원에게 보내지 마세요." if owner
              else "이 파일을 카톡으로 직원들에게 보내면 됩니다. (리베이트·마진·대리점은 안 들어 있음)"))


# ============================================================
# 탭: 판매가 · 가격표
# ============================================================
class PriceTab(QWidget):
    HEAD = ["구분", "통신사", "모델/상품", "요금제/약정", "가입유형", "대리점", "출고가", "공시지원금", "리베이트", "차감",
            "정산예상", "고객지원금/사은품", "할부원금(공시)", "할부원금(선약)", "매장마진", "조건"]

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.rows = []
        self._init = False
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>손님에게 얼마에 팔지 계산하는 화면</b>입니다. 목표마진을 정하면 고객지원금·할부원금이 자동 계산됩니다. "
                             "<b>[가격표 만들기]</b>를 누르면 손님용 가격표(리베이트·마진 안 보임)가 인터넷 창으로 열려서 바로 인쇄할 수 있어요.<br>"
                             "※ 여기는 기본 정책 기준입니다. 부가·보험·기변 조건까지 따진 정확한 비교는 <b>🏆 최적 대리점</b> 탭에서 하세요."))
        top = QHBoxLayout()
        self.date = date_edit()
        self.cat = QComboBox()
        self.cat.addItems(["전체"] + CATEGORIES)
        self.carrier = QComboBox()
        self.agency = QComboBox()
        self.search = QLineEdit()
        self.search.setPlaceholderText("모델 검색 (예: S25)")
        self.margin = money_spin(0, 5_000_000)
        self.unit = QComboBox()
        for t, v in [("절사 없음", 1), ("천원 단위", 1000), ("만원 단위", 10000)]:
            self.unit.addItem(t, v)
        for w in [QLabel("날짜"), self.date, QLabel(" 구분"), self.cat, QLabel(" 통신사"), self.carrier,
                  QLabel(" 대리점"), self.agency, self.search, QLabel(" 목표마진"), self.margin,
                  QLabel(" 지원금"), self.unit]:
            top.addWidget(w)
        top.addWidget(btn("계산", self.calc, primary=True))
        top.addStretch()
        lay.addLayout(top)
        bar = QHBoxLayout()
        bar.addWidget(btn("🖨 손님용 가격표 만들기", self.make_html, primary=True))
        bar.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"판매가_{dstr(self.date)}.csv")))
        self.summary = QLabel()
        bar.addStretch()
        bar.addWidget(self.summary)
        lay.addLayout(bar)
        self.table = make_table(self.HEAD)
        lay.addWidget(self.table)
        lay.addWidget(hint("고객지원금 = (리베이트 - 차감) - 목표마진  ·  할부원금(공시) = 출고가 - 공시지원금 - 고객지원금  ·  "
                           "할부원금(선약) = 출고가 - 고객지원금.   대리점을 '전체'로 두면 같은 조건에서 제일 많이 주는 대리점 기준으로 계산합니다. "
                           "빨간 줄 = 목표마진이 안 나오는 조건."))
        self.search.returnPressed.connect(self.calc)
        for cb in (self.cat, self.carrier, self.agency):
            cb.currentIndexChanged.connect(lambda *_: self.calc())
        self.carrier.currentIndexChanged.connect(lambda *_: self._fill_agencies())

    def _fill_agencies(self):
        car = self.carrier.currentText()
        fill_combo(self.agency, self.db.agency_items(None if car == "전체" else car), "전체(제일 좋은 정책)")

    def refresh(self):
        fill_text_combo(self.carrier, ["전체"] + self.db.carriers())
        self._fill_agencies()
        if not self._init:
            self._init = True
            self.margin.setValue(to_int(self.db.get("target_margin", "100000")))
            i = self.unit.findData(to_int(self.db.get("round_unit", "10000")))
            self.unit.setCurrentIndex(max(0, i))
        self.calc()

    def calc(self):
        cat, car = self.cat.currentText(), self.carrier.currentText()
        target = self.margin.value()
        self.rows = compute_prices(self.db, dstr(self.date), self.agency.currentData(),
                                   None if cat == "전체" else cat, None if car == "전체" else car,
                                   self.search.text().strip().lower(), target, self.unit.currentData() or 1)
        data, colors = [], []
        for r in self.rows:
            wired = r["category"] == "유선"
            data.append([r["category"], r["carrier"], r["model"], r["plan"], r["join_type"], r["agency"],
                         "" if wired else r["release_price"], "" if wired else r["public_subsidy"],
                         r["rebate"], r["deduction"], r["net"], r["support"],
                         "" if wired else r["hal_public"], "" if wired else r["hal_select"],
                         r["margin"], r["conditions"]])
            colors.append(RED if r["margin"] < target else None)
        fill_table(self.table, data, colors=colors)
        self.summary.setText(f"{len(self.rows)}개 조건")

    def make_html(self):
        if not self.rows:
            return warn(self, "계산된 정책이 없습니다. 정책을 먼저 저장해 주세요.")
        store = self.db.get("store_name", "우리매장")
        date = dstr(self.date)
        path, _ = QFileDialog.getSaveFileName(self, "가격표 저장", os.path.join(BASE_DIR, f"가격표_{date}.html"),
                                              "HTML (*.html)")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(build_price_html(store, date, self.rows))
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        info(self, "가격표를 열었습니다.\n인쇄하려면 열린 창에서 Ctrl+P 를 누르세요.")


# ============================================================
# 탭: 대리점 비교
# ============================================================
class CompareTab(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>같은 통신사 대리점끼리 누가 더 많이 주는지</b> 한눈에 비교합니다. "
                             "초록색 칸이 그 조건에서 제일 좋은 대리점이에요. (금액 = 리베이트 - 차감)<br>"
                             "※ 기본 정책끼리 비교입니다. 손님 조건(부가·보험·기변)까지 따지려면 <b>🏆 최적 대리점</b> 탭을 쓰세요."))
        top = QHBoxLayout()
        self.date = date_edit()
        self.cat = QComboBox()
        self.cat.addItems(CATEGORIES)
        self.carrier = QComboBox()
        self.join = QComboBox()
        self.search = QLineEdit()
        self.search.setPlaceholderText("모델 검색")
        for w in [QLabel("날짜"), self.date, QLabel(" 구분"), self.cat, QLabel(" 통신사"), self.carrier,
                  QLabel(" 가입유형"), self.join, self.search]:
            top.addWidget(w)
        top.addWidget(btn("비교", self.calc, primary=True))
        top.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"대리점비교_{dstr(self.date)}.csv")))
        top.addStretch()
        lay.addLayout(top)
        self.table = make_table(["통신사"])
        lay.addWidget(self.table)
        self.search.returnPressed.connect(self.calc)
        self.cat.currentIndexChanged.connect(lambda *_: (self._joins(), self.calc()))
        self.carrier.currentIndexChanged.connect(lambda *_: self.calc())
        self.join.currentIndexChanged.connect(lambda *_: self.calc())
        self._joins()

    def _joins(self):
        self.join.blockSignals(True)
        self.join.clear()
        self.join.addItems(["전체"] + JOIN_BY_CAT[self.cat.currentText()])
        self.join.blockSignals(False)

    def refresh(self):
        fill_text_combo(self.carrier, self.db.carriers())
        self.calc()

    def calc(self):
        cat, car, jt = self.cat.currentText(), self.carrier.currentText(), self.join.currentText()
        q = self.search.text().strip().lower()
        ags = [a["name"] for a in self.db.agencies(car)]
        data = {}
        for p in self.db.visible_policies(dstr(self.date)):
            if p["category"] != cat or p["carrier"] != car:
                continue
            if jt != "전체" and p["join_type"] != jt:
                continue
            if q and q not in p["model"].lower() and q not in (p["plan"] or "").lower():
                continue
            if p["agency"] not in ags:
                ags.append(p["agency"])
            k = (p["model"], p["plan"], p["join_type"])
            data.setdefault(k, {})[p["agency"]] = p["rebate"] - p["deduction"]
        heads = ["모델/상품", "요금제/약정", "가입유형"] + ags + ["제일 좋은 곳", "1-2위 차이"]
        self.table.setSortingEnabled(False)
        self.table.setColumnCount(len(heads))
        self.table.setHorizontalHeaderLabels(heads)
        rows, cell_colors = [], {}
        keys = sorted(data, key=lambda k: (k[0], k[1], ALL_JOINS.index(k[2]) if k[2] in ALL_JOINS else 9))
        for r, k in enumerate(keys):
            vals = data[k]
            ranked = sorted(vals.items(), key=lambda x: -x[1])
            best = ranked[0][0]
            gap = ranked[0][1] - ranked[1][1] if len(ranked) > 1 else 0
            row = list(k)
            for i, a in enumerate(ags):
                row.append(NumItem(vals[a]) if a in vals else txt_item("-"))
                if a == best and len(vals) > 1:
                    cell_colors[(r, 3 + i)] = GREEN
            rows.append(row + [best, gap])
        fill_table(self.table, rows, cell_colors=cell_colors)


# ============================================================
# 탭: 정책 변경 이력
# ============================================================
class HistoryTab(QWidget):
    HEAD = ["적용일", "대리점", "구분", "통신사", "모델/상품", "요금제/약정", "가입유형", "변화", "리베이트", "변동금액",
            "차감", "공시지원금", "출고가", "조건", "등록시각"]

    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>정책이 언제 얼마나 바뀌었는지</b> 기록입니다. 변동금액 초록 = 올랐음, 빨강 = 내렸음. "
                             "예전 날짜 정책도 모두 남아 있어서 지난 개통 건 정산 확인할 때 씁니다."))
        top = QHBoxLayout()
        self.agency = QComboBox()
        self.cat = QComboBox()
        self.cat.addItems(["전체"] + CATEGORIES)
        self.search = QLineEdit()
        self.search.setPlaceholderText("모델 검색")
        self.d_from = date_edit(d=QDate.currentDate().addMonths(-3))
        self.d_to = date_edit(d=QDate.currentDate().addDays(60))
        self.only_change = QComboBox()
        self.only_change.addItems(["전부 보기", "금액 바뀐 것만"])
        for w in [QLabel("대리점"), self.agency, QLabel(" 구분"), self.cat, self.search,
                  QLabel(" 기간"), self.d_from, QLabel("~"), self.d_to, self.only_change]:
            top.addWidget(w)
        top.addWidget(btn("조회", self.calc, primary=True))
        top.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, "정책이력.csv")))
        top.addWidget(btn("🗂 정책 원본 보관함", lambda: ArchiveDialog(self, self.db).exec()))
        top.addStretch()
        lay.addLayout(top)
        self.table = make_table(self.HEAD)
        lay.addWidget(self.table)
        self.search.returnPressed.connect(self.calc)
        for cb in (self.agency, self.cat, self.only_change):
            cb.currentIndexChanged.connect(lambda *_: self.calc())

    def refresh(self):
        fill_combo(self.agency, self.db.agency_items(), "전체")
        self.calc()

    def calc(self):
        aid, cat = self.agency.currentData(), self.cat.currentText()
        q = self.search.text().strip().lower()
        f, t = dstr(self.d_from), dstr(self.d_to)
        only = self.only_change.currentIndex() == 1
        rows = self.db.q("SELECT p.*, a.name AS agency FROM policies p JOIN agencies a ON a.id=p.agency_id "
                         "ORDER BY p.agency_id, p.category, p.carrier, p.model, p.plan, p.join_type, p.valid_from, p.id")
        prev, out = {}, []
        for p in rows:
            k = (p["agency_id"], p["category"], p["carrier"], p["model"], p["plan"], p["join_type"])
            pv = prev.get(k)
            prev[k] = p
            if p["deleted"]:
                kind, change = "종료", None
            elif pv is None or pv["deleted"]:
                kind, change = "신규", None
            else:
                kind, change = "변경", p["rebate"] - pv["rebate"]
            if not (f <= p["valid_from"] <= t):
                continue
            if aid and p["agency_id"] != aid:
                continue
            if cat != "전체" and p["category"] != cat:
                continue
            if q and q not in p["model"].lower() and q not in (p["plan"] or "").lower():
                continue
            if only and not change:
                continue
            out.append((p, kind, change))
        out.sort(key=lambda x: (x[0]["valid_from"], x[0]["id"]), reverse=True)
        data, cc = [], {}
        for r, (p, kind, change) in enumerate(out):
            ch = NumItem(change, signed(change)) if change is not None else txt_item("")
            if change:
                cc[(r, 9)] = GREEN if change > 0 else RED
            cc[(r, 7)] = {"종료": RED, "신규": BLUE}.get(kind, YELLOW)
            end = kind == "종료"
            data.append([p["valid_from"], p["agency"], p["category"], p["carrier"], p["model"], p["plan"],
                         p["join_type"], kind, "" if end else p["rebate"], ch, "" if end else p["deduction"],
                         "" if end else p["public_subsidy"], "" if end else p["release_price"],
                         p["conditions"] or "", p["created_at"] or ""])
        fill_table(self.table, data, cell_colors=cc)


# ============================================================
# 손님 조건 체크박스 묶음
# ============================================================
class FlagBox(QWidget):
    def __init__(self, on_change=None, cols=5):
        super().__init__()
        g = QGridLayout(self)
        g.setContentsMargins(0, 0, 0, 0)
        self.boxes = {}
        for i, (k, lab) in enumerate(FLAGS):
            cb = QCheckBox(lab)
            self.boxes[k] = cb
            g.addWidget(cb, i // cols, i % cols)
        self.boxes["초단기기변"].toggled.connect(lambda on: on and self.boxes["단기기변"].setChecked(True))
        self.boxes["단기기변"].toggled.connect(lambda on: (not on) and self.boxes["초단기기변"].setChecked(False))
        if on_change:
            for cb in self.boxes.values():
                cb.toggled.connect(lambda *_: on_change())

    def get(self):
        return {k: cb.isChecked() for k, cb in self.boxes.items()}

    def text(self):
        return ",".join(k for k, v in self.get().items() if v)

    def set(self, flags):
        if isinstance(flags, str):
            flags = {k: True for k in _split(flags)}
        for k, cb in self.boxes.items():
            cb.blockSignals(True)
            cb.setChecked(bool(flags.get(k)))
            cb.blockSignals(False)


# ============================================================
# 탭: 최적 대리점 찾기 (부가·차감 조건까지 계산)
# ============================================================
class BestTab(QWidget):
    HEAD = ["순위", "대리점", "통신사", "모델/상품", "요금제 구간", "가입유형", "기본 리베이트", "고정 차감", "조건 가감",
            "실제 받는 돈", "1등과 차이", "공시지원금", "출고가", "고객지원금(목표마진)", "할부원금(공시)", "적용된 조건", "⚠ 주의"]

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.on_register = None
        self.results = []
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>손님 조건을 넣으면 대리점 전체의 부가서비스·보험·요금제·기변 차감까지 계산해서 "
                             "실제로 제일 많이 받는 곳</b>을 찾아줍니다. ① 모델·요금제·가입유형 고르기 → ② 손님이 가입할 부가/보험 체크 → "
                             "③ [🏆 찾기]. 마음에 들면 [이 조건으로 개통 등록]을 누르세요."))
        r1 = QHBoxLayout()
        self.date = date_edit()
        self.cat = QComboBox()
        self.cat.addItems(CATEGORIES)
        self.carrier = QComboBox()
        self.join = QComboBox()
        self.fee = QComboBox()
        self.fee.setEditable(True)
        self.fee.addItems(["상관없음", "115", "105", "95", "85", "75", "69", "61", "55", "44", "33"])
        self.fee.setCurrentText("115")
        self.fee.setToolTip("손님 요금제 월정액(천원). 예: 월 69,000원 요금제면 69")
        self.model = QComboBox()
        self.model.setEditable(True)
        self.model.setMinimumWidth(260)
        self.model.lineEdit().setPlaceholderText("모델 검색 (예: S26 256)")
        for w in [QLabel("날짜"), self.date, QLabel(" 구분"), self.cat, QLabel(" 통신사"), self.carrier,
                  QLabel(" 모델"), self.model, QLabel(" 요금제(천원)"), self.fee, QLabel(" 가입유형"), self.join]:
            r1.addWidget(w)
        self.disc = discount_combo()
        for w in [QLabel(" 할인"), self.disc]:
            r1.addWidget(w)
        r1.addStretch()
        lay.addLayout(r1)
        fb = QGroupBox("손님 조건 (해당하는 것만 체크)")
        fl = QVBoxLayout(fb)
        self.flags = FlagBox(on_change=lambda: self.results and self.calc())
        fl.addWidget(self.flags)
        lay.addWidget(fb)
        r2 = QHBoxLayout()
        r2.addWidget(btn("🏆 제일 좋은 곳 찾기", self.calc, primary=True, big=True))
        r2.addWidget(btn("📱 선택한 줄(없으면 1등)로 개통 등록", self.register))
        r2.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"최적대리점_{dstr(self.date)}.csv")))
        r2.addStretch()
        lay.addLayout(r2)
        self.summary = QLabel("")
        self.summary.setStyleSheet("font-size:13pt;font-weight:bold;color:#1a7f37;padding:4px")
        self.summary.setWordWrap(True)
        lay.addWidget(self.summary)
        self.table = make_table(self.HEAD)
        lay.addWidget(self.table, 1)
        lay.addWidget(hint("실제 받는 돈 = 기본 리베이트 - 고정 차감 + 조건 가감(부가·보험·요금제·기변 규칙). "
                           "규칙은 '📋 정책 입력 → 📐 부가·차감 규칙'에 대리점별로 저장된 것을 씁니다. "
                           "⚠ 주의 칸은 개통 후 환수될 수 있는 조건이라 금액에는 안 들어갑니다."))
        self.cat.currentIndexChanged.connect(lambda *_: self._on_cat())
        self.date.dateChanged.connect(lambda *_: self._fill_models())
        self.carrier.currentIndexChanged.connect(lambda *_: self._fill_models())
        self.model.lineEdit().returnPressed.connect(self.calc)
        self._on_cat(first=True)

    def _on_cat(self, first=False):
        cat = self.cat.currentText()
        cur = self.join.currentText()
        self.join.clear()
        self.join.addItems(JOIN_BY_CAT[cat])
        i = self.join.findText(cur)
        self.join.setCurrentIndex(i if i >= 0 else (1 if cat == "무선" else 0))
        if not first:
            self._fill_models()

    def refresh(self):
        fill_text_combo(self.carrier, ["전체"] + self.db.carriers())
        self._fill_models()

    def _fill_models(self):
        cat = self.cat.currentText()
        car = self.carrier.currentText()
        names = {}
        for p in self.db.visible_policies(dstr(self.date)):
            if p["category"] != cat or (car not in ("", "전체") and p["carrier"] != car):
                continue
            k = norm_model(p["model"])
            if k and (k not in names or len(p["model"]) < len(names[k])):
                names[k] = p["model"]
        cur = self.model.currentText()
        self.model.blockSignals(True)
        self.model.clear()
        for k, n in sorted(names.items(), key=lambda x: x[1]):
            self.model.addItem(n, k)
        self.model.setCurrentIndex(-1)
        self.model.setEditText(cur)
        self.model.blockSignals(False)

    def calc(self):
        date = dstr(self.date)
        cat, car, join = self.cat.currentText(), self.carrier.currentText(), self.join.currentText()
        ft = self.fee.currentText().strip()
        fee = None if (not ft or ft == "상관없음") else to_int(ft)
        typed = self.model.currentText().strip()
        idx = self.model.findText(typed)
        exact_key = self.model.itemData(idx) if idx >= 0 else None
        q = norm_model(typed)
        if not q:
            self.summary.setText("")
            return warn(self, "모델을 고르거나 입력하세요. (예: S26 256)")
        flags = self.flags.get()
        target = to_int(self.db.get("target_margin", "100000"))
        unit = to_int(self.db.get("round_unit", "10000")) or 1
        groups = {}
        for p in self.db.visible_policies(date):
            if p["category"] != cat or p["join_type"] != join:
                continue
            if car not in ("", "전체") and p["carrier"] != car:
                continue
            mk = norm_model(p["model"])
            if (exact_key and mk != exact_key) or (not exact_key and q not in mk):
                continue
            groups.setdefault((p["agency_id"], mk), []).append(p)
        rules_cache, res = {}, []
        for (aid, mk), cands in groups.items():
            if aid not in rules_cache:
                rules_cache[aid] = self.db.effective_rules(date, aid)
            for p in pick_by_fee(by_discount(cands, self.disc.currentData()), fee):
                use_fee = fee if fee is not None else plan_fee(p["plan"])
                adj, applied, warns = apply_rules(rules_cache[aid], p, join, use_fee, flags)
                net = p["rebate"] - p["deduction"] + adj
                sup = max(0, net - target)
                if unit > 1:
                    sup = sup // unit * unit
                res.append(dict(p=p, mk=mk, adj=adj, net=net, applied=applied, warns=warns, sup=sup,
                                hal=max(0, p["release_price"] - p["public_subsidy"] - sup)))
        if not res:
            self.results = []
            fill_table(self.table, [])
            self.summary.setText("")
            return warn(self, "조건에 맞는 정책이 없습니다.\n모델 이름, 가입유형, 요금제 금액대를 확인하거나 '상관없음'으로 바꿔보세요.")
        by_model = {}
        for r in res:
            by_model.setdefault(r["mk"], []).append(r)
        order = sorted(by_model.values(), key=lambda L: -max(x["net"] for x in L))
        self.results, data, colors = [], [], []
        for L in order:
            L.sort(key=lambda x: -x["net"])
            top = L[0]["net"]
            for i, r in enumerate(L, 1):
                p = r["p"]
                self.results.append(r)
                data.append([f"{i}등", p["agency"], p["carrier"], p["model"], p["plan"], p["join_type"], p["rebate"],
                             p["deduction"], NumItem(r["adj"], signed(r["adj"])), r["net"],
                             NumItem(r["net"] - top, signed(r["net"] - top)), p["public_subsidy"], p["release_price"],
                             r["sup"], r["hal"], " / ".join(r["applied"]) or "-", " / ".join(r["warns"])])
                colors.append(GREEN if i == 1 else (RED if r["net"] < 0 else None))
        self.table.setProperty("_user_sorted", False)
        self.table._user_sorted = False
        fill_table(self.table, data, ids=list(range(len(self.results))), colors=colors)
        best = order[0]
        msg = f"🏆 1등: {best[0]['p']['agency']} — 실제 받는 돈 {won(best[0]['net'])}원"
        if len(best) > 1:
            msg += f"  (2등 {best[1]['p']['agency']}보다 {won(best[0]['net'] - best[1]['net'])}원 더)"
        per_car = {}
        for r in best:
            c = r["p"]["carrier"]
            if c not in per_car:
                per_car[c] = r
        msg += "\n📡 통신사별 최고:  " + "   |   ".join(
            f"{c} {per_car[c]['p']['agency']} {per_car[c]['net'] / 10000:g}만"
            for c in sorted(per_car, key=lambda c: CARRIERS.index(c) if c in CARRIERS else 9))
        if len(order) > 1:
            msg += f"\n※ 검색에 모델 {len(order)}개가 걸렸습니다. 모델별로 순위를 따로 매겼어요."
        self.summary.setText(msg)

    def register(self):
        if not self.results:
            return warn(self, "먼저 [🏆 제일 좋은 곳 찾기]를 누르세요.")
        rows = selected_rows(self.table)
        r = self.results[row_id(self.table, rows[0])] if rows else self.results[0]
        p = r["p"]
        if self.on_register:
            self.on_register(dict(date=dstr(self.date), agency_id=p["agency_id"], category=p["category"],
                                  model=p["model"], plan=p["plan"], join=p["join_type"], flags=self.flags.get()))


# ============================================================
# 탭: 요금제 구간별 비교 (통신사·대리점 × 요금 구간)
# ============================================================
DEFAULT_FEES = "115,105,95,85,75,69,61,55,44,33"


class TierTab(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.shown = False
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>한 모델을 요금제 구간별로</b> 대리점·통신사끼리 비교합니다. 통신사마다 요금제 이름이 달라도 "
                             "<b>월 요금(천원) 기준</b>으로 맞춰서 비교해요. 칸 = 부가·차감 조건까지 반영한 <b>실제 받는 돈</b>, "
                             "초록 = 그 요금 구간 1등. 칸에 마우스를 올리면 어떤 요금제 구간·조건이 적용됐는지 보입니다."))
        r1 = QHBoxLayout()
        self.date = date_edit()
        self.carrier = QComboBox()
        self.model = QComboBox()
        self.model.setEditable(True)
        self.model.setMinimumWidth(260)
        self.model.lineEdit().setPlaceholderText("모델 선택 (예: S26 256)")
        self.join = QComboBox()
        self.join.addItems(JOIN_BY_CAT["무선"])
        self.join.setCurrentText("번호이동")
        self.fees = QLineEdit("자동")
        self.fees.setToolTip("'자동' = SK·KT·LG 정책표에 있는 요금 구간을 모아서 자동으로 칸을 만듦.\n"
                             "직접 정하려면 월 요금(천원)을 쉼표로. 예: 115,95,69,55")
        self.fees.setMinimumWidth(230)
        for w in [QLabel("날짜"), self.date, QLabel(" 통신사"), self.carrier, QLabel(" 모델"), self.model,
                  QLabel(" 가입유형"), self.join, QLabel(" 요금(천원)"), self.fees]:
            r1.addWidget(w)
        self.disc = discount_combo()
        for w in [QLabel(" 할인"), self.disc]:
            r1.addWidget(w)
        r1.addStretch()
        lay.addLayout(r1)
        fb = QGroupBox("손님 조건 (해당하는 것만 체크)")
        fl = QVBoxLayout(fb)
        self.flags = FlagBox(on_change=lambda: self.shown and self.calc())
        fl.addWidget(self.flags)
        lay.addWidget(fb)
        r2 = QHBoxLayout()
        r2.addWidget(btn("📶 구간별 비교하기", self.calc, primary=True, big=True))
        r2.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"구간비교_{dstr(self.date)}.csv")))
        r2.addStretch()
        lay.addLayout(r2)
        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        self.summary.setStyleSheet("font-size:11pt;font-weight:bold;color:#1a7f37;padding:4px")
        lay.addWidget(self.summary)
        self.table = make_table(["대리점", "통신사"])
        lay.addWidget(self.table, 1)
        lay.addWidget(hint("요금 구간 맞추기: 손님 요금이 정확히 있는 구간이 있으면 그 구간, 없으면 그보다 낮은 구간 중 가장 높은 구간 기준. "
                           "'-' = 그 대리점에 이 모델·가입유형 정책이 없음. 아래 굵은 줄 = 통신사별 최고 금액."))
        self.carrier.currentIndexChanged.connect(lambda *_: self._fill_models())
        self.date.dateChanged.connect(lambda *_: self._fill_models())
        self.join.currentIndexChanged.connect(lambda *_: self.shown and self.calc())
        self.model.lineEdit().returnPressed.connect(self.calc)

    def refresh(self):
        fill_text_combo(self.carrier, ["전체"] + self.db.carriers())
        self._fill_models()

    def _fill_models(self):
        car = self.carrier.currentText()
        names = {}
        for p in self.db.visible_policies(dstr(self.date)):
            if p["category"] != "무선" or (car not in ("", "전체") and p["carrier"] != car):
                continue
            if not fee_numbers(p["plan"]):
                continue
            k = norm_model(p["model"])
            if k and (k not in names or len(p["model"]) < len(names[k])):
                names[k] = p["model"]
        cur = self.model.currentText()
        self.model.blockSignals(True)
        self.model.clear()
        for k, n in sorted(names.items(), key=lambda x: x[1]):
            self.model.addItem(n, k)
        self.model.setCurrentIndex(-1)
        self.model.setEditText(cur)
        self.model.blockSignals(False)

    def calc(self):
        date, car, join = dstr(self.date), self.carrier.currentText(), self.join.currentText()
        fees = [int(x) for x in re.findall(r"\d+", self.fees.text())]
        typed = self.model.currentText().strip()
        idx = self.model.findText(typed)
        key = self.model.itemData(idx) if idx >= 0 else None
        q = norm_model(typed)
        if not q:
            return warn(self, "모델을 고르세요.")
        pols = [p for p in self.db.visible_policies(date)
                if p["category"] == "무선" and p["join_type"] == join
                and (car in ("", "전체") or p["carrier"] == car) and fee_numbers(p["plan"])]
        if not key:
            keys = {norm_model(p["model"]) for p in pols if q in norm_model(p["model"])}
            if len(keys) > 1:
                names = sorted({p["model"] for p in pols if norm_model(p["model"]) in keys})[:8]
                return warn(self, "검색에 모델이 여러 개 걸렸습니다. 목록에서 하나를 골라주세요.\n\n" + "\n".join(names))
            key = next(iter(keys), None)
        pols = [p for p in pols if norm_model(p["model"]) == key]
        if not fees:     # 자동: 모든 통신사 정책표의 요금 구간을 모아서 칸 만들기 (SK·KT·LG 이름이 달라도 월 요금으로)
            fees = sorted({min(fee_numbers(p["plan"])) for p in pols if fee_numbers(p["plan"])}, reverse=True)[:14] \
                or [int(x) for x in DEFAULT_FEES.split(",")]
        if not pols:
            fill_table(self.table, [])
            self.summary.setText("")
            return warn(self, "이 모델·가입유형의 정책이 없습니다.")
        flags = self.flags.get()
        by_ag, rules = {}, {}
        for p in pols:
            by_ag.setdefault((p["agency_id"], p["agency"], p["carrier"]), []).append(p)
        order = sorted(by_ag, key=lambda k: (CARRIERS.index(k[2]) if k[2] in CARRIERS else 9, k[1]))
        grid, tips = [], []
        for (aid, name, carrier) in order:
            if aid not in rules:
                rules[aid] = self.db.effective_rules(date, aid)
            row, tip, first_best = [], [], None
            for fee in fees:
                best = None
                for p in pick_by_fee(by_discount(by_ag[(aid, name, carrier)], self.disc.currentData()), fee):
                    if not fee_numbers(p["plan"]):
                        continue
                    adj, applied, warns = apply_rules(rules[aid], p, join, fee, flags)
                    net = p["rebate"] - p["deduction"] + adj
                    if best is None or net > best[0]:
                        best = (net, p["plan"], applied, warns)
                row.append(best[0] if best else None)
                if best and first_best is None:
                    first_best = best
                tip.append(f"{best[1]}\n" + ("\n".join(best[2]) if best[2] else "조건 가감 없음") +
                           (f"\n⚠ {' / '.join(best[3])}" if best[3] else "") if best else "정책 없음")
            note = "-"
            if first_best:
                bits = [re.sub(r"\s*\[.*?\]", "", first_best[1])]
                if first_best[2]:
                    bits.append("적용: " + ", ".join(short_name(x, 18) for x in first_best[2][:3]))
                ads = [short_name(clean_addon_name(r["name"]), 14) for r in rules[aid] if r["kind"] in ADDON_MAIN]
                if ads:
                    bits.append("부가: " + ", ".join(list(dict.fromkeys(ads))[:3]))
                if first_best[3]:
                    bits.append(f"⚠ 주의 {len(first_best[3])}건")
                note = " · ".join(bits)
            grid.append((name, carrier, row, note))
            tips.append(tip)
        # 통신사별 최고 줄
        carriers = sorted({c for _, c, _, _ in grid}, key=lambda c: CARRIERS.index(c) if c in CARRIERS else 9)
        top_rows = []
        for c in carriers:
            vals, who = [], []
            for j in range(len(fees)):
                xs = [(r[j], n) for n, cc, r, _ in grid if cc == c and r[j] is not None]
                vals.append(max(xs)[0] if xs else None)
                if xs:
                    who.append(max(xs)[1])
            wn = list(dict.fromkeys(who))
            top_rows.append((f"【{c} 최고】", c, vals, ("모든 구간 " + wn[0]) if len(wn) == 1 else " / ".join(wn[:3])))
        heads = ["대리점", "통신사"] + [f"{f}요금" for f in fees] + ["비고 (적용 요금제 · 조건 · 부가서비스)"]
        self.table.setSortingEnabled(False)
        self.table.setColumnCount(len(heads))
        self.table.setHorizontalHeaderLabels(heads)
        data, cc = [], {}
        col_max = [max([r[j] for _, _, r, _ in grid if r[j] is not None] or [None], key=lambda v: -10**12 if v is None else v)
                   for j in range(len(fees))]
        for i, (name, carrier, row, note) in enumerate(grid + top_rows):
            line = [name, carrier]
            for j, v in enumerate(row):
                if v is None:
                    it = txt_item("-")
                else:
                    it = NumItem(v, f"{v / 10000:g}만")
                    if i < len(grid):
                        it.setToolTip(tips[i][j])
                    if v == col_max[j] and len(grid) > 1:
                        cc[(i, 2 + j)] = GREEN
                    elif v < 0:
                        cc[(i, 2 + j)] = RED
                line.append(it)
            line.append(note)
            data.append(line)
        fill_table(self.table, data, cell_colors=cc, bold_rows=set(range(len(grid), len(data))))
        self.shown = True
        model_name = min((p["model"] for p in pols), key=len)
        parts = []
        for j, f in enumerate(fees[:5]):
            if col_max[j] is None:
                continue
            winners = [n for n, _, r, _ in grid if r[j] == col_max[j]]
            parts.append(f"{f}요금 → {winners[0]} {col_max[j] / 10000:g}만")
        self.summary.setText(f"📶 {model_name} · {join}   |   " + "   ·   ".join(parts))


# ============================================================
# 탭: 개통 등록
# ============================================================
class SalesTab(QWidget):
    HEAD = ["ID", "개통일", "담당", "대리점", "구분", "통신사", "모델/상품", "요금제", "가입", "할인", "고객", "뒷4",
            "리베이트", "차감", "조건가감", "고객지원금", "기타수입", "마진", "입금", "손님조건", "메모"]

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.editing_id = None
        self.busy = False
        self.lock = False
        self._pols = []
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>개통할 때마다 한 건씩 등록</b>하세요. 대리점·모델·요금제·가입유형을 고르면 그날 정책이 자동으로 채워지고 "
                             "마진이 바로 계산됩니다. 아래 목록을 더블클릭하면 수정할 수 있어요."))

        box = QGroupBox("개통 등록")
        g = QGridLayout(box)
        self.f_date = date_edit()
        self.f_staff = QComboBox()
        self.f_agency = QComboBox()
        self.f_agency.setMinimumWidth(200)
        self.f_cat = QComboBox()
        self.f_cat.addItems(CATEGORIES)
        self.f_carrier = QLabel("-")
        self.f_carrier.setStyleSheet("font-weight:bold")
        self.f_model = QComboBox()
        self.f_model.setEditable(True)
        self.f_model.setMinimumWidth(180)
        self.f_plan = QComboBox()
        self.f_plan.setEditable(True)
        self.f_plan.setMinimumWidth(150)
        self.f_join = QComboBox()
        self.f_join.addItems(JOIN_BY_CAT["무선"])
        self.f_disc = QComboBox()
        self.f_disc.addItems(DISCOUNT_TYPES)
        self.f_customer = QLineEdit()
        self.f_phone4 = QLineEdit()
        self.f_phone4.setMaxLength(4)
        self.f_phone4.setPlaceholderText("1234")
        self.f_memo = QLineEdit()
        self.f_phone = QLineEdit()
        self.f_phone.setPlaceholderText("010-0000-0000 (알림 문자용, 선택)")
        self.f_months = QComboBox()
        for t, v in (("24개월", 24), ("36개월", 36), ("12개월", 12), ("30개월", 30), ("일시불", 0)):
            self.f_months.addItem(t, v)
        self.f_release = money_spin()
        self.f_public = money_spin()
        self.f_rebate = money_spin()
        self.f_deduct = money_spin()
        self.f_support = money_spin()
        self.f_extra = money_spin()
        self.f_adjust = money_spin()
        self.f_adjust.setToolTip("대리점 부가·차감 규칙으로 자동 계산됩니다 (직접 고쳐도 됨)")
        self.f_flags = FlagBox(on_change=self.try_load, cols=5)
        self.lbl_margin = QLabel()
        self.lbl_margin.setStyleSheet("font-size:13pt;font-weight:bold")
        self.lbl_policy = QLabel()
        self.lbl_policy.setStyleSheet("color:#667")

        grid = [
            [("개통일", self.f_date), ("담당 직원", self.f_staff), ("대리점", self.f_agency), ("통신사", self.f_carrier)],
            [("구분", self.f_cat), ("모델/상품", self.f_model), ("요금제/약정", self.f_plan), ("가입유형", self.f_join)],
            [("할인유형", self.f_disc), ("고객명", self.f_customer), ("휴대폰 뒷4자리", self.f_phone4), ("메모", self.f_memo)],
            [("출고가", self.f_release), ("공시지원금", self.f_public), ("리베이트", self.f_rebate), ("차감", self.f_deduct)],
            [("부가·차감 가감", self.f_adjust), ("고객지원금/사은품", self.f_support),
             ("기타수입(중고 등)", self.f_extra), ("마진", self.lbl_margin)],
            [("고객 연락처", self.f_phone), ("할부", self.f_months)],
        ]
        for r, row in enumerate(grid):
            for c, (label, w) in enumerate(row):
                g.addWidget(QLabel(label), r, c * 2)
                g.addWidget(w, r, c * 2 + 1, 1, 1)
        g.addWidget(QLabel("손님 조건"), 6, 0)
        g.addWidget(self.f_flags, 6, 1, 1, 7)
        self.lbl_policy.setWordWrap(True)
        g.addWidget(self.lbl_policy, 7, 0, 1, 8)
        bar = QHBoxLayout()
        bar.addWidget(btn("정책 다시 불러오기", lambda: self.load_policy(False)))
        bar.addWidget(btn("지원금 = 목표마진 기준으로", self.auto_support))
        bar.addStretch()
        bar.addWidget(btn("새로 입력", self.reset_form))
        self.b_save = btn("등록", self.save, primary=True, big=True)
        bar.addWidget(self.b_save)
        g.addLayout(bar, 8, 0, 1, 8)
        lay.addWidget(box)

        lst = QHBoxLayout()
        self.month = date_edit("yyyy-MM")
        lst.addWidget(QLabel("조회 월"))
        lst.addWidget(self.month)
        lst.addWidget(btn("조회", self.load_list))
        lst.addWidget(btn("선택 건 수정", self.edit_selected))
        lst.addWidget(btn("선택 건 삭제", self.delete_selected))
        lst.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"개통내역_{mstr(self.month)}.csv")))
        self.list_sum = QLabel()
        lst.addStretch()
        lst.addWidget(self.list_sum)
        lay.addLayout(lst)
        self.table = make_table(self.HEAD)
        self.table.cellDoubleClicked.connect(lambda r, _c: self.load_sale(row_id(self.table, r)))
        lay.addWidget(self.table, 1)

        self.f_date.dateChanged.connect(lambda *_: self.refresh_models())
        self.f_agency.currentIndexChanged.connect(lambda *_: self.refresh_models())
        self.f_cat.currentIndexChanged.connect(lambda *_: self.on_cat())
        self.f_model.currentTextChanged.connect(lambda *_: self.refresh_plans())
        self.f_plan.currentTextChanged.connect(lambda *_: self.try_load())
        self.f_join.currentIndexChanged.connect(lambda *_: self.try_load())
        self.f_disc.currentIndexChanged.connect(lambda *_: self.try_load())
        for s in (self.f_rebate, self.f_deduct, self.f_support, self.f_extra, self.f_release, self.f_public,
                  self.f_adjust):
            s.valueChanged.connect(lambda *_: self.update_margin())
        self.month.dateChanged.connect(lambda *_: self.load_list())
        self.update_margin()

    def refresh(self):
        fill_combo(self.f_staff, [(s["id"], s["name"]) for s in self.db.staff(active_only=True)], "(미지정)")
        fill_combo(self.f_agency, self.db.agency_items())
        self.refresh_models()
        self.load_list()

    def on_cat(self):
        cat = self.f_cat.currentText()
        self.f_join.blockSignals(True)
        cur = self.f_join.currentText()
        self.f_join.clear()
        self.f_join.addItems(JOIN_BY_CAT[cat])
        i = self.f_join.findText(cur)
        self.f_join.setCurrentIndex(max(0, i))
        self.f_join.blockSignals(False)
        wired = cat == "유선"
        for w in (self.f_release, self.f_public, self.f_disc):
            w.setEnabled(not wired)
        if wired:
            self.f_release.setValue(0)
            self.f_public.setValue(0)
        self.refresh_models()

    def refresh_models(self):
        if self.busy:
            return
        self.busy = True
        aid = self.f_agency.currentData()
        self.f_carrier.setText(self.db.agency_carrier(aid) or "-")
        cur = self.f_model.currentText()
        pols = self.db.effective_policies(dstr(self.f_date), aid) if aid else []
        self._pols = [p for p in pols if p["category"] == self.f_cat.currentText()]
        self.f_model.clear()
        self.f_model.addItems(sorted({p["model"] for p in self._pols}))
        self.f_model.setCurrentText(cur)
        self.busy = False
        self.refresh_plans()

    def refresh_plans(self):
        if self.busy:
            return
        self.busy = True
        cur = self.f_plan.currentText()
        m = self.f_model.currentText()
        plans = sorted({p["plan"] for p in self._pols if p["model"] == m})
        self.f_plan.clear()
        self.f_plan.addItems(plans)
        if cur in plans or not plans:
            self.f_plan.setCurrentText(cur)
        self.busy = False
        self.try_load()

    def try_load(self):
        if self.busy or self.lock:
            return
        self.load_policy(silent=True)

    def _carrier(self):
        c = self.f_carrier.text()
        return "" if c == "-" else c

    def load_policy(self, silent=False):
        aid = self.f_agency.currentData()
        if not aid:
            if not silent:
                warn(self, "대리점을 선택하세요.")
            return
        p = self.db.find_policy(dstr(self.f_date), aid, self.f_cat.currentText(), self._carrier(),
                                self.f_model.currentText().strip(), self.f_plan.currentText().strip(),
                                self.f_join.currentText())
        if not p:
            self.lbl_policy.setText("⚠ 이 조건의 정책이 없습니다. 금액을 직접 넣어도 됩니다.")
            if not silent:
                warn(self, "이 날짜·대리점·조건의 정책이 없습니다.")
            return
        wired = self.f_cat.currentText() == "유선"
        self.f_release.setValue(0 if wired else p["release_price"])
        self.f_public.setValue(0 if (wired or self.f_disc.currentText() == "선택약정") else p["public_subsidy"])
        self.f_rebate.setValue(p["rebate"])
        self.f_deduct.setValue(p["deduction"])
        adj, applied, warns = apply_rules(self.db.effective_rules(dstr(self.f_date), aid), p, self.f_join.currentText(),
                                          plan_fee(p["plan"]), self.f_flags.get())
        self.f_adjust.setValue(adj)
        self.auto_support()
        cond = f" · 메모: {p['conditions']}" if p["conditions"] else ""
        rule_txt = f"\n   조건 가감: {' / '.join(applied)}" if applied else ""
        warn_txt = f"\n   ⚠ 주의: {' / '.join(warns)}" if warns else ""
        self.lbl_policy.setText(f"✔ 정책 자동 적용 ({p['valid_from']} 정책){cond}{rule_txt}{warn_txt}")

    def auto_support(self):
        target = to_int(self.db.get("target_margin", "100000"))
        unit = to_int(self.db.get("round_unit", "10000")) or 1
        s = max(0, self.f_rebate.value() - self.f_deduct.value() + self.f_adjust.value() - target)
        self.f_support.setValue(s // unit * unit if unit > 1 else s)

    def calc_margin(self):
        return (self.f_rebate.value() - self.f_deduct.value() + self.f_adjust.value()
                - self.f_support.value() + self.f_extra.value())

    def update_margin(self):
        m = self.calc_margin()
        color = "#1a7f37" if m >= 0 else "#d62828"
        extra = ""
        if self.f_cat.currentText() == "무선":
            hal = max(0, self.f_release.value() - self.f_public.value() - self.f_support.value())
            extra = f"  <span style='font-size:10pt;color:#555'>(손님 할부원금 {won(hal)} 원)</span>"
        target = to_int(self.db.get("target_margin", "100000"))
        if m < target:
            extra += f"  <span style='color:#d62828;font-size:10pt'>⚠ 목표마진보다 {won(target - m)}원 적음</span>"
        self.lbl_margin.setText(f"<span style='color:{color}'>{won(m)} 원</span>{extra}")

    def reset_form(self):
        self.editing_id = None
        self.b_save.setText("등록")
        for w in (self.f_customer, self.f_phone4, self.f_memo):
            w.clear()
        self.f_extra.setValue(0)
        self.f_phone.clear()
        self.f_flags.set({})
        self.lbl_policy.setText("")
        self.try_load()

    def save(self):
        aid = self.f_agency.currentData()
        model = self.f_model.currentText().strip()
        if not aid or not model:
            return warn(self, "대리점과 모델/상품은 꼭 넣어주세요.")
        wired = self.f_cat.currentText() == "유선"
        target = to_int(self.db.get("target_margin", "100000"))
        if self.calc_margin() < target and not ask(
                self, f"⚠ 이 개통은 마진이 {won(self.calc_margin())}원으로 목표마진({won(target)}원)보다 적습니다.\n그래도 등록할까요?"):
            return
        vals = dict(
            sale_date=dstr(self.f_date), staff_id=self.f_staff.currentData(), agency_id=aid,
            category=self.f_cat.currentText(), carrier=self._carrier(), model=model,
            plan=self.f_plan.currentText().strip(), join_type=self.f_join.currentText(),
            discount_type="" if wired else self.f_disc.currentText(),
            customer=self.f_customer.text().strip(), phone4=last4(self.f_phone4.text()),
            release_price=self.f_release.value(), public_subsidy=self.f_public.value(),
            rebate=self.f_rebate.value(), deduction=self.f_deduct.value(), support=self.f_support.value(),
            extra_income=self.f_extra.value(), margin=self.calc_margin(), memo=self.f_memo.text().strip(),
            adjust=self.f_adjust.value(), flags=self.f_flags.text(),
            phone=self.f_phone.text().strip(), months=self.f_months.currentData())
        if not vals["phone4"] and vals["phone"]:
            vals["phone4"] = last4(vals["phone"])
        if self.editing_id:
            sets = ",".join(f"{k}=?" for k in vals)
            self.db.x(f"UPDATE sales SET {sets} WHERE id=?", (*vals.values(), self.editing_id))
            sid = self.editing_id
        else:
            vals["created_at"] = now()
            sid = self.db.x(f"INSERT INTO sales({','.join(vals)}) VALUES({','.join('?' * len(vals))})",
                            tuple(vals.values()))
        try:
            sync_terms(self.db, sid)
        except Exception:
            pass
        self.month.setDate(self.f_date.date())
        self.reset_form()
        self.load_list()
        self.db.notify("sales")

    def load_sale(self, sid):
        if not sid:
            return
        s = self.db.one("SELECT * FROM sales WHERE id=?", (sid,))
        if not s:
            return
        self.lock = True
        try:
            self.f_date.setDate(QDate.fromString(s["sale_date"], "yyyy-MM-dd"))
            self.f_staff.setCurrentIndex(max(0, self.f_staff.findData(s["staff_id"])))
            i = self.f_agency.findData(s["agency_id"])
            if i >= 0:
                self.f_agency.setCurrentIndex(i)
            self.f_cat.setCurrentText(s["category"] or "무선")
            self.on_cat()
            self.refresh_models()
            self.f_model.setCurrentText(s["model"] or "")
            self.f_plan.setCurrentText(s["plan"] or "")
            self.f_join.setCurrentText(s["join_type"] or "신규")
            if s["discount_type"]:
                self.f_disc.setCurrentText(s["discount_type"])
            self.f_customer.setText(s["customer"] or "")
            self.f_phone4.setText(s["phone4"] or "")
            self.f_memo.setText(s["memo"] or "")
            for w, f in ((self.f_release, "release_price"), (self.f_public, "public_subsidy"),
                         (self.f_rebate, "rebate"), (self.f_deduct, "deduction"), (self.f_adjust, "adjust"),
                         (self.f_support, "support"), (self.f_extra, "extra_income")):
                w.setValue(s[f] or 0)
            self.f_flags.set(s["flags"] or "")
            self.f_phone.setText(s["phone"] or "")
            self.f_months.setCurrentIndex(max(0, self.f_months.findData(s["months"] if s["months"] is not None else 24)))
        finally:
            self.lock = False
        self.editing_id = sid
        self.b_save.setText(f"수정 저장 (#{sid})")
        self.lbl_policy.setText(f"✏ #{sid} 수정 중 — 취소하려면 '새로 입력'")

    def prefill(self, d):
        """최적 대리점 탭에서 넘어온 조건으로 입력칸 채우기"""
        self.reset_form()
        self.lock = True
        try:
            self.f_date.setDate(QDate.fromString(d["date"], "yyyy-MM-dd"))
            i = self.f_agency.findData(d["agency_id"])
            if i >= 0:
                self.f_agency.setCurrentIndex(i)
            self.f_cat.setCurrentText(d["category"])
            self.on_cat()
            self.refresh_models()
            self.f_model.setCurrentText(d["model"])
            self.refresh_plans()
            self.f_plan.setCurrentText(d["plan"])
            self.f_join.setCurrentText(d["join"])
            self.f_flags.set(d["flags"])
        finally:
            self.lock = False
        self.load_policy(silent=False)
        self.f_customer.setFocus()

    def edit_selected(self):
        rows = selected_rows(self.table)
        if rows:
            self.load_sale(row_id(self.table, rows[0]))

    def delete_selected(self):
        ids = [row_id(self.table, r) for r in selected_rows(self.table)]
        if not ids or not ask(self, f"{len(ids)}건을 삭제할까요? (되돌릴 수 없어요)"):
            return
        for i in ids:
            self.db.x("DELETE FROM sales WHERE id=?", (i,), commit=False)
        self.db.commit()
        self.load_list()
        self.db.notify("sales")

    def load_list(self):
        rows = self.db.q(
            "SELECT s.*, a.name AS agency, st.name AS staff FROM sales s "
            "LEFT JOIN agencies a ON a.id=s.agency_id LEFT JOIN staff st ON st.id=s.staff_id "
            "WHERE substr(s.sale_date,1,7)=? ORDER BY s.sale_date DESC, s.id DESC", (mstr(self.month),))
        data, ids, colors = [], [], []
        for s in rows:
            ids.append(s["id"])
            data.append([s["id"], s["sale_date"], s["staff"] or "", s["agency"] or "", s["category"], s["carrier"],
                         s["model"], s["plan"], s["join_type"], s["discount_type"] or "", s["customer"], s["phone4"],
                         s["rebate"], s["deduction"], s["adjust"] or 0, s["support"], s["extra_income"], s["margin"],
                         "" if s["paid"] is None else s["paid"], s["flags"] or "", s["memo"]])
            colors.append(RED if s["margin"] < 0 else None)
        fill_table(self.table, data, ids=ids, colors=colors)
        tot = sum(s["margin"] for s in rows)
        nw = sum(1 for s in rows if s["category"] == "유선")
        self.list_sum.setText(f"{len(rows)}건 (무선 {len(rows) - nw} · 유선 {nw}) · 마진 합계 {won(tot)}원")


# ============================================================
# 탭: 정산 대조
# ============================================================
class SettleTab(QWidget):
    HEAD = ["ID", "개통일", "대리점", "구분", "담당", "고객", "뒷4", "모델/상품", "가입유형",
            "받을 돈", "실제 입금 ✎", "입금일 ✎", "차액", "상태"]
    PAID_COL, DATE_COL = 10, 11

    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>대리점에서 돈이 제대로 들어왔는지</b> 확인하는 화면입니다. 노랑 = 아직 안 들어옴, 빨강 = 덜/더 들어옴, 초록 = 정상.<br>"
                             "들어온 건은 선택 후 [받을 돈 그대로 입금처리], 금액이 다르면 '실제 입금' 칸을 더블클릭해 고치고 [저장]. "
                             "대리점 정산 엑셀이 있으면 [입금내역 가져오기]로 한 번에 맞춰집니다."))
        top = QHBoxLayout()
        self.month = date_edit("yyyy-MM")
        self.agency = QComboBox()
        self.status = QComboBox()
        self.status.addItems(["전체", "미입금", "차액 발생", "완료"])
        for w in [QLabel("개통 월"), self.month, QLabel(" 대리점"), self.agency, QLabel(" 상태"), self.status]:
            top.addWidget(w)
        top.addWidget(btn("조회", self.load, primary=True))
        top.addStretch()
        lay.addLayout(top)
        bar = QHBoxLayout()
        bar.addWidget(btn("✔ 선택 건 받을 돈 그대로 입금처리", self.mark_paid))
        bar.addWidget(btn("선택 건 입금 취소", self.unmark))
        bar.addWidget(btn("입금내역 가져오기(엑셀)", self.import_payments))
        bar.addWidget(btn("💾 고친 금액 저장", self.save, primary=True))
        bar.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"정산대조_{mstr(self.month)}.csv")))
        bar.addStretch()
        lay.addLayout(bar)
        self.table = make_table(self.HEAD, editable=True)
        lay.addWidget(self.table)
        self.summary = QLabel()
        self.summary.setStyleSheet("font-weight:bold;padding:4px;font-size:11pt")
        lay.addWidget(self.summary)
        lay.addWidget(hint("받을 돈 = 리베이트 - 차감 + 부가·차감 가감.  입금내역 엑셀은 '개통일', '휴대폰번호(뒷자리)', '금액' 열이 있으면 자동으로 맞춥니다. "
                           "대리점을 먼저 고르면 그 대리점 건만 맞춥니다."))
        for cb in (self.status, self.agency):
            cb.currentIndexChanged.connect(lambda *_: self.load())
        self.month.dateChanged.connect(lambda *_: self.load())

    def refresh(self):
        fill_combo(self.agency, self.db.agency_items(), "전체")
        self.load()

    def load(self):
        sql = ("SELECT s.*, a.name AS agency, st.name AS staff FROM sales s "
               "LEFT JOIN agencies a ON a.id=s.agency_id LEFT JOIN staff st ON st.id=s.staff_id "
               "WHERE substr(s.sale_date,1,7)=?")
        params = [mstr(self.month)]
        aid = self.agency.currentData()
        if aid:
            sql += " AND s.agency_id=?"
            params.append(aid)
        rows = self.db.q(sql + " ORDER BY s.sale_date, s.id", params)
        filt = self.status.currentText()
        data, ids, colors = [], [], []
        tot_exp = tot_paid = un_n = un_amt = df_n = df_amt = 0
        for s in rows:
            exp = s["rebate"] - s["deduction"] + (s["adjust"] or 0)
            paid = s["paid"]
            st = pay_status(exp, paid)
            tot_exp += exp
            if paid is None:
                un_n += 1
                un_amt += exp
            else:
                tot_paid += paid
                if paid != exp:
                    df_n += 1
                    df_amt += paid - exp
            if (filt == "미입금" and st != "미입금") or (filt == "완료" and st != "완료") or \
                    (filt == "차액 발생" and st not in ("차감/부족", "초과입금")):
                continue
            paid_item = NumItem(paid or 0, "" if paid is None else won(paid), editable=True)
            date_item = txt_item(s["paid_date"] or "", editable=True)
            diff = NumItem(paid - exp, signed(paid - exp)) if paid is not None else txt_item("")
            ids.append(s["id"])
            data.append([s["id"], s["sale_date"], s["agency"] or "", s["category"], s["staff"] or "", s["customer"],
                         s["phone4"], s["model"], s["join_type"], exp, paid_item, date_item, diff, st])
            colors.append({"미입금": YELLOW, "완료": GREEN}.get(st, RED))
        fill_table(self.table, data, ids=ids, colors=colors)
        self.summary.setText(f"총 {len(rows)}건 · 받을 돈 {won(tot_exp)}원 · 들어온 돈 {won(tot_paid)}원   |   "
                             f"미입금 {un_n}건 {won(un_amt)}원   |   차액 {df_n}건 {signed(df_amt)}원")

    def save(self):
        if staff_block(self, self.db, "입금 처리"):
            return
        n = 0
        for r in range(self.table.rowCount()):
            sid = row_id(self.table, r)
            ptxt = self.table.item(r, self.PAID_COL).text().strip()
            dtxt = self.table.item(r, self.DATE_COL).text().strip()
            paid = None if ptxt == "" else to_int(ptxt)
            pdate = (norm_date(dtxt) or today()) if paid is not None else None
            cur = self.db.one("SELECT paid, paid_date FROM sales WHERE id=?", (sid,))
            if cur and (cur["paid"] != paid or (cur["paid_date"] or None) != pdate):
                self.db.x("UPDATE sales SET paid=?, paid_date=? WHERE id=?", (paid, pdate, sid), commit=False)
                n += 1
        self.db.commit()
        self.load()
        info(self, f"{n}건 저장했습니다.")

    def mark_paid(self):
        if staff_block(self, self.db, "입금 처리"):
            return
        ids = [row_id(self.table, r) for r in selected_rows(self.table)]
        if not ids:
            return warn(self, "표에서 입금된 건을 먼저 클릭해서 선택하세요.\n(여러 건은 Ctrl 또는 Shift 누르고 클릭)")
        for i in ids:
            self.db.x("UPDATE sales SET paid=rebate-deduction+COALESCE(adjust,0), paid_date=? WHERE id=?", (today(), i), commit=False)
        self.db.commit()
        self.load()

    def unmark(self):
        if staff_block(self, self.db, "입금 처리"):
            return
        ids = [row_id(self.table, r) for r in selected_rows(self.table)]
        if not ids or not ask(self, f"{len(ids)}건을 '미입금'으로 되돌릴까요?"):
            return
        for i in ids:
            self.db.x("UPDATE sales SET paid=NULL, paid_date=NULL WHERE id=?", (i,), commit=False)
        self.db.commit()
        self.load()

    def import_payments(self):
        if staff_block(self, self.db, "입금내역 가져오기"):
            return
        path, _ = QFileDialog.getOpenFileName(self, "대리점 입금내역 파일", "", "엑셀/CSV (*.xlsx *.xlsm *.xls *.csv)")
        if not path:
            return
        try:
            sheets = read_sheets(path)
        except Exception as ex:
            return warn(self, str(ex))
        aliases = [["개통일", "개통일자", "일자", "날짜"], ["뒷", "휴대폰", "전화", "개통번호", "번호"],
                   ["입금", "정산", "지급", "금액"]]
        sums = {}
        for _name, data in sheets:
            hi, mp = map_headers(data, aliases, required=0)
            if hi is None or 1 not in mp or 2 not in mp:
                continue
            for row in data[hi + 1:]:
                def g(c):
                    i = mp[c]
                    return row[i] if i < len(row) else ""
                d, p = norm_date(g(0)), last4(g(1))
                if d and len(p) == 4:
                    sums[(d, p)] = sums.get((d, p), 0) + to_int(g(2))
        if not sums:
            return warn(self, "'개통일', '휴대폰번호', '금액' 열을 찾지 못했습니다.\n파일 첫 줄에 이 제목들이 있는지 확인해 주세요.")
        aid = self.agency.currentData()
        matched, unmatched = 0, []
        for (d, p), amt in sums.items():
            sql = "SELECT id FROM sales WHERE sale_date=? AND phone4=?"
            params = [d, p]
            if aid:
                sql += " AND agency_id=?"
                params.append(aid)
            cand = self.db.q(sql + " ORDER BY (paid IS NOT NULL), id", params)
            if not cand:
                unmatched.append(f"{d} / {p} / {won(amt)}원")
                continue
            self.db.x("UPDATE sales SET paid=?, paid_date=? WHERE id=?", (amt, today(), cand[0]["id"]), commit=False)
            matched += 1
        self.db.commit()
        self.load()
        msg = f"✅ {matched}건 맞춰서 입금 처리했습니다."
        if unmatched:
            msg += f"\n\n개통 등록에서 못 찾은 {len(unmatched)}건 (개통 등록 누락 확인):\n" + "\n".join(unmatched[:25])
            if len(unmatched) > 25:
                msg += f"\n… 외 {len(unmatched) - 25}건"
        info(self, msg)


# ============================================================
# 탭: 직원 · 수당
# ============================================================
class StaffTab(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        outer = QVBoxLayout(self)
        outer.addWidget(banner("<b>직원 등록과 월별 수당</b>. 수당 = 건당 수당 × 개통 건수 + 마진 × 비율(%). "
                               "그만둔 직원은 재직을 N으로 바꾸면 목록에서 빠지고 기록은 남습니다."))
        split = QSplitter()
        outer.addWidget(split, 1)
        left = QGroupBox("직원 관리")
        ll = QVBoxLayout(left)
        self.staff_table = make_table(["이름", "건당 수당(원)", "마진 비율(%)", "재직(Y/N)"], editable=True)
        ll.addWidget(self.staff_table)
        b = QHBoxLayout()
        b.addWidget(btn("직원 추가", self.add_staff))
        b.addWidget(btn("💾 저장", self.save_staff, primary=True))
        ll.addLayout(b)
        split.addWidget(left)
        right = QGroupBox("월별 실적 · 수당")
        rl = QVBoxLayout(right)
        t = QHBoxLayout()
        self.month = date_edit("yyyy-MM")
        t.addWidget(QLabel("월"))
        t.addWidget(self.month)
        t.addWidget(btn("조회", self.calc, primary=True))
        t.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"직원수당_{mstr(self.month)}.csv")))
        t.addStretch()
        rl.addLayout(t)
        self.table = make_table(["직원", "건수", "무선", "유선", "마진 합계", "건당 수당", "마진 수당", "총 수당"])
        rl.addWidget(self.table)
        split.addWidget(right)
        split.setSizes([420, 900])
        self.month.dateChanged.connect(lambda *_: self.calc())

    def refresh(self):
        self.load_staff()
        self.calc()

    def load_staff(self):
        rows = self.db.staff()
        self.staff_table.setSortingEnabled(False)
        self.staff_table.setRowCount(len(rows))
        for r, s in enumerate(rows):
            it = txt_item(s["name"], editable=True)
            it.setData(ID_ROLE, s["id"])
            self.staff_table.setItem(r, 0, it)
            self.staff_table.setItem(r, 1, txt_item(won(s["per_unit"]), editable=True))
            self.staff_table.setItem(r, 2, txt_item(f"{s['margin_rate']:g}", editable=True))
            self.staff_table.setItem(r, 3, txt_item("Y" if s["active"] else "N", editable=True))
        self.staff_table.resizeColumnsToContents()

    def add_staff(self):
        r = self.staff_table.rowCount()
        self.staff_table.insertRow(r)
        for c, v in enumerate(["", "0", "0", "Y"]):
            self.staff_table.setItem(r, c, txt_item(v, editable=True))
        self.staff_table.editItem(self.staff_table.item(r, 0))

    def save_staff(self):
        if staff_block(self, self.db, "직원 목록 수정"):
            return
        try:
            for r in range(self.staff_table.rowCount()):
                name = self.staff_table.item(r, 0).text().strip()
                if not name:
                    continue
                sid = self.staff_table.item(r, 0).data(ID_ROLE)
                per = to_int(self.staff_table.item(r, 1).text())
                rate = to_num(self.staff_table.item(r, 2).text())
                active = 0 if self.staff_table.item(r, 3).text().strip().upper() == "N" else 1
                if sid:
                    self.db.x("UPDATE staff SET name=?, per_unit=?, margin_rate=?, active=? WHERE id=?",
                              (name, per, rate, active, sid), commit=False)
                else:
                    self.db.x("INSERT INTO staff(name,per_unit,margin_rate,active) VALUES(?,?,?,?)",
                              (name, per, rate, active), commit=False)
            self.db.commit()
        except sqlite3.IntegrityError:
            self.db.con.rollback()
            return warn(self, "같은 이름의 직원이 있습니다. (예: 김민수A 처럼 구분해 주세요)")
        self.refresh()
        self.db.notify("staff")
        info(self, "저장했습니다.")

    def calc(self):
        rows = self.db.q(
            "SELECT s.category, s.margin, s.staff_id, st.name, st.per_unit, st.margin_rate FROM sales s "
            "LEFT JOIN staff st ON st.id=s.staff_id WHERE substr(s.sale_date,1,7)=?", (mstr(self.month),))
        agg = {}
        for s in rows:
            a = agg.setdefault(s["staff_id"], dict(name=s["name"] or "(미지정)", per=s["per_unit"] or 0,
                                                   rate=s["margin_rate"] or 0, n=0, w=0, l=0, m=0, pm=0))
            a["n"] += 1
            a["l" if s["category"] == "유선" else "w"] += 1
            a["m"] += s["margin"]
            a["pm"] += max(0, s["margin"])
        data, tot = [], [0] * 7
        for a in sorted(agg.values(), key=lambda x: -x["n"]):
            uc = a["per"] * a["n"]
            rc = int(a["pm"] * a["rate"] / 100)
            vals = [a["n"], a["w"], a["l"], a["m"], uc, rc, uc + rc]
            tot = [x + y for x, y in zip(tot, vals)]
            data.append([a["name"]] + vals)
        if data:
            data.append(["합계"] + tot)
        fill_table(self.table, data, bold_rows={len(data) - 1} if data else set())


# ============================================================
# 탭: 월별 리포트 (+ 경비)
# ============================================================
class ReportTab(QWidget):
    COLS = [("n", "개통건수"), ("w", "무선"), ("l", "유선"), ("rebate", "리베이트"), ("ded", "차감"),
            ("sup", "고객지원금"), ("extra", "기타수입"), ("margin", "마진(매출총이익)"), ("comm", "직원수당"),
            ("exp", "경비"), ("profit", "순이익"), ("paid", "대리점 입금"), ("unpaid", "못 받은 돈")]

    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>한 달에 얼마 남았는지</b> 보는 화면입니다. 순이익 = 마진 - 직원수당 - 경비. "
                             "임대료·월급·광고비 같은 경비는 아래에 월별로 넣어주세요. '못 받은 돈' = 대리점에서 아직 안 들어온 금액."))
        top = QHBoxLayout()
        self.year = QSpinBox()
        self.year.setRange(2020, 2100)
        self.year.setValue(QDate.currentDate().year())
        top.addWidget(QLabel("연도"))
        top.addWidget(self.year)
        top.addWidget(btn("조회", self.calc, primary=True))
        top.addWidget(btn("엑셀로 내보내기", lambda: export_csv(self, self.table, f"월별리포트_{self.year.value()}.csv")))
        top.addStretch()
        lay.addLayout(top)
        self.table = make_table(["월"] + [c[1] for c in self.COLS])
        lay.addWidget(self.table, 3)
        box = QGroupBox("월 경비 입력")
        bl = QVBoxLayout(box)
        e = QHBoxLayout()
        self.e_month = date_edit("yyyy-MM")
        self.e_item = QLineEdit()
        self.e_item.setPlaceholderText("항목 (예: 임대료)")
        self.e_amount = money_spin(0, 1_000_000_000)
        for w in [QLabel("월"), self.e_month, self.e_item, self.e_amount]:
            e.addWidget(w)
        e.addWidget(btn("추가", self.add_expense, primary=True))
        e.addWidget(btn("선택 삭제", self.del_expense))
        e.addWidget(btn("지난달 경비 복사", self.copy_prev))
        e.addStretch()
        bl.addLayout(e)
        self.e_table = make_table(["항목", "금액"])
        bl.addWidget(self.e_table)
        lay.addWidget(box, 2)
        self.e_month.dateChanged.connect(lambda *_: self.load_expenses())
        self.year.valueChanged.connect(lambda *_: self.calc())

    def refresh(self):
        self.calc()
        self.load_expenses()

    def calc(self):
        y = str(self.year.value())
        sales = self.db.q("SELECT s.*, st.per_unit, st.margin_rate FROM sales s LEFT JOIN staff st ON st.id=s.staff_id "
                          "WHERE substr(s.sale_date,1,4)=?", (y,))
        exps = {r["m"]: r["t"] for r in self.db.q(
            "SELECT substr(month,6,2) AS m, SUM(amount) AS t FROM expenses WHERE substr(month,1,4)=? GROUP BY m", (y,))}
        keys = [c[0] for c in self.COLS]
        agg = {m: dict.fromkeys(keys, 0) for m in range(1, 13)}
        for s in sales:
            a = agg[int(s["sale_date"][5:7])]
            a["n"] += 1
            a["l" if s["category"] == "유선" else "w"] += 1
            a["rebate"] += s["rebate"]
            a["ded"] += s["deduction"]
            a["sup"] += s["support"]
            a["extra"] += s["extra_income"]
            a["margin"] += s["margin"]
            a["comm"] += commission(s["per_unit"], s["margin_rate"], s["margin"])
            a["paid"] += s["paid"] or 0
            a["unpaid"] += (s["rebate"] - s["deduction"] + (s["adjust"] or 0)) - (s["paid"] or 0)
        data, cc = [], {}
        tot = dict.fromkeys(keys, 0)
        pi, ui = keys.index("profit") + 1, keys.index("unpaid") + 1
        for m in range(1, 13):
            a = agg[m]
            a["exp"] = exps.get(f"{m:02d}", 0) or 0
            a["profit"] = a["margin"] - a["comm"] - a["exp"]
            for k in keys:
                tot[k] += a[k]
            r = len(data)
            if a["profit"] < 0:
                cc[(r, pi)] = RED
            if a["unpaid"] > 0:
                cc[(r, ui)] = YELLOW
            data.append([f"{y}-{m:02d}"] + [a[k] for k in keys])
        data.append(["합계"] + [tot[k] for k in keys])
        fill_table(self.table, data, cell_colors=cc, bold_rows={12})

    def load_expenses(self):
        rows = self.db.q("SELECT * FROM expenses WHERE month=? ORDER BY id", (mstr(self.e_month),))
        fill_table(self.e_table, [[r["item"], r["amount"]] for r in rows], ids=[r["id"] for r in rows])

    def add_expense(self):
        item = self.e_item.text().strip()
        if not item or not self.e_amount.value():
            return warn(self, "항목과 금액을 넣어주세요.")
        self.db.x("INSERT INTO expenses(month,item,amount) VALUES(?,?,?)",
                  (mstr(self.e_month), item, self.e_amount.value()))
        self.e_item.clear()
        self.e_amount.setValue(0)
        self.load_expenses()
        self.calc()

    def del_expense(self):
        for i in [row_id(self.e_table, r) for r in selected_rows(self.e_table)]:
            self.db.x("DELETE FROM expenses WHERE id=?", (i,), commit=False)
        self.db.commit()
        self.load_expenses()
        self.calc()

    def copy_prev(self):
        cur = mstr(self.e_month)
        prev = self.e_month.date().addMonths(-1).toString("yyyy-MM")
        rows = self.db.q("SELECT item, amount FROM expenses WHERE month=?", (prev,))
        if not rows:
            return warn(self, f"{prev} 경비가 없습니다.")
        if not ask(self, f"{prev} 경비 {len(rows)}건을 {cur}로 복사할까요?"):
            return
        for r in rows:
            self.db.x("INSERT INTO expenses(month,item,amount) VALUES(?,?,?)", (cur, r["item"], r["amount"]),
                      commit=False)
        self.db.commit()
        self.load_expenses()
        self.calc()


# ============================================================
# 탭: 설정
# ============================================================
class SettingsTab(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        lay = QVBoxLayout(self)
        lay.addWidget(banner("<b>처음 한 번만 해두면 되는 설정</b>입니다. 대리점 이름을 실제 이름으로 바꾸고, AI 키를 넣어주세요. "
                             "모르겠으면 '❓ 사용법' 탭을 보세요."))
        top = QHBoxLayout()

        box = QGroupBox("기본 설정")
        g = QGridLayout(box)
        self.store = QLineEdit()
        self.margin = money_spin(0, 5_000_000)
        self.unit = QComboBox()
        for t, v in [("절사 없음", 1), ("천원 단위", 1000), ("만원 단위", 10000)]:
            self.unit.addItem(t, v)
        g.addWidget(QLabel("매장 이름 (가격표 제목)"), 0, 0)
        g.addWidget(self.store, 0, 1)
        g.addWidget(QLabel("목표마진 (한 건당 남길 돈)"), 1, 0)
        g.addWidget(self.margin, 1, 1)
        g.addWidget(QLabel("고객지원금 끝자리 정리"), 2, 0)
        g.addWidget(self.unit, 2, 1)
        g.addWidget(btn("💾 저장", self.save_settings, primary=True), 3, 1)
        top.addWidget(box, 1)

        abox = QGroupBox("AI 정책 자동읽기 설정")
        ag = QGridLayout(abox)
        self.api_key = QLineEdit()
        self.api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key.setPlaceholderText("Claude(sk-ant-) · GPT(sk-) · Gemini(AIza) 중 아무 키나 붙여넣기")
        self.model = QLineEdit()
        self.model.setPlaceholderText("비워두면 키에 맞는 모델 자동 선택")
        self.ai_detect = QLabel("")
        self.ai_detect.setStyleSheet("font-weight:bold;color:#1f5fbf")
        self.api_key.textChanged.connect(lambda *_: self._show_detect())
        show = btn("보기", lambda: self.api_key.setEchoMode(
            QLineEdit.EchoMode.Normal if self.api_key.echoMode() == QLineEdit.EchoMode.Password
            else QLineEdit.EchoMode.Password))
        ag.addWidget(QLabel("AI 키"), 0, 0)
        ag.addWidget(self.api_key, 0, 1)
        ag.addWidget(show, 0, 2)
        ag.addWidget(QLabel("모델"), 1, 0)
        ag.addWidget(self.model, 1, 1)
        ag.addWidget(btn("💾 저장", self.save_ai, primary=True), 2, 1)
        ag.addWidget(btn("연결 테스트", self.test_ai), 2, 2)
        ag.addWidget(self.ai_detect, 3, 0, 1, 3)
        ag.addWidget(hint("Claude·GPT·Gemini 키 중 무엇을 넣어도 알아서 인식합니다. 직원마다 다른 키를 써도 됩니다. "
                          "키는 이 컴퓨터에만 저장되고 다른 PC로 넘어가지 않습니다. 모델 칸은 보통 비워두세요."), 4, 0, 1, 3)
        top.addWidget(abox, 1)
        lay.addLayout(top)

        agbox = QGroupBox("대리점 관리 (이름을 더블클릭해서 실제 대리점 이름으로 바꾸세요)")
        al = QVBoxLayout(agbox)
        self.a_table = make_table(["대리점 이름", "통신사", "담당자/연락처", "메모"], editable=True)
        al.addWidget(self.a_table)
        b = QHBoxLayout()
        self.ag_btns = [btn("➕ 대리점 추가", self.add_agency), btn("선택 삭제", self.del_agency),
                        btn("💾 대리점 저장", self.save_agencies, primary=True)]
        for x in self.ag_btns:
            b.addWidget(x)
        b.addStretch()
        al.addLayout(b)
        lay.addWidget(agbox, 1)

        # 직원 PC에 보이는 간단한 업데이트 안내 (토큰·배포 버튼 없음)
        self.staff_box = QGroupBox("🌐 자동 업데이트")
        stl = QHBoxLayout(self.staff_box)
        self.staff_upd_lbl = QLabel("")
        stl.addWidget(self.staff_upd_lbl)
        stl.addStretch()
        stl.addWidget(btn("🔄 새 버전 확인", self.check_github))
        stl.addWidget(btn("🔒 사장님 PC로 전환", self.become_owner))
        lay.addWidget(self.staff_box)

        gbox = QGroupBox("🌐 인터넷 자동 업데이트 (GitHub) — 카카오톡처럼 직원 PC에 업데이트 창")
        self.gbox = gbox
        gg = QGridLayout(gbox)
        self.gh_repo = QLineEdit()
        self.gh_repo.setPlaceholderText("GitHub아이디/저장소이름   예: yeogida/phone-policy")
        self.gh_token = QLineEdit()
        self.gh_token.setEchoMode(QLineEdit.EchoMode.Password)
        self.gh_token.setPlaceholderText("github_pat_ 로 시작하는 토큰 (사장님 PC에만 저장, 직원 PC엔 필요 없음)")
        gg.addWidget(QLabel("저장소"), 0, 0)
        gg.addWidget(self.gh_repo, 0, 1)
        gg.addWidget(QLabel("토큰"), 1, 0)
        gg.addWidget(self.gh_token, 1, 1)
        gb = QHBoxLayout()
        gb.addWidget(btn("💾 저장 + 연결 확인", self.save_github, primary=True))
        self.b_gh_pub = btn("📤 직원들에게 업데이트 배포", self.publish_github, primary=True)
        gb.addWidget(self.b_gh_pub)
        gb.addWidget(btn("🔄 새 버전 확인", self.check_github))
        gb.addStretch()
        gg.addLayout(gb, 2, 0, 1, 2)
        gg.addWidget(hint("사장님이 [📤 업데이트 배포]를 누르면, 직원 PC는 켤 때와 10분마다 확인해서 "
                          "'새 버전이 있습니다. 업데이트할까요?' 창이 뜨고 [예]를 누르면 인터넷에서 받아 자동 설치됩니다. "
                          "정책·개통·API키 같은 매장 데이터는 올라가지 않고 프로그램만 올라갑니다. (설정 방법: ❓ 사용법 12번)"),
                     3, 0, 1, 2)
        lay.addWidget(gbox)

        sbox = QGroupBox("직원 PC 공유 폴더 (선택) — 정책·개통 내역 주고받기")
        sg = QGridLayout(sbox)
        self.role = QComboBox()
        self.role.addItem("사장님 PC (정책 입력·배포)", "관리자")
        self.role.addItem("직원 PC (자동으로 받기)", "직원")
        self.pc_name = QLineEdit()
        self.pc_name.setPlaceholderText("예: 김민수PC")
        self.share_dir = QLineEdit()
        self.share_dir.setPlaceholderText("예: G:\\내 드라이브\\폰정책공유   또는   \\\\사장님PC\\폰정책공유")
        self.role_lbl = QLabel("이 PC 역할")
        sg.addWidget(self.role_lbl, 0, 0)
        sg.addWidget(self.role, 0, 1)
        sg.addWidget(QLabel("이 PC 이름"), 0, 2)
        sg.addWidget(self.pc_name, 0, 3)
        sg.addWidget(QLabel("공유 폴더"), 1, 0)
        sg.addWidget(self.share_dir, 1, 1, 1, 3)
        sg.addWidget(btn("폴더 선택", self.pick_share), 1, 4)
        sb = QHBoxLayout()
        sb.addWidget(btn("💾 공유 설정 저장", self.save_share, primary=True))
        self.b_publish = btn("📤 지금 직원들에게 배포", self.publish_now)
        sb.addWidget(self.b_publish)
        sb.addWidget(btn("🔄 지금 업데이트 확인", self.check_now))
        self.ver_lbl = QLabel()
        sb.addStretch()
        sb.addWidget(self.ver_lbl)
        sg.addLayout(sb, 2, 0, 1, 5)
        sg.addWidget(hint("사장님 PC에서 정책을 저장하면 공유 폴더로 자동 전달되고, 직원 PC는 켤 때와 10분마다 자동으로 받아옵니다. "
                          "프로그램을 새 버전으로 바꾸면 직원 PC에 '업데이트할까요?' 창이 뜹니다. "
                          "직원 PC에서 등록한 개통 건은 사장님 PC의 정산·리포트에 자동으로 모입니다. (설정 방법: ❓ 사용법 11번)"),
                     3, 0, 1, 5)
        lay.addWidget(sbox)
        self.main = None    # MainWindow가 연결

        kbox = QGroupBox("⏰ 유지기간 기본값 · 💾 자동 백업")
        kg = QGridLayout(kbox)
        self.k_add, self.k_plan, self.k_line = QSpinBox(), QSpinBox(), QSpinBox()
        for w in (self.k_add, self.k_plan, self.k_line):
            w.setRange(1, 1000)
            w.setSuffix(" 일")
        self.bk_dir = QLineEdit()
        self.bk_dir.setPlaceholderText("예: G:\\내 드라이브\\폰정책백업  (비우면 이 PC에만 백업)")
        self.bk_lbl = QLabel("")
        kg.addWidget(QLabel("부가서비스 유지"), 0, 0)
        kg.addWidget(self.k_add, 0, 1)
        kg.addWidget(QLabel("요금제 유지"), 0, 2)
        kg.addWidget(self.k_plan, 0, 3)
        kg.addWidget(QLabel("회선 유지(해지 환수)"), 0, 4)
        kg.addWidget(self.k_line, 0, 5)
        kg.addWidget(QLabel("백업 폴더(구글 드라이브 추천)"), 1, 0)
        kg.addWidget(self.bk_dir, 1, 1, 1, 4)
        kg.addWidget(btn("폴더 선택", lambda: self.bk_dir.setText(
            os.path.normpath(QFileDialog.getExistingDirectory(self, "백업 폴더") or self.bk_dir.text()))), 1, 5)
        kg.addWidget(btn("💾 저장", self.save_keep, primary=True), 2, 1)
        kg.addWidget(btn("지금 백업하기", self.backup_now), 2, 2)
        kg.addWidget(self.bk_lbl, 2, 3, 1, 3)
        kg.addWidget(hint("대리점 규칙에 '93일' 같은 날짜가 있으면 그걸 우선 쓰고, 없으면 이 기본값을 씁니다. "
                          "데이터는 하루 한 번 자동 백업(최근 30일치)되고, 백업 폴더를 정하면 거기에도 복사됩니다."), 3, 0, 1, 6)
        lay.addWidget(kbox)

        dbox = QGroupBox("데이터 보관")
        dl = QHBoxLayout(dbox)
        lb = QLabel(f"저장 위치: {db.path}")
        lb.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        dl.addWidget(lb)
        dl.addStretch()
        dl.addWidget(btn("📦 백업 파일 만들기 (USB·메일에 보관)", self.backup))
        dl.addWidget(btn("📥 새 버전 파일 적용", self.apply_file,
                         tip="새로 받은 phone_policy_manager.py 파일을 골라서 이 PC 프로그램을 업데이트합니다."))
        if not FROZEN:
            pass
        self.b_bat = btn("💿 직원용 설치 bat 만들기 (바로)", self.build_installer)
        dl.addWidget(self.b_bat)
        if not FROZEN:
            self.b_setup = btn("🧙 직원용 Setup.exe 만들기", self.build_setup_exe, primary=True,
                               tip="카카오톡처럼 다음→설치→마침 화면이 뜨는 Setup.exe를 만듭니다 (약 10분)")
            dl.addWidget(self.b_setup)
        lay.addWidget(dbox)

    def refresh(self):
        self.store.setText(self.db.get("store_name", "우리매장"))
        self.margin.setValue(to_int(self.db.get("target_margin", "100000")))
        self.unit.setCurrentIndex(max(0, self.unit.findData(to_int(self.db.get("round_unit", "10000")))))
        self.api_key.setText(self.db.get("api_key", ""))
        self.k_add.setValue(to_int(self.db.get("keep_addon_days", "93")) or 93)
        self.k_plan.setValue(to_int(self.db.get("keep_plan_days", "93")) or 93)
        self.k_line.setValue(to_int(self.db.get("keep_line_days", "183")) or 183)
        self.bk_dir.setText(self.db.get("backup_dir", ""))
        self.bk_lbl.setText(f"마지막 백업: {self.db.get('last_backup', '아직 없음')}")
        self.gh_repo.setText(gh_repo(self.db))
        self.gh_token.setText(self.db.get("gh_token", ""))
        self.b_gh_pub.setVisible(not self.db.is_staff_pc())
        self.role.setCurrentIndex(max(0, self.role.findData(self.db.get("role", "관리자"))))
        self.pc_name.setText(self.db.get("pc_name", "") or socket.gethostname())
        self.share_dir.setText(self.db.get("share_dir", ""))
        self._role_ui()
        m = self.db.get("ai_model", "")
        self.model.setText("" if m == DEFAULT_MODEL else m)
        self._show_detect()
        rows = self.db.agencies()
        self.a_table.setSortingEnabled(False)
        self.a_table.setRowCount(0)
        for a in rows:
            self._agency_row(a["id"], a["name"], a["carrier"], a["contact"], a["memo"])
        self.a_table.resizeColumnsToContents()

    def _agency_row(self, aid, name, carrier, contact, memo):
        r = self.a_table.rowCount()
        self.a_table.insertRow(r)
        it = txt_item(name, editable=True)
        it.setData(ID_ROLE, aid)
        self.a_table.setItem(r, 0, it)
        cb = QComboBox()
        cb.setEditable(True)
        cb.addItems(CARRIERS)
        cb.setCurrentText(carrier or "SKT")
        self.a_table.setCellWidget(r, 1, cb)
        self.a_table.setItem(r, 2, txt_item(contact or "", editable=True))
        self.a_table.setItem(r, 3, txt_item(memo or "", editable=True))
        return r

    def save_github(self):
        repo = re.sub(r"^https?://github\.com/", "", self.gh_repo.text().strip()).strip("/")
        if repo and not re.fullmatch(r"[\w.-]+/[\w.-]+", repo):
            return warn(self, "저장소는 'GitHub아이디/저장소이름' 모양으로 넣어주세요.\n예: yeogida/phone-policy")
        self.db.set("gh_repo", repo)
        self.db.set("gh_token", self.gh_token.text().strip())
        if self.gh_token.text().strip():
            self.db.set("role", "관리자")
        if not repo:
            return info(self, "인터넷 업데이트를 끕니다.")
        tok = self.gh_token.text().strip()
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            d = _gh(f"{GH_API}/repos/{repo}", tok or None)
            QApplication.restoreOverrideCursor()
            if d.get("private"):
                return warn(self, "저장소가 Private(비공개)입니다. 직원 PC가 받을 수 있게 Public(공개)으로 바꿔주세요.\n"
                                  "(저장소 → Settings → 맨 아래 Change visibility)")
            if tok and not (d.get("permissions") or {}).get("push", True):
                return warn(self, "토큰에 쓰기 권한이 없습니다. 토큰 만들 때 'Contents'를 'Read and write'로 했는지,\n"
                                  "그 저장소를 선택했는지 확인해 주세요.")
            info(self, f"✅ GitHub 연결 정상입니다. ({repo})\n" +
                 ("이제 [📤 직원들에게 업데이트 배포]를 누르면 됩니다." if tok else "이 PC는 업데이트를 받기만 합니다."))
        except Exception as ex:
            QApplication.restoreOverrideCursor()
            warn(self, "❌ 연결 실패\n\n" + gh_error(ex))

    def publish_github(self):
        repo, tok = gh_repo(self.db), self.db.get("gh_token", "").strip()
        if not repo or not tok:
            return warn(self, "저장소와 토큰을 넣고 [💾 저장 + 연결 확인]을 먼저 눌러주세요.")
        notes, ok = QInputDialog.getText(self, "업데이트 배포", f"버전 {APP_VERSION}을(를) 직원들에게 배포합니다.\n"
                                                          "직원 화면에 보여줄 한 줄 설명 (비워도 됨):")
        if not ok:
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            src = publish_source(self.db).encode("utf-8")
            gh_put_file(repo, tok, "phone_policy_manager.py", src, f"폰정책매니저 {APP_VERSION}")
            ver = {"version": APP_VERSION, "notes": notes.strip(), "time": now(),
                   "sha256": hashlib.sha256(src).hexdigest()}
            gh_put_file(repo, tok, "version.json", json.dumps(ver, ensure_ascii=False, indent=1).encode("utf-8"),
                        f"version {APP_VERSION}")
            # 이 PC 프로그램 파일에도 저장소 주소를 박아둠 (다음에 만드는 설치파일이 알도록)
            me = this_program()
            with open(me, encoding="utf-8") as f:
                cur = f.read()
            if gh_repo(self.db) and f'UPDATE_REPO = "{repo}"' not in cur:
                _atomic_write(me, source_with_repo(cur, repo))
            QApplication.restoreOverrideCursor()
            info(self, f"📤 배포 완료! (버전 {APP_VERSION})\n\n직원 PC는 켤 때 또는 10분 안에 "
                       "'업데이트할까요?' 창이 뜹니다.\n\n※ 인터넷 업데이트를 처음 쓰는 거라면, 직원 PC에 "
                       "설치파일을 한 번만 새로 깔아주세요. (그 뒤로는 계속 자동)")
        except Exception as ex:
            QApplication.restoreOverrideCursor()
            warn(self, "❌ 배포 실패\n\n" + gh_error(ex))

    def check_github(self):
        if self.main:
            self.main.github_check(loud=True)

    def save_keep(self):
        self.db.set("keep_addon_days", self.k_add.value())
        self.db.set("keep_plan_days", self.k_plan.value())
        self.db.set("keep_line_days", self.k_line.value())
        d = self.bk_dir.text().strip()
        if d and not os.path.isdir(d):
            return warn(self, "백업 폴더를 찾을 수 없습니다.")
        self.db.set("backup_dir", d)
        info(self, "저장했습니다.")

    def backup_now(self):
        self.save_keep() if False else None
        self.db.set("backup_dir", self.bk_dir.text().strip())
        try:
            p = auto_backup(self.db, force=True)
        except Exception as ex:
            return warn(self, f"백업 실패: {ex}")
        self.bk_lbl.setText(f"마지막 백업: {self.db.get('last_backup', '')}")
        info(self, f"✅ 백업했습니다.\n{p}" + (f"\n+ {self.bk_dir.text()}" if self.bk_dir.text().strip() else ""))

    def pick_share(self):
        d = QFileDialog.getExistingDirectory(self, "공유 폴더 선택", self.share_dir.text() or BASE_DIR)
        if d:
            self.share_dir.setText(os.path.normpath(d))

    def save_share(self):
        d = self.share_dir.text().strip()
        if d and not os.path.isdir(d):
            return warn(self, "그 폴더를 찾을 수 없습니다. [폴더 선택]으로 골라주세요.")
        if d and os.path.normcase(os.path.abspath(d)) == os.path.normcase(BASE_DIR):
            return warn(self, "공유 폴더는 프로그램이 있는 폴더와 다른 폴더여야 합니다.\n"
                              "(프로그램과 데이터는 각 PC 바탕화면 폴더에 두고, 공유 폴더는 따로 만드세요)")
        if not self.db.is_staff_pc():
            self.db.set("role", self.role.currentData())
        self.db.set("pc_name", self.pc_name.text().strip())
        self.db.set("share_dir", d)
        self._role_ui()
        if not d:
            return info(self, "공유를 끕니다. (이 PC 혼자 사용)")
        if self.main:
            self.main.share_sync(startup=True, loud=True)

    def publish_now(self):
        if not self.main or not self.main.share.folder():
            return warn(self, "공유 폴더를 먼저 정하고 [공유 설정 저장]을 누르세요.")
        try:
            self.main.share.publish_program()
            self.main.share.publish_policies()
            n = self.main.share.collect_sales()
        except Exception as ex:
            return warn(self, f"배포 중 오류: {ex}")
        info(self, f"📤 배포했습니다.\n\n· 프로그램 버전 {APP_VERSION}\n· 대리점·정책·규칙·직원 목록\n"
                   f"· 직원 개통 내역 새로 모음: {n}건\n\n직원 PC는 켤 때 또는 10분 안에 자동으로 받습니다.")

    def check_now(self):
        if self.main:
            self.main.share_sync(startup=True, loud=True)

    def become_owner(self):
        """직원 PC → 사장님 PC 전환: GitHub 토큰을 아는 사람(사장님)만 가능"""
        tok, ok = QInputDialog.getText(self, "사장님 PC로 전환", "사장님 PC로 바꾸려면 GitHub 토큰(github_pat_…)을 넣어주세요.",
                                       QLineEdit.EchoMode.Password)
        if not ok or not tok.strip():
            return
        repo = gh_repo(self.db)
        try:
            d = _gh(f"{GH_API}/repos/{repo}", tok.strip())
            if not (d.get("permissions") or {}).get("push"):
                raise RuntimeError("권한 없음")
        except Exception:
            return warn(self, "토큰이 맞지 않습니다. 사장님만 전환할 수 있습니다.")
        self.db.set("role", "관리자")
        self.db.set("gh_token", tok.strip())
        self.refresh()
        info(self, "사장님 PC로 전환했습니다.")

    def _role_ui(self):
        staff = self.db.is_staff_pc()
        self.b_publish.setVisible(not staff)
        self.gbox.setVisible(not staff)
        self.staff_box.setVisible(staff)
        self.staff_upd_lbl.setText(f"사장님이 새 버전을 올리면 자동으로 '업데이트할까요?' 창이 뜹니다.   "
                                   f"지금 버전: {APP_VERSION}")
        self.role.setVisible(not staff)
        self.role_lbl.setVisible(not staff)
        for x in self.ag_btns + [self.b_bat] + ([self.b_setup] if hasattr(self, "b_setup") else []):
            x.setVisible(not staff)
        self.ver_lbl.setText(f"이 PC 프로그램 버전: {APP_VERSION}")

    def save_settings(self):
        self.db.set("store_name", self.store.text().strip() or "우리매장")
        self.db.set("target_margin", self.margin.value())
        self.db.set("round_unit", self.unit.currentData())
        self.db.notify("settings")
        info(self, "저장했습니다.")

    def _show_detect(self):
        key = self.api_key.text().strip()
        prov = ai_provider(key)
        if not key:
            self.ai_detect.setText("● AI 키 없음")
            self.ai_detect.setStyleSheet("font-weight:bold;color:#999")
        elif not prov:
            self.ai_detect.setText("● 키 모양을 알 수 없음 (Claude sk-ant- / GPT sk- / Gemini AIza 로 시작해야 함)")
            self.ai_detect.setStyleSheet("font-weight:bold;color:#d62828")
        else:
            self.ai_detect.setText(f"● {AI_NAMES[prov]} 키로 인식됨 → 모델 {pick_model(key, self.model.text())}")
            self.ai_detect.setStyleSheet("font-weight:bold;color:#1f5fbf")

    def save_ai(self):
        self.db.set("api_key", self.api_key.text().strip())
        self.db.set("ai_model", self.model.text().strip())
        info(self, "AI 설정을 저장했습니다. [연결 테스트]로 확인해 보세요.")

    def test_ai(self):
        key = self.api_key.text().strip()
        if not key:
            return warn(self, "AI 키를 먼저 붙여넣으세요.")
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            text, _ = api_request(key, self.model.text().strip(), "짧게 답해.",
                                  [{"type": "text", "text": "연결 확인. '정상'이라고만 답해."}],
                                  max_tokens=20, stream=False, timeout=40)
            QApplication.restoreOverrideCursor()
            self.save_ai_silent(key)
            self.ai_detect.setText(f"● {AI_NAMES[ai_provider(key)]} 연결됨 ✅ (모델 {pick_model(key, self.model.text())})")
            self.ai_detect.setStyleSheet("font-weight:bold;color:#1a7f37")
            info(self, f"✅ AI 연결 정상입니다! ({AI_NAMES[ai_provider(key)]})\n이제 정책 입력 탭에서 파일을 올려보세요.")
        except Exception as ex:
            QApplication.restoreOverrideCursor()
            self.ai_detect.setText("● 연결 안 됨 ❌")
            self.ai_detect.setStyleSheet("font-weight:bold;color:#d62828")
            warn(self, "❌ 연결 실패\n\n" + friendly_error(ex))

    def save_ai_silent(self, key):
        self.db.set("api_key", key)
        self.db.set("ai_model", self.model.text().strip())

    def add_agency(self):
        r = self._agency_row(None, "", "SKT", "", "")
        self.a_table.editItem(self.a_table.item(r, 0))

    def save_agencies(self):
        if staff_block(self, self.db, "대리점 수정"):
            return
        try:
            for r in range(self.a_table.rowCount()):
                name = self.a_table.item(r, 0).text().strip()
                if not name:
                    continue
                aid = self.a_table.item(r, 0).data(ID_ROLE)
                raw = self.a_table.cellWidget(r, 1).currentText().strip()
                carrier = norm_carrier(raw) or raw
                contact = self.a_table.item(r, 2).text().strip()
                memo = self.a_table.item(r, 3).text().strip()
                if aid:
                    self.db.x("UPDATE agencies SET name=?, carrier=?, contact=?, memo=? WHERE id=?",
                              (name, carrier, contact, memo, aid), commit=False)
                else:
                    self.db.x("INSERT INTO agencies(name,carrier,contact,memo) VALUES(?,?,?,?)",
                              (name, carrier, contact, memo), commit=False)
            self.db.commit()
        except sqlite3.IntegrityError:
            self.db.con.rollback()
            return warn(self, "같은 이름의 대리점이 두 개 있습니다. 이름을 다르게 해주세요.")
        self.refresh()
        self.db.notify("agency")
        info(self, "대리점을 저장했습니다.")

    def del_agency(self):
        rows = selected_rows(self.a_table)
        if not rows:
            return warn(self, "지울 대리점 줄을 먼저 클릭하세요.")
        r = rows[0]
        aid = self.a_table.item(r, 0).data(ID_ROLE)
        if not aid:
            self.a_table.removeRow(r)
            return
        used = self.db.one("SELECT (SELECT COUNT(*) FROM policies WHERE agency_id=?)"
                           " + (SELECT COUNT(*) FROM sales WHERE agency_id=?) AS n", (aid, aid))["n"]
        if used:
            return warn(self, "이 대리점은 정책/개통 기록이 있어서 지울 수 없습니다.\n"
                              "이름 뒤에 '(거래중단)'을 붙여두세요.")
        if ask(self, "이 대리점을 삭제할까요?"):
            self.db.x("DELETE FROM agencies WHERE id=?", (aid,))
            self.refresh()

    def apply_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "새 버전 프로그램 파일 선택", os.path.expanduser("~"),
                                              "프로그램 파일 (*.py)")
        if not path:
            return
        v = read_version(path)
        if not v:
            return warn(self, "폰정책매니저 프로그램 파일이 아닙니다.")
        if not ask(self, f"지금 버전 {APP_VERSION} → 선택한 파일 버전 {v}\n\n적용하고 다시 켤까요?"):
            return
        if self.main:
            self.main.do_update(path)

    def build_installer(self):
        data = make_setup_bat(publish_source(self.db))
        desk = os.path.join(os.path.expanduser("~"), "Desktop")
        try:
            out = _ps("[Environment]::GetFolderPath('Desktop')").stdout.strip()
            if out and os.path.isdir(out):
                desk = out
        except Exception:
            pass
        places = [os.path.join(desk, SETUP_BAT_NAME)]
        d = self.db.get("share_dir", "")
        if d and os.path.isdir(d):
            places.append(os.path.join(d, SETUP_BAT_NAME))
        for p in places:
            with open(p, "wb") as f:
                f.write(data)
        info(self, "✅ 설치파일을 만들었습니다.\n\n" + "\n".join(places) +
             "\n\n직원에게 카톡으로 보내서 더블클릭하게 하면 됩니다.\n"
             "(파이썬이 없는 PC면 자동으로 설치하고, 바탕화면에 폰정책매니저 아이콘을 만듭니다)")

    def build_setup_exe(self):
        if not ask(self, "설치 마법사(폰정책매니저-Setup.exe)를 만듭니다.\n\n"
                         "· 더블클릭하면 카카오톡처럼 '다음 → 설치 → 마침' 화면이 뜨고,\n"
                         "  바탕화면 아이콘 · 시작 메뉴 · 제어판(앱 제거)까지 등록됩니다.\n"
                         "· 받는 PC에 파이썬이 없어도 됩니다.\n\n"
                         "· 인터넷 필요, 처음엔 10분 정도 걸립니다. 다 되면 바탕화면에 생깁니다.\n\n지금 만들까요?"):
            return
        BuildDialog(self, self.db).exec()

    def backup(self):
        name = f"폰정책_백업_{datetime.datetime.now():%Y%m%d_%H%M}.db"
        path, _ = QFileDialog.getSaveFileName(self, "백업 저장", name, "백업 파일 (*.db)")
        if not path:
            return
        self.db.commit()
        shutil.copyfile(self.db.path, path)
        info(self, f"백업했습니다.\n{path}\n\n복구할 때는 이 파일 이름을 phone_policy.db 로 바꿔서 프로그램 폴더에 넣으면 됩니다.")


# ============================================================
# 탭: 사용법
# ============================================================
GUIDE_HTML = """
<style>
 body{font-family:'Malgun Gothic';font-size:11pt;line-height:1.6;color:#222}
 h1{color:#1f3a5f;font-size:18pt} h2{color:#fff;background:#1f5fbf;padding:6px 10px;font-size:13pt;margin-top:22px}
 h3{color:#1f5fbf;font-size:12pt;margin-bottom:2px} .box{background:#fff8e1;border:1px solid #f0d58c;padding:8px}
 b{color:#c0392b} li{margin-bottom:4px} td{padding:4px 8px;border:1px solid #ddd}
</style>
<h1>📘 폰정책·정산 매니저 사용법</h1>
<p>위쪽 탭(📋 정책 입력, 💰 판매가 …)을 눌러서 화면을 바꿉니다. 모든 화면 맨 위 파란 상자에 그 화면 설명이 있어요.</p>

<h2>0. 한눈에 보기 — 매일 하는 일은 딱 3개</h2>
<table>
<tr><td>🌅 대리점 정책이 오면</td><td><b>📥 정책 일괄 등록</b> → 9곳 파일 전부 끌어다 놓기 → 대리점 확인 → [🤖 전부 읽고 저장]</td></tr>
<tr><td>📣 우리 판매정책 만들기</td><td><b>📣 판매정책</b> → 요금 구간·목표마진·부가 기준 → [판매정책 만들기] → [📤 직원용 엑셀] → 카톡 배포</td></tr>
<tr><td>⭐ 켜자마자</td><td><b>⭐ 한눈에</b> → 즐겨찾기 모델의 SK·KT·LG 최고 금액 + 최근 정책 오른/내린 것</td></tr>
<tr><td>📦 모델 정리</td><td><b>📦 모델 관리</b> → 옛날·재고 없는 모델 🙈숨김, 재고 수량, ⭐즐겨찾기 ([🧹 숨길 후보 자동 체크])</td></tr>
<tr><td>🧾 손님 견적</td><td><b>🧾 견적서</b> → 모델·요금제 고르면 3사 할부원금·월 납부액 비교 → [🖼 카톡용 이미지]</td></tr>
<tr><td>🔎 모델별 마진 보기</td><td><b>🔎 정책마진 조회</b> → 'S26' 검색 → SK·KT·LG 대리점 전부의 기본·부가다함·부가없이 금액과 부가 내역이 한 화면에</td></tr>
<tr><td>🙋 손님이 오면</td><td><b>🏆 최적 대리점</b> 탭 → 모델·요금제·부가 체크 → [찾기] → 제일 많이 주는 곳 확인</td></tr>
<tr><td>📱 손님 개통하면</td><td><b>📱 개통 등록</b> 탭 → 한 건 입력 → [등록]</td></tr>
<tr><td>💵 대리점 돈 들어오면</td><td><b>🧾 정산 대조</b> 탭 → 들어온 건 선택 → [입금처리]</td></tr>
</table>
<p>월말에는 <b>👥 직원·수당</b>, <b>📊 월별 리포트</b> 탭만 보면 됩니다.</p>

<h2>1. 처음 한 번만 — 대리점 이름 바꾸기</h2>
<ol>
<li><b>⚙ 설정</b> 탭을 누릅니다.</li>
<li>아래 '대리점 관리'에 <b>SK 대리점1~3, KT 대리점1~3, LG 대리점1~3</b> 이 미리 만들어져 있어요.</li>
<li>이름 칸을 <b>더블클릭</b> → 실제 대리점 이름으로 고치고 → <b>[💾 대리점 저장]</b>.</li>
<li>대리점이 늘어나면 <b>[➕ 대리점 추가]</b> → 이름 쓰고 통신사 고르고 → 저장.</li>
</ol>

<h2>2. 처음 한 번만 — 매장 설정</h2>
<ol>
<li>같은 ⚙ 설정 탭 '기본 설정'에서 <b>매장 이름</b>(손님 가격표 제목에 나옴)을 씁니다.</li>
<li><b>목표마진</b> = 한 건 팔 때 남기고 싶은 돈. 예: 10만원이면 100,000.</li>
<li><b>[💾 저장]</b>.</li>
</ol>

<h2>3. 처음 한 번만 — 직원 등록</h2>
<ol>
<li><b>👥 직원·수당</b> 탭 → [직원 추가] → 이름, 건당 수당(예: 20000), 마진 비율(예: 10) 입력 → [저장].</li>
<li>수당이 없으면 0으로 두면 됩니다.</li>
</ol>

<h2>4. 처음 한 번만 — AI 키 넣기 (사진·엑셀 자동읽기용)</h2>
<p>대리점 정책 사진/엑셀을 AI가 읽으려면 'AI 키'가 필요합니다. <b>Claude · GPT · Gemini 중 아무거나</b> 이미 쓰는 키를 넣으면 됩니다.
직원마다 다른 키를 써도 되고, 키는 그 PC에만 저장됩니다.</p>
<ul>
<li><b>GPT 키</b>(sk-로 시작): platform.openai.com → API keys</li>
<li><b>Gemini 키</b>(AIza로 시작): aistudio.google.com → Get API key</li>
<li><b>Claude 키</b>(sk-ant-로 시작): 아래 순서</li>
</ul>
<ol>
<li>인터넷 창에서 <b>console.anthropic.com</b> 접속 → 이메일로 가입.</li>
<li>왼쪽 메뉴 <b>Billing(결제)</b> → 카드 등록 후 충전 (예: $10). 정책표 한 장 읽을 때마다 조금씩 차감됩니다.</li>
<li>왼쪽 메뉴 <b>API Keys</b> → <b>Create Key</b> → 이름 아무거나(예: 매장) → 만들어진 키(<b>sk-ant-</b>로 시작)를 <b>복사</b>.<br>
   ※ 키는 그때 한 번만 보여주니 꼭 바로 복사하세요.</li>
<li>이 프로그램 <b>⚙ 설정</b> 탭 → 'AI 키' 칸에 붙여넣기(Ctrl+V) → <b>[연결 테스트]</b> → '정상' 나오면 끝.</li>
</ol>
<p class='box'>💡 AI 키 없이도 프로그램은 다 쓸 수 있어요. 정책을 손으로 입력하거나 엑셀 양식으로 넣으면 됩니다.</p>

<h2>5. 매일 — 대리점 정책 넣기 (가장 중요)</h2>
<h3>① 대리점과 날짜 고르기</h3>
<p><b>📋 정책 입력</b> 탭 맨 위에서 정책을 보낸 <b>대리점</b>을 고르고, 정책이 <b>시작되는 날짜</b>를 고릅니다. (보통 오늘)</p>
<h3>② 파일 올리기 — 셋 중 편한 방법</h3>
<ul>
<li><b>엑셀 파일</b>: 카톡에서 받은 엑셀을 PC에 저장 → 점선 상자에 <b>끌어다 놓기</b> (또는 상자를 눌러 파일 선택)</li>
<li><b>카톡 사진</b>: PC카톡에서 사진 우클릭 → '다른 이름으로 저장' → 끌어다 놓기</li>
<li><b>캡처</b>: 사진을 화면에 크게 띄우고 <b>Win + Shift + S</b> → 표 부분을 드래그 → 프로그램으로 돌아와서 <b>[📋 캡처 붙여넣기]</b> (또는 Ctrl+V)</li>
</ul>
<ul>
<li><b>카톡 글</b>: 대리점이 글로 보낸 정책은 카톡에서 글을 복사(Ctrl+C) → 프로그램에서 <b>[📋 캡처 붙여넣기]</b></li>
<li><b>엑셀 일부만</b>: 엑셀에서 표 부분을 드래그해서 복사 → 똑같이 [📋 붙여넣기]</li>
</ul>
<p class='box'>💡 <b>어떤 대리점 양식이든</b> 됩니다. 엑셀은 AI가 표 구조(어느 칸이 모델·요금제·신규/번이/기변인지)만 파악하고,
금액은 프로그램이 <b>원본 칸에서 그대로</b> 읽어서 틀리지 않습니다. 사진·PDF·카톡 글은 AI가 옮겨 적으니 금액 몇 개만 확인하세요.
엑셀 안에 사진으로 붙여 보낸 정책도 알아서 꺼내서 읽습니다.</p>
<p>여러 장이면 여러 개 올려도 됩니다. 그다음 <b>[🤖 AI로 읽기]</b>를 누르고 기다립니다. (보통 30초~2분)</p>
<h3>③ 확인하고 저장</h3>
<ul>
<li>AI가 채운 표를 훑어봅니다. 특히 <b>리베이트 금액</b>이 맞는지 몇 개만 확인하세요.</li>
<li>틀린 칸은 <b>더블클릭</b>해서 고치고 엔터. 필요 없는 줄은 클릭 후 [선택 행 삭제].</li>
<li><b>[💾 저장]</b>을 누르면 끝. 예전 정책은 지워지지 않고 이력에 남습니다.</li>
<li>'새 정책표에 없는 기존 정책을 종료할까요?' 가 뜨면: 대리점이 준 게 <b>전체 정책표</b>면 '예', <b>일부만</b> 바뀐 거면 '아니오'.</li>
</ul>
<p class='box'>💡 유선(인터넷·TV) 정책도 똑같이 올리면 AI가 '유선'으로 구분합니다. 헷갈리면 위쪽 '종류'를 유선으로 골라주세요.</p>
<h3>④ 부가·차감 규칙 확인 (대리점마다 다른 조건)</h3>
<p>표 위에 <b>💰 정책 금액</b> / <b>📐 부가·차감 규칙</b> 두 칸이 있습니다. AI가 정책표의 조건을 규칙으로 바꿔서 두 번째 칸에 넣어둡니다.</p>
<ul>
<li>예) <b>부가서비스 미유치 4차감</b> → 조건 '부가', 해당 시 0, 미해당 시 -40,000</li>
<li>예) <b>보험 미유치 3차감, 유치 1추가</b> → 조건 '보험', 해당 시 +10,000, 미해당 시 -30,000</li>
<li>예) <b>115군 기변 1차감</b> → 조건 '자동', 해당 시 -10,000, 가입유형 기기변경, 요금 최소 115</li>
<li>예) <b>93일 내 해지 시 환수</b> → 조건 '주의' (금액 계산 안 하고 경고로만 보여줌)</li>
</ul>
<p>'적용 그룹'은 정책 금액 표 맨 오른쪽 '그룹'(프리미엄/중저가 등)과 같은 이름이면 그 모델들에만 적용됩니다.
규칙은 대리점마다 한 번 맞춰두면, 다음 정책 때 규칙이 안 바뀌었으면 그대로 두고 금액만 저장하면 됩니다.</p>

<h2>6. 손님 응대 — 🏆 제일 좋은 대리점 찾기 (가장 많이 쓰게 될 화면)</h2>
<ol>
<li><b>🏆 최적 대리점</b> 탭 → 모델 칸에 <b>S26 256</b> 처럼 입력하거나 목록에서 고릅니다.</li>
<li><b>요금제(천원)</b>: 손님 요금제 월정액. 월 69,000원이면 69. 모르면 '상관없음'.</li>
<li><b>가입유형</b>: 신규 / 번호이동 / 기기변경.</li>
<li><b>손님 조건</b>: 부가서비스·보험·컬러링 가입하면 체크, 이전 폰 13개월 안 됐으면 '단기 기변' 체크.</li>
<li><b>[🏆 제일 좋은 곳 찾기]</b> → 대리점 9곳이 <b>실제로 주는 돈</b> 순서대로 나옵니다. 초록 줄이 1등.</li>
<li>'적용된 조건' 칸에 왜 그 금액인지(예: 보험 미유치 -3만) 나오고, '⚠ 주의' 칸에 환수 조건이 나옵니다.</li>
<li><b>[이 조건으로 개통 등록]</b>을 누르면 개통 등록 화면에 다 채워져서 넘어갑니다. 고객명만 넣고 [등록].</li>
</ol>
<p class='box'>💡 체크를 바꾸면 순위가 바로 다시 계산됩니다. "부가 안 하면 어디가 1등?" 같은 것도 바로 확인 가능.</p>

<h2>6-1. 📶 요금제 구간별 비교</h2>
<ul>
<li>모델과 가입유형을 고르면 <b>115 · 105 · 95 · 85 · 75 · 69 · 61 · 55 · 44 · 33 요금</b>마다 대리점별 실제 받는 돈이 한 표로 나옵니다.</li>
<li>SKT·KT·LG 요금제 이름이 달라도 <b>월 요금 기준</b>으로 맞춰서 비교합니다. 맨 아래 굵은 줄 = 통신사별 최고 금액.</li>
<li>초록 칸 = 그 요금 구간 1등. 칸에 마우스를 올리면 적용된 요금제 구간과 부가·차감 조건이 보입니다.</li>
<li>'요금(천원)' 칸이 <b>자동</b>이면 SK·KT·LG 정책표에 있는 요금 구간을 모아서 칸을 만듭니다. 직접 정하려면 115,95,69 처럼 적기.</li>
</ul>

<h2>6-2. 판매가 보기 / 가격표 뽑기</h2>
<ul>
<li><b>💰 판매가·가격표</b> 탭: 모델 검색하면 손님 할부원금과 우리 마진이 바로 나옵니다.</li>
<li>대리점을 '전체'로 두면 <b>제일 많이 주는 대리점</b> 기준으로 계산해줍니다. (어느 대리점으로 개통할지도 표에 나옴)</li>
<li><b>[🖨 손님용 가격표 만들기]</b> → 인터넷 창이 열리면 <b>Ctrl+P</b> 로 인쇄. (리베이트·마진은 안 나옵니다)</li>
<li><b>⚖ 대리점 비교</b> 탭: 같은 통신사 대리점 3곳 중 어디가 제일 많이 주는지 초록색으로 표시.</li>
</ul>

<h2>7. 개통하면 — 개통 등록</h2>
<ol>
<li><b>📱 개통 등록</b> 탭 → 날짜, 담당 직원, 대리점, 구분(무선/유선) 선택.</li>
<li>모델 → 요금제 → 가입유형을 고르면 <b>정책 금액이 자동으로</b> 들어갑니다.</li>
<li>손님에게 실제로 준 <b>고객지원금</b>(유선은 사은품)을 맞게 고칩니다. 마진이 바로 보여요.</li>
<li>고객명, <b>휴대폰 뒷4자리</b>(정산 맞출 때 씀) 입력 → <b>[등록]</b>.</li>
<li>잘못 넣었으면 아래 목록에서 <b>더블클릭</b> → 고치고 [수정 저장].</li>
</ol>

<h2>8. 돈 들어오면 — 정산 대조</h2>
<ul>
<li><b>🧾 정산 대조</b> 탭 → 월과 대리점 선택.</li>
<li>노란 줄 = 아직 안 들어온 돈. 들어온 건 클릭(여러 개는 Ctrl+클릭) → <b>[✔ 받을 돈 그대로 입금처리]</b>.</li>
<li>금액이 다르게 들어왔으면 '실제 입금' 칸 더블클릭 → 실제 금액 입력 → <b>[💾 고친 금액 저장]</b>. 빨간 줄로 표시돼서 따지기 좋아요.</li>
<li>대리점이 정산 엑셀을 주면 <b>[입금내역 가져오기]</b>로 한 번에 맞춰집니다.</li>
</ul>

<h2>9. 월말</h2>
<ul>
<li><b>👥 직원·수당</b>: 월 고르면 직원별 건수와 수당이 계산됩니다.</li>
<li><b>📊 월별 리포트</b>: 아래에 임대료·월급 등 경비를 넣으면 월별 <b>순이익</b>이 나옵니다. [지난달 경비 복사]로 매달 편하게.</li>
<li>모든 표는 <b>[엑셀로 내보내기]</b>로 엑셀 파일로 저장할 수 있어요.</li>
</ul>

<h2>11. 직원 PC에 나눠주기 · 자동 업데이트</h2>
<h3>🧙 설치 마법사 Setup.exe (카카오톡과 똑같은 방식)</h3>
<ol>
<li>사장님 PC: ⚙ 설정 → 맨 아래 <b>[🧙 설치 마법사 Setup.exe 만들기]</b> → 10분 정도 기다리면 바탕화면에 <b>폰정책매니저-Setup.exe</b>.</li>
<li>직원은 더블클릭 → <b>다음 → 설치 → 마침</b>. 바탕화면·시작 메뉴 아이콘, 제어판 '앱 제거' 등록. 파이썬 필요 없음.</li>
<li>Setup.exe는 처음 한 번만 만들면 됩니다. 이후 업데이트는 공유 폴더로 자동 전달됩니다.</li>
</ol>
<h3>⭐ 간단한 방법: 설치 bat</h3>
<ol>
<li>사장님 PC: ⚙ 설정 → 직원 PC 공유에서 공유 폴더를 먼저 정하고 → 맨 아래 <b>[💿 직원용 설치파일 만들기]</b> (바로 만들어짐).</li>
<li>바탕화면에 <b>폰정책매니저_설치.bat</b> 이 생깁니다. (공유 폴더에도 자동 복사)</li>
<li>이 파일을 카톡·USB로 직원에게 주면, 직원은 <b>더블클릭</b>만 하면 설치 끝. 파이썬이 없으면 자동으로 깔고, 바탕화면에 '폰정책매니저' 아이콘이 생깁니다.</li>
<li>처음 켤 때 "직원 PC / 사장님 PC"를 고르고 이름만 적으면 공유 폴더에서 정책을 자동으로 받아옵니다.</li>
<li>파란 창 "Windows의 PC 보호"가 뜨면 <b>[추가 정보] → [실행]</b>. (개인이 만든 프로그램이라 뜨는 정상 안내입니다)</li>
</ol>
<p class='box'>💡 설치한 뒤의 업데이트는 설치파일을 다시 안 만들어도 됩니다. 사장님이 새 프로그램 파일로 바꾸고 켜기만 하면 직원 PC에 업데이트 창이 뜹니다.</p>
<h3>직접 설정하는 방법</h3>
<p>사장님 PC가 <b>공유 폴더</b>에 최신 프로그램·정책을 올려두고, 직원 PC가 그걸 자동으로 받아가는 방식입니다.</p>
<h3>공유 폴더 만들기 (둘 중 하나)</h3>
<ul>
<li><b>매장이 여러 곳이거나 집에서도 볼 때</b>: 구글 드라이브 PC용(Google Drive for desktop)을 모든 PC에 설치 →
사장님이 드라이브에 '폰정책공유' 폴더를 만들어 직원 계정에 <b>편집자</b>로 공유 → 각 PC 탐색기에 G: 드라이브로 보입니다.</li>
<li><b>한 매장 안 PC끼리만</b>: 사장님 PC에 '폰정책공유' 폴더를 만들고 우클릭 → 속성 → 공유 → 직원이 쓸 수 있게 공유.</li>
</ul>
<h3>사장님 PC (한 번만)</h3>
<ol>
<li>⚙ 설정 → '직원 PC 공유' → 역할 <b>사장님 PC</b>, 공유 폴더에 위 폴더 선택 → <b>[공유 설정 저장]</b>.</li>
<li>이제 정책·대리점·직원을 저장할 때마다 공유 폴더로 자동 전달됩니다.</li>
</ol>
<h3>직원 PC (한 번만)</h3>
<ol>
<li>Python 설치 후, 공유 폴더의 <b>phone_policy_manager.py</b>를 직원 PC <b>바탕화면 새 폴더에 복사</b>해서 실행.
(공유 폴더 안에서 바로 실행하지 마세요)</li>
<li>⚙ 설정 → 역할 <b>직원 PC</b>, PC 이름(예: 김민수PC), 같은 공유 폴더 선택 → <b>[공유 설정 저장]</b>.</li>
<li>대리점·정책·직원 목록이 바로 들어옵니다. 직원 PC는 정책을 고칠 수 없고, 받기만 합니다.</li>
</ol>
<h3>업데이트할 때 (사장님)</h3>
<ol>
<li>새 프로그램 파일을 받으면 사장님 PC의 기존 파일에 <b>덮어쓰고 실행</b>만 하면 됩니다.</li>
<li>켜지면서 공유 폴더에 새 버전이 자동으로 올라가고, 직원 PC에는 <b>"새 버전이 있습니다. 업데이트할까요?"</b> 창이 뜹니다.</li>
<li>직원이 [예]를 누르면 자동으로 바뀌고 다시 켜집니다. 직원이 입력한 개통 내역은 그대로 남습니다.</li>
</ol>
<p class='box'>💡 직원 PC에서 등록한 개통 건은 사장님 PC의 정산 대조·직원 수당·월별 리포트에 자동으로 모입니다.
입금 처리는 사장님 PC에서만 합니다.</p>

<h2>12. 인터넷 자동 업데이트 (GitHub) 설정 — 처음 한 번만</h2>
<h3>① 저장소 만들기</h3>
<ol>
<li>github.com 로그인 → 오른쪽 위 <b>+</b> → <b>New repository</b></li>
<li>Repository name: 예) <b>phone-policy</b> → <b>Public</b> 선택 → 맨 아래 <b>Create repository</b></li>
</ol>
<h3>② 토큰(배포용 열쇠) 만들기</h3>
<ol>
<li>오른쪽 위 프로필 사진 → <b>Settings</b> → 왼쪽 맨 아래 <b>Developer settings</b></li>
<li><b>Personal access tokens → Fine-grained tokens → Generate new token</b></li>
<li>Token name: 아무거나(예: 폰정책배포) / Expiration: <b>No expiration</b>(또는 1년)</li>
<li>Repository access: <b>Only select repositories</b> → 방금 만든 저장소 선택</li>
<li>Permissions → Repository permissions → <b>Contents</b> 를 <b>Read and write</b> 로</li>
<li>맨 아래 <b>Generate token</b> → 나온 토큰(github_pat_…) <b>복사</b> (한 번만 보여줌)</li>
</ol>
<h3>③ 프로그램에 넣기 (사장님 PC)</h3>
<ol>
<li>⚙ 설정 → 🌐 인터넷 자동 업데이트 → 저장소: <b>GitHub아이디/phone-policy</b>, 토큰 붙여넣기 → <b>[💾 저장 + 연결 확인]</b></li>
<li><b>[📤 직원들에게 업데이트 배포]</b> 한 번 누르기</li>
<li>그다음 <b>설치파일(Setup.exe 또는 설치 bat)을 새로 만들어서</b> 직원 PC에 한 번만 다시 설치 → 이후로는 자동</li>
</ol>
<h3>업데이트할 때마다</h3>
<p>새 파일로 사장님 PC 업데이트 → 설정 탭 <b>[📤 직원들에게 업데이트 배포]</b> → 직원 PC에 "업데이트할까요?" 창 → [예]. 끝.</p>

<h2>10. 자주 묻는 질문</h2>
<p><b>Q. AI가 틀리게 읽으면?</b><br>표에서 칸을 더블클릭해서 고치면 됩니다. 사진이 흐리면 틀리기 쉬워서, 가능하면 <b>엑셀 원본</b>을 올리는 게 제일 정확해요.</p>
<p><b>Q. AI 비용은 얼마나?</b><br>쓴 만큼만 충전금에서 빠집니다. 보통 정책표 한 장에 몇백 원 수준이지만 표 크기에 따라 다르니, console.anthropic.com 의 Usage 에서 확인하세요.</p>
<p><b>Q. 인터넷 끊기면?</b><br>AI 읽기만 안 되고 나머지는 전부 됩니다.</p>
<p><b>Q. 데이터 날아가면?</b><br>⚙ 설정 → [백업 파일 만들기]를 <b>일주일에 한 번</b> 눌러서 USB나 메일에 보관하세요.</p>
<p><b>Q. 컴퓨터 바꾸면?</b><br>프로그램 파일과 phone_policy.db 파일을 같이 새 컴퓨터 같은 폴더에 복사하면 그대로 씁니다.</p>
"""


class GuideTab(QWidget):
    def __init__(self, db):
        super().__init__()
        lay = QVBoxLayout(self)
        tb = QTextBrowser()
        tb.setOpenExternalLinks(True)
        tb.setHtml(GUIDE_HTML)
        lay.addWidget(tb)


# ============================================================
# 설치파일 만들기 (사장님 PC, 파이썬 실행 상태에서만)
# ============================================================
class BuildWorker(QThread):
    line = Signal(str)
    finished_ok = Signal(str)
    failed = Signal(str)

    def __init__(self, share_dir, store="", source=""):
        super().__init__()
        self.share_dir = share_dir
        self.store = store
        self.source = source

    @staticmethod
    def py():
        p = os.path.join(os.path.dirname(sys.executable), "python.exe")
        return p if os.path.isfile(p) else sys.executable

    def _run(self, args, cwd):
        p = subprocess.Popen(args, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        for raw in p.stdout:
            self.line.emit(raw.decode("utf-8", "ignore").rstrip())
        if p.wait() != 0:
            raise RuntimeError(f"명령 실패: {' '.join(args[:4])} …")

    def _find_iscc(self):
        cands = [os.path.join(_LOCAL, "Programs", "Inno Setup 6", "ISCC.exe"),
                 os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "Inno Setup 6", "ISCC.exe"),
                 os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "Inno Setup 6", "ISCC.exe")]
        return next((c for c in cands if os.path.isfile(c)), None)

    def _get_inno(self, work):
        iscc = self._find_iscc()
        if iscc:
            return iscc
        self.line.emit("③ 설치 화면 제작 도구(Inno Setup) 받는 중…")
        setup = os.path.join(work, "innosetup.exe")
        try:
            req = urllib.request.Request("https://jrsoftware.org/download.php/is.exe",
                                         headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=180) as r, open(setup, "wb") as f:
                shutil.copyfileobj(r, f)
            self._run([setup, "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-", "/CURRENTUSER"], work)
        except Exception as ex:
            self.line.emit(f"직접 받기 실패({ex}) → winget으로 시도")
            try:
                self._run(["winget", "install", "-e", "--id", "JRSoftware.InnoSetup", "--silent",
                           "--accept-package-agreements", "--accept-source-agreements"], work)
            except Exception:
                pass
        iscc = self._find_iscc()
        if not iscc:
            raise RuntimeError("Inno Setup을 설치하지 못했습니다. https://jrsoftware.org/isdl.php 에서 "
                               "innosetup 설치 후 다시 눌러주세요.")
        return iscc

    def _iss(self, work, iscc):
        kor = os.path.join(os.path.dirname(iscc), "Languages", "Korean.isl")
        lang = 'Name: "ko"; MessagesFile: "compiler:Languages\\Korean.isl"' if os.path.isfile(kor) \
            else 'Name: "ko"; MessagesFile: "compiler:Default.isl"'
        msgs = "" if os.path.isfile(kor) else """
[Messages]
SetupWindowTitle=%1 설치
WelcomeLabel1=[name] 설치를 시작합니다
WelcomeLabel2=이 프로그램은 [name/ver]을(를) 이 컴퓨터에 설치합니다.%n%n"다음"을 눌러 계속하세요.
ButtonBack=< 뒤로
ButtonNext=다음 >
ButtonInstall=설치
ButtonCancel=취소
ButtonFinish=마침
WizardSelectTasks=추가 작업
SelectTasksDesc=어떤 작업을 함께 할까요?
SelectTasksLabel2=원하는 항목을 고르고 "다음"을 누르세요.
WizardReady=설치 준비 완료
ReadyLabel1=[name]을(를) 설치할 준비가 끝났습니다.
ReadyLabel2a="설치"를 누르면 설치를 시작합니다.
WizardInstalling=설치 중
InstallingLabel=[name]을(를) 설치하는 중입니다. 잠시만 기다려 주세요.
FinishedHeadingLabel=[name] 설치 완료
FinishedLabel=설치가 끝났습니다. 바탕화면의 [name] 아이콘으로 실행할 수 있습니다.
ExitSetupTitle=설치 취소
ExitSetupMessage=설치가 끝나지 않았습니다.%n%n지금 나가면 설치되지 않습니다. 나갈까요?
"""
        store = (self.store or "우리매장").replace('"', "")
        iss = f"""[Setup]
AppId={{{{8C3F2A5E-4B7D-4E21-9A6B-3D5F7A9C1E24}}
AppName=폰정책매니저
AppVersion={APP_VERSION}
AppVerName=폰정책매니저 {APP_VERSION}
AppPublisher={store}
DefaultDirName={{localappdata}}\\Programs\\PhonePolicyManager
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir={os.path.join(work, "setup")}
OutputBaseFilename=폰정책매니저-Setup
SetupIconFile={os.path.join(work, "app.ico")}
UninstallDisplayIcon={{app}}\\{EXE_NAME}
UninstallDisplayName=폰정책매니저
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes

[Languages]
{lang}
{msgs}
[Tasks]
Name: "desktopicon"; Description: "바탕화면에 폰정책매니저 아이콘 만들기"

[Files]
Source: "{os.path.join(work, "dist", EXE_NAME[:-4])}\\*"; DestDir: "{{app}}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{{autoprograms}}\\폰정책매니저"; Filename: "{{app}}\\{EXE_NAME}"
Name: "{{autodesktop}}\\폰정책매니저"; Filename: "{{app}}\\{EXE_NAME}"; Tasks: desktopicon

[Run]
Filename: "{{app}}\\{EXE_NAME}"; Description: "지금 폰정책매니저 실행"; Flags: nowait postinstall skipifsilent
"""
        path = os.path.join(work, "setup.iss")
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(iss)
        return path

    def run(self):
        try:
            work = os.path.join(BASE_DIR, "_build")
            shutil.rmtree(work, ignore_errors=True)
            os.makedirs(work)
            src = os.path.join(work, "ppm_app.py")
            with open(src, "w", encoding="utf-8") as f:
                f.write(self.source)
            write_ico(os.path.join(work, "app.ico"))
            self.line.emit("① 필요한 도구 설치 중 (PyInstaller)…")
            self._run([self.py(), "-m", "pip", "install", "--upgrade", "pyinstaller", "PySide6", "openpyxl", "xlrd"],
                      work)
            self.line.emit("② 프로그램 묶는 중… (3~8분, 창을 닫지 마세요)")
            defaults = os.path.join(work, "ppm_defaults.json")
            with open(defaults, "w", encoding="utf-8") as f:
                json.dump({"share_dir": self.share_dir}, f, ensure_ascii=False)
            self._run([self.py(), "-m", "PyInstaller", "--noconfirm", "--onedir", "--windowed",
                       "--name", EXE_NAME[:-4], "--icon", os.path.join(work, "app.ico"),
                       "--add-data", f"{src}{os.pathsep}.", "--add-data", f"{defaults}{os.pathsep}.",
                       "--hidden-import", "openpyxl", "--hidden-import", "xlrd",
                       "--distpath", os.path.join(work, "dist"), "--workpath", os.path.join(work, "build"),
                       "--specpath", work, src], work)
            if not os.path.isfile(os.path.join(work, "dist", EXE_NAME[:-4], EXE_NAME)):
                raise RuntimeError("프로그램 묶기에 실패했습니다.")
            iscc = self._get_inno(work)
            self.line.emit("④ 설치파일(설치 화면 포함) 만드는 중…")
            self._run([iscc, self._iss(work, iscc)], work)
            exe = os.path.join(work, "setup", SETUP_NAME)
            if not os.path.isfile(exe):
                raise RuntimeError("설치파일이 만들어지지 않았습니다.")
            desk = os.path.join(os.path.expanduser("~"), "Desktop")
            try:
                out = subprocess.run(["powershell", "-NoProfile", "-Command", "[Environment]::GetFolderPath('Desktop')"],
                                     capture_output=True, text=True,
                                     creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout.strip()
                if out and os.path.isdir(out):
                    desk = out
            except Exception:
                pass
            dst = os.path.join(desk, SETUP_NAME)
            shutil.copyfile(exe, dst)
            if self.share_dir and os.path.isdir(self.share_dir):
                shutil.copyfile(exe, os.path.join(self.share_dir, SETUP_NAME))
                self.line.emit("공유 폴더에도 복사했습니다.")
            self.finished_ok.emit(dst)
        except Exception as ex:
            self.failed.emit(str(ex))


class BuildDialog(QDialog):
    def __init__(self, parent, db, auto_install=False):
        super().__init__(parent)
        self.auto_install = auto_install
        self.setup_path = None
        self.setWindowTitle("직원용 설치파일 만들기")
        self.resize(760, 460)
        lay = QVBoxLayout(self)
        self.status = QLabel("준비 중…")
        self.status.setStyleSheet("font-weight:bold;font-size:12pt")
        lay.addWidget(self.status)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        lay.addWidget(self.log)
        self.b_close = btn("닫기", self.accept)
        self.b_close.setEnabled(False)
        lay.addWidget(self.b_close)
        self.w = BuildWorker(db.get("share_dir", ""), db.get("store_name", ""), publish_source(db))
        self.w.line.connect(self._line)
        self.w.finished_ok.connect(self._ok)
        self.w.failed.connect(self._fail)
        self.w.start()

    def _line(self, t):
        if t.startswith(("①", "②", "③", "④", "공유")):
            self.status.setText(t)
        self.log.appendPlainText(t)

    def _ok(self, path):
        self.b_close.setEnabled(True)
        self.status.setText("✅ 완성!")
        self.setup_path = path
        if self.auto_install:
            info(self, "✅ 설치파일을 만들었습니다. 이제 설치 화면이 뜹니다.\n'다음 → 설치 → 마침'을 누르세요.\n\n"
                       f"직원에게 줄 설치파일은 바탕화면에 있습니다:\n{path}")
            self.accept()
            return
        info(self, f"✅ 설치파일을 만들었습니다.\n{path}\n\n"
                   "카카오톡 설치파일처럼 더블클릭하면 설치 화면(다음 → 설치 → 마침)이 뜨고,\n"
                   "바탕화면 아이콘이 생기고, 제어판 '앱 제거'에서 지울 수도 있습니다.\n\n"
                   "이 파일을 카톡·USB·공유 폴더로 직원에게 주면 됩니다.\n"
                   "(처음 실행 때 '알 수 없는 게시자' 파란 창이 뜨면 [추가 정보] → [실행])")

    def _fail(self, msg):
        self.b_close.setEnabled(True)
        self.status.setText("❌ 실패")
        warn(self, f"설치파일을 만들지 못했습니다:\n{msg}\n\n아래 기록 창 내용을 캡처해서 보내주세요.")


# ============================================================
# 설치형(exe) 시작 처리: 설치 · 바로가기 · 최신 본체 실행
# ============================================================
def _ps(cmd):
    return subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
                          capture_output=True, text=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def make_shortcuts():
    for folder in ("Desktop", "Programs"):
        _ps(f"$d=[Environment]::GetFolderPath('{folder}');"
            f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d '{SHORTCUT_NAME}.lnk'));"
            f"$s.TargetPath='{INSTALLED_EXE}';$s.WorkingDirectory='{INSTALL_DIR}';$s.Save()")


def bundled_file(name):
    return os.path.join(getattr(sys, "_MEIPASS", BASE_DIR), name)


def frozen_install_if_needed():
    """설치 안 된 곳(다운로드 폴더 등)에서 실행됐으면 설치하고 True(=이번 실행은 종료)"""
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    if os.path.normcase(os.path.abspath(sys.executable)) == os.path.normcase(INSTALLED_EXE) \
            or glob.glob(os.path.join(exe_dir, "unins*.exe")):
        return False
    app = QApplication.instance() or QApplication(sys.argv)
    app.setFont(QFont("Malgun Gothic", 10))
    app.setWindowIcon(app_icon())
    again = os.path.isfile(INSTALLED_EXE)
    q = ("폰정책매니저를 새 버전으로 다시 설치할까요?" if again else "폰정책매니저를 이 컴퓨터에 설치할까요?")
    if QMessageBox.question(None, APP_NAME, q + "\n\n바탕화면에 바로가기가 만들어집니다. (입력한 데이터는 지워지지 않아요)") \
            != QMessageBox.StandardButton.Yes:
        return True
    try:
        os.makedirs(INSTALL_DIR, exist_ok=True)
        shutil.copyfile(sys.executable, INSTALLED_EXE + ".new")
        os.replace(INSTALLED_EXE + ".new", INSTALLED_EXE)
        # 설치파일에 들어있는 본체가 더 새것이면 본체도 교체
        b = bundled_file("ppm_app.py")
        if os.path.isfile(b) and ver_tuple(read_version(b)) >= ver_tuple(read_version(APP_PY) or "0"):
            shutil.copyfile(b, APP_PY)
        make_shortcuts()
    except PermissionError:
        QMessageBox.warning(None, APP_NAME, "폰정책매니저가 켜져 있어서 설치할 수 없습니다.\n프로그램을 닫고 다시 실행해 주세요.")
        return True
    except Exception as ex:
        QMessageBox.warning(None, APP_NAME, f"설치 중 오류: {ex}")
        return True
    QMessageBox.information(None, APP_NAME, "✅ 설치가 끝났습니다!\n바탕화면의 '폰정책매니저' 아이콘으로 실행하세요.\n지금 바로 켤게요.")
    subprocess.Popen([INSTALLED_EXE], cwd=INSTALL_DIR)
    return True


def frozen_prepare_body():
    """설치형: 데이터 폴더의 본체(.py)를 최신으로 맞춤. 더 최신 본체가 있으면 그걸로 실행해야 하므로 경로 반환"""
    b = bundled_file("ppm_app.py")
    if os.path.isfile(b) and ver_tuple(read_version(b)) > ver_tuple(read_version(APP_PY) or "0"):
        shutil.copyfile(b, APP_PY)
    v = read_version(APP_PY)
    if v and ver_tuple(v) > ver_tuple(APP_VERSION):
        return APP_PY
    return None


def first_run_setup(db):
    """설치형 첫 실행: 공유 폴더 기본값 적용, 사장님/직원 PC 선택, 예전 데이터 가져오기"""
    if db.get("first_setup_done", "") == "1":
        return
    db.set("first_setup_done", "1")
    try:
        with open(bundled_file("ppm_defaults.json"), encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    has_data = db.one("SELECT (SELECT COUNT(*) FROM policies) + (SELECT COUNT(*) FROM sales) AS n")["n"] > 0
    if d.get("share_dir") and not db.get("share_dir", ""):
        db.set("share_dir", d["share_dir"])
    if d.get("share_dir") and not db.get("role", "") and DEFAULT_ROLE != "직원":
        box = QMessageBox()
        box.setWindowTitle(APP_NAME)
        box.setText("이 컴퓨터는 누가 쓰나요?")
        b_staff = box.addButton("직원 PC", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("사장님 PC", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        is_staff = box.clickedButton() == b_staff
        db.set("role", "직원" if is_staff else "관리자")
        if is_staff:
            name, ok = QInputDialog.getText(None, APP_NAME, "이 PC 이름(직원 이름)을 적어주세요. 예: 김민수",
                                            text=socket.gethostname())
            if ok and name.strip():
                db.set("pc_name", name.strip())
            if not os.path.isdir(d["share_dir"]):
                QMessageBox.information(None, APP_NAME, f"공유 폴더({d['share_dir']})가 이 PC에서 안 보입니다.\n"
                                                        "구글 드라이브 설치 후 설정 탭에서 공유 폴더를 다시 골라주세요.")
            return
    if has_data or db.is_staff_pc():
        return
    if ask(None, "예전에 쓰던 프로그램 데이터(phone_policy.db)가 있으면 가져올까요?\n(처음 쓰시면 '아니오')"):
        path, _ = QFileDialog.getOpenFileName(None, "예전 데이터 파일 선택", os.path.expanduser("~"),
                                              "데이터 파일 (phone_policy.db *.db)")
        if path:
            db.con.close()
            shutil.copyfile(path, DB_PATH)
            return "reload"


# ============================================================
# 프로그램 아이콘 (파일 안에 내장)
# ============================================================
ICON_PNG_B64 = (
    "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"
)
ICON_ICO_B64 = (
    "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"
)


def app_icon():
    pm = QPixmap()
    pm.loadFromData(base64.b64decode(ICON_PNG_B64), "PNG")
    return QIcon(pm)


def write_ico(path):
    with open(path, "wb") as f:
        f.write(base64.b64decode(ICON_ICO_B64))
    return path


# ============================================================
# 메인
# ============================================================
class MainWindow(QMainWindow):
    def __init__(self, db, quiet=False):
        super().__init__()
        self.quiet = quiet
        self.db = db
        self.share = Share(db)
        db.after_change = self._on_data_change
        self._asked_version = None
        self._title()
        self.resize(1460, 900)
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.pages = [
            ("❓ 사용법", GuideTab(db)),
            ("⭐ 한눈에", DashboardTab(db)),
            ("⏰ 알림", AlertTab(db)),
            ("🔎 정책마진 조회", MarginTab(db)),
            ("📥 정책 일괄 등록", BatchTab(db)),
            ("📣 판매정책", StorePolicyTab(db)),
            ("🧾 견적서", QuoteTab(db)),
            ("📦 모델 관리", ModelsTab(db)),
            ("📋 정책 입력(개별)", PolicyEditTab(db)),
            ("🏆 최적 대리점", BestTab(db)),
            ("📶 구간별 비교", TierTab(db)),
            ("💰 판매가·가격표", PriceTab(db)),
            ("⚖ 대리점 비교", CompareTab(db)),
            ("🕘 정책 이력", HistoryTab(db)),
            ("📱 개통 등록", SalesTab(db)),
            ("🧾 정산 대조", SettleTab(db)),
            ("👥 직원·수당", StaffTab(db)),
            ("📊 월별 리포트", ReportTab(db)),
            ("⚙ 설정", SettingsTab(db)),
        ]
        for name, w in self.pages:
            self.tabs.addTab(w, name)
        pages = dict(self.pages)
        sales = pages["📱 개통 등록"]

        def go_register(d):
            self.tabs.setCurrentWidget(sales)
            sales.prefill(d)
        pages["🏆 최적 대리점"].on_register = go_register
        pages["⚙ 설정"].main = self
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(lambda: self.share_sync(startup=False))
        if not quiet:
            self.sync_timer.start(10 * 60 * 1000)
            QTimer.singleShot(1500, lambda: self.share_sync(startup=True))
            QTimer.singleShot(4000, self.daily_jobs)
        self.tabs.currentChanged.connect(self.on_tab)
        first = db.get("first_run_done", "") != "1" and not quiet
        self.tabs.setCurrentIndex(0 if first else 1)
        self.on_tab(self.tabs.currentIndex())
        if first:
            db.set("first_run_done", "1")
            QTimer.singleShot(400, self.welcome)

    def daily_jobs(self):
        try:
            auto_backup(self.db)
        except Exception as ex:
            self.statusBar().showMessage(f"자동 백업 실패: {ex}", 10000)
        try:
            at = dict(self.pages)["⏰ 알림"]
            n_today, n_rev = at.refresh()
            i = self.tabs.indexOf(at)
            self.tabs.setTabText(i, f"⏰ 알림 ({n_today + n_rev})" if n_today + n_rev else "⏰ 알림")
            if (n_today or n_rev) and self.db.get("alert_shown", "") != today():
                self.db.set("alert_shown", today())
                if ask(self, f"⏰ 오늘 챙길 손님이 있습니다.\n\n· 부가서비스 해지·요금제 변경 안내: {n_today}명\n"
                             f"· 할부 만료 재방문 대상: {n_rev}명\n\n알림 화면을 열까요?"):
                    self.tabs.setCurrentWidget(at)
        except Exception:
            pass

    def closeEvent(self, e):
        try:
            auto_backup(self.db, force=True)
        except Exception:
            pass
        super().closeEvent(e)

    def _title(self):
        role = "직원 PC" if self.db.is_staff_pc() else "사장님 PC"
        shared = " · 공유 켜짐" if self.db.get("share_dir", "") else ""
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}  [{role}{shared}]")

    def _on_data_change(self, what):
        try:
            if self.share.staff():
                if what == "sales":
                    self.share.publish_sales()
            else:
                if what in ("policy", "agency", "staff", "settings"):
                    self.share.publish_policies()
        except Exception as ex:
            self.statusBar().showMessage(f"공유 폴더에 쓰지 못했습니다: {ex}", 10000)

    def github_check(self, loud=False):
        """인터넷(GitHub)에 새 버전이 있으면 업데이트 창 → [예] 누르면 받아서 설치"""
        repo = gh_repo(self.db)
        if not repo:
            if loud:
                warn(self, "인터넷 업데이트 저장소가 정해지지 않았습니다. (설정 탭 🌐 인터넷 자동 업데이트)")
            return False
        try:
            v = json.loads(gh_get_file(repo, "version.json", timeout=8).decode("utf-8"))
        except Exception as ex:
            if loud:
                if isinstance(ex, urllib.error.HTTPError) and ex.code == 404:
                    warn(self, "아직 배포된 버전이 없습니다.\n사장님 PC에서 [📤 직원들에게 업데이트 배포]를 먼저 눌러주세요.")
                else:
                    warn(self, "새 버전 확인 실패\n\n" + gh_error(ex))
            return False
        rv = str(v.get("version", ""))
        if not rv or ver_tuple(rv) <= ver_tuple(APP_VERSION):
            if loud:
                info(self, f"✅ 최신 버전입니다. ({APP_VERSION})")
            return False
        if not loud and self._asked_version == rv:
            return False
        self._asked_version = rv
        note = f"\n\n바뀐 내용: {v['notes']}" if v.get("notes") else ""
        if not ask(self, f"🔔 새 버전이 있습니다!\n\n지금 버전: {APP_VERSION}\n새 버전: {rv}{note}\n\n"
                         "지금 업데이트할까요? (자동으로 받아서 다시 켜집니다. 입력한 데이터는 그대로 남아요)"):
            return False
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            data = gh_get_file(repo, "phone_policy_manager.py", timeout=60)
            if v.get("sha256") and hashlib.sha256(data).hexdigest() != v["sha256"]:
                raise RuntimeError("받은 파일이 올바르지 않습니다. 잠시 뒤 다시 시도해 주세요.")
            tmp = os.path.join(BASE_DIR, "update_download.py")
            with open(tmp, "wb") as f:
                f.write(data)
        except Exception as ex:
            QApplication.restoreOverrideCursor()
            warn(self, "업데이트 받기 실패\n\n" + gh_error(ex))
            return False
        QApplication.restoreOverrideCursor()
        self.do_update(tmp)
        return True

    def share_sync(self, startup=False, loud=False):
        """켤 때 / 10분마다: 업데이트 확인, 정책 받기(직원), 개통 모으기(사장님)"""
        self._title()
        if not loud and self.github_check():
            return
        sh = self.share
        if not sh.folder():
            if loud and not gh_repo(self.db):
                warn(self, "공유 폴더가 정해지지 않았거나 찾을 수 없습니다.\n설정 탭 '직원 PC 공유'에서 폴더를 정해주세요.")
            return
        msgs = []
        try:
            path, rv = sh.remote_program()
            if sh.staff():
                if rv and ver_tuple(rv) > ver_tuple(APP_VERSION) and (loud or self._asked_version != rv):
                    self._asked_version = rv
                    if ask(self, f"🔔 사장님이 새 버전 프로그램을 올렸습니다.\n\n지금 버전: {APP_VERSION}\n새 버전: {rv}\n\n"
                                 "지금 업데이트할까요? (자동으로 다시 켜집니다. 입력한 데이터는 그대로 남아요)"):
                        return self.do_update(path)
                got = sh.import_policies(force=loud)
                if got:
                    msgs.append(f"사장님이 올린 최신 정책을 받았습니다 ({got})")
                    w = self.tabs.currentWidget()
                    if hasattr(w, "refresh"):
                        w.refresh()
                sh.publish_sales()
            else:
                if not rv or ver_tuple(APP_VERSION) > ver_tuple(rv):
                    sh.publish_program()
                    if rv:
                        msgs.append(f"직원 PC에 새 버전 {APP_VERSION} 배포 완료")
                if startup:
                    sh.publish_policies()
                n = sh.collect_sales()
                if n:
                    msgs.append(f"직원 PC 개통 {n}건을 새로 모았습니다")
        except Exception as ex:
            msgs.append(f"공유 폴더 동기화 오류: {ex}")
        if msgs:
            self.statusBar().showMessage("  ·  ".join(msgs), 15000)
        if loud:
            info(self, "\n".join(msgs) if msgs else "✅ 최신 상태입니다.")

    def do_update(self, path):
        try:
            me = self.share.install_update(path)
        except Exception as ex:
            return warn(self, f"업데이트를 설치하지 못했습니다:\n{ex}\n\n(프로그램은 기존 버전 그대로입니다)")
        info(self, "✅ 업데이트했습니다. 프로그램을 다시 켭니다.")
        try:
            self.db.con.close()
        except Exception:
            pass
        if FROZEN:
            subprocess.Popen([sys.executable], cwd=INSTALL_DIR)
        else:
            subprocess.Popen([sys.executable, me], cwd=os.path.dirname(me))
        QApplication.quit()

    def on_tab(self, i):
        w = self.tabs.widget(i)
        if not self.share.staff() and isinstance(w, (SettleTab, ReportTab, StaffTab)) and self.share.folder():
            try:
                self.share.collect_sales()
            except Exception:
                pass
        if hasattr(w, "refresh"):
            try:
                w.refresh()
            except Exception as ex:
                if getattr(self, "quiet", False):
                    raise
                warn(self, f"화면을 불러오는 중 오류가 났습니다:\n{ex}")

    def welcome(self):
        info(self, "환영합니다! 👋\n\n지금 보이는 '❓ 사용법' 탭을 한 번만 쭉 읽어주세요.\n\n"
                   "처음 할 일:\n 1) ⚙ 설정 → 대리점 이름 바꾸기, 매장 이름, AI 키\n 2) 👥 직원·수당 → 직원 등록\n"
                   " 3) 📋 정책 입력 → 대리점 정책 올리기")


SETUP_BAT_NAME = "폰정책매니저_설치.bat"

_BAT_HEAD = r"""@echo off
chcp 65001 >nul
title 폰정책매니저 설치
setlocal EnableExtensions
set "APPDIR=%LOCALAPPDATA%\PhonePolicyManager\app"
set "APPPY=%APPDIR%\phone_policy_manager.py"
echo.
echo  ==========================================
echo      폰정책매니저 설치를 시작합니다
echo  ==========================================
echo.
if not exist "%APPDIR%" mkdir "%APPDIR%"

set "PYEXE="
for %%V in (3.13 3.12 3.11 3.14) do call :trypy %%V
if not defined PYEXE call :trypython
if defined PYEXE goto :havepy

echo  [1/3] 파이썬이 없어서 설치합니다. 처음 한 번만 2~5분 걸립니다...
winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements >nul 2>&1
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if defined PYEXE goto :havepy
echo  인터넷에서 파이썬을 받는 중...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe' -OutFile $env:TEMP\ppm_python.exe"
if exist "%TEMP%\ppm_python.exe" "%TEMP%\ppm_python.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if defined PYEXE goto :havepy
echo.
echo  파이썬을 자동으로 설치하지 못했습니다.
echo  https://www.python.org 에서 설치한 뒤 이 파일을 다시 더블클릭해 주세요.
pause
exit /b 1

:havepy
echo  [2/3] 프로그램 설치 중...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$m='#__PPM_'+'PAYLOAD__';$t=[IO.File]::ReadAllText('%~f0',[Text.Encoding]::UTF8);$i=$t.LastIndexOf($m);$j=$t.IndexOf([char]10,$i);[IO.File]::WriteAllText($env:APPPY,$t.Substring($j+1),(New-Object Text.UTF8Encoding $false))"
if not exist "%APPPY%" goto :copyfail
echo  [3/3] 필요한 구성요소 설치 중... 처음엔 1~3분 걸립니다
"%PYEXE%" -m pip install --disable-pip-version-check -q PySide6 openpyxl xlrd
if errorlevel 1 goto :pipfail
echo  점검 중...
"%PYEXE%" "%APPPY%" --selftest > "%TEMP%\ppm_selftest.txt" 2>&1
findstr /c:"SELFTEST_OK" "%TEMP%\ppm_selftest.txt" >nul
if errorlevel 1 goto :testfail
"%PYEXE%" "%APPPY%" --install-shortcut
echo.
echo  ------------------------------------------------------------
echo   직원 PC에 나눠줄 [폰정책매니저-Setup.exe] 도 지금 만들까요?
echo   카카오톡처럼 설치 화면이 뜨고, 파이썬 없는 PC에서도 됩니다.
echo   만드는 데 10분 정도 걸립니다. 다 되면 바탕화면에 생깁니다.
echo  ------------------------------------------------------------
choice /c YN /n /m "  만들려면 Y, 나중에 하려면 N 을 누르세요: "
if errorlevel 2 goto :skipbuild
"%PYEXE%" "%APPPY%" --build-setup
:skipbuild
for %%D in ("%PYEXE%") do set "PYW=%%~dpDpythonw.exe"
if not exist "%PYW%" set "PYW=%PYEXE%"
if exist "%LOCALAPPDATA%\Programs\PhonePolicyManager\PhonePolicyManager.exe" (
  start "" "%LOCALAPPDATA%\Programs\PhonePolicyManager\PhonePolicyManager.exe"
) else (
  start "" "%PYW%" "%APPPY%"
)
echo.
echo  설치 완료! 바탕화면의 [폰정책매니저] 아이콘으로 실행하세요.
timeout /t 4 >nul
exit /b 0

:testfail
echo.
echo  !!! 프로그램 점검에서 오류가 났습니다. 아래 내용을 캡처해서 보내주세요 !!!
echo  파이썬: %PYEXE%
echo  ------------------------------------------------------------
type "%TEMP%\ppm_selftest.txt"
echo  ------------------------------------------------------------
pause
exit /b 1

:copyfail
echo  프로그램을 복사하지 못했습니다. 파일 이름을 짧게 바꾸거나 바탕화면으로 옮긴 뒤 다시 실행해 주세요.
pause
exit /b 1

:pipfail
echo  구성요소 설치에 실패했습니다. 인터넷 연결을 확인하고 다시 실행해 주세요.
pause
exit /b 1

:trypy
if defined PYEXE exit /b 0
py -%1 -c "import sys" >nul 2>&1 || exit /b 0
for /f "delims=" %%P in ('py -%1 -c "import sys;print(sys.executable)"') do set "PYEXE=%%P"
exit /b 0

:trypython
python -c "import sys" >nul 2>&1 || exit /b 0
for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)"') do set "PYEXE=%%P"
exit /b 0
"""


def make_setup_bat(program_source):
    """설치용 .bat 한 파일 = 설치 명령 + 프로그램 본체"""
    marker = "#__PPM_" + "PAYLOAD__"
    text = _BAT_HEAD.strip("\n") + "\n" + marker + "\n" + program_source.replace("\r\n", "\n")
    return text.replace("\n", "\r\n").encode("utf-8")


def install_shortcuts_for_script():
    """파이썬 설치 방식: 바탕화면·시작메뉴에 전용 아이콘 바로가기 만들기"""
    if os.path.isfile(INSTALLED_EXE):        # Setup.exe로 설치된 PC면 그 프로그램 아이콘 유지
        return
    ico = write_ico(os.path.join(BASE_DIR, "app.ico"))
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.isfile(pyw):
        pyw = sys.executable
    for folder in ("Desktop", "Programs"):
        _ps(f"$d=[Environment]::GetFolderPath('{folder}');"
            f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d '{SHORTCUT_NAME}.lnk'));"
            f"$s.TargetPath='{pyw}';$s.Arguments='\"{APP_PY}\"';$s.WorkingDirectory='{BASE_DIR}';"
            f"$s.IconLocation='{ico},0';$s.Description='폰정책매니저';$s.Save()")
    try:   # 바탕화면 아이콘 모양이 바로 바뀌도록 윈도우 아이콘 캐시 새로고침
        subprocess.run(["ie4uinit.exe", "-show"], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=15)
    except Exception:
        pass


def launch_installed():
    if os.path.isfile(INSTALLED_EXE):
        subprocess.Popen([INSTALLED_EXE], cwd=INSTALL_DIR)
        return
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    subprocess.Popen([pyw if os.path.isfile(pyw) else sys.executable, APP_PY], cwd=BASE_DIR)


def script_launcher(app):
    """설치 폴더 밖의 .py 파일을 더블클릭했을 때: 설치 / 업데이트. True면 여기서 끝"""
    if FROZEN or os.name != "nt":
        return False
    me = this_program_script()
    if os.path.normcase(me) == os.path.normcase(APP_PY):
        return False
    cur = read_version(APP_PY) if os.path.isfile(APP_PY) else None
    if cur and ver_tuple(cur) >= ver_tuple(APP_VERSION):
        install_shortcuts_for_script()
        launch_installed()
        return True
    q = (f"폰정책매니저를 새 버전({APP_VERSION})으로 업데이트할까요?\n(지금 버전: {cur})" if cur
         else "폰정책매니저를 이 PC에 설치할까요?\n바탕화면에 폰정책매니저 아이콘이 생깁니다.")
    if QMessageBox.question(None, APP_NAME, q) != QMessageBox.StandardButton.Yes:
        return False
    os.makedirs(os.path.dirname(APP_PY), exist_ok=True)
    if cur:
        shutil.copyfile(APP_PY, APP_PY + ".bak")
    shutil.copyfile(me, APP_PY)
    install_shortcuts_for_script()
    QMessageBox.information(None, APP_NAME, "✅ 완료! 바탕화면의 '폰정책매니저' 아이콘으로 실행하면 됩니다.\n지금 바로 켤게요.")
    launch_installed()
    return True


def selftest():
    """설치 bat이 부름: 실제로 켜질 수 있는지 미리 점검 (창은 안 띄움)"""
    app = QApplication.instance() or QApplication(sys.argv)
    this_program_script(), this_program(), read_version(APP_PY), app_icon()   # 시작 경로에서 쓰는 함수 점검
    db = DB(DB_PATH)
    w = MainWindow(db, quiet=True)
    for i in range(w.tabs.count()):
        w.on_tab(i)
    w.deleteLater()
    print("SELFTEST_OK")


def main():
    step("main 시작 " + " ".join(sys.argv[1:]))
    if "--selftest" in sys.argv:
        selftest()
        return
    if "--build-setup" in sys.argv:              # 설치 bat 끝에서 부름: Setup.exe 만들기
        app = QApplication.instance() or QApplication(sys.argv)
        app.setWindowIcon(app_icon())
        app.setFont(QFont("Malgun Gothic", 10))
        dlg = BuildDialog(None, DB(DB_PATH))
        dlg.setWindowIcon(app_icon())
        dlg.exec()
        return
    if "--install-shortcut" in sys.argv:          # 설치 .bat 이 부름
        install_shortcuts_for_script()
        return
    if FROZEN and not os.environ.get("PPM_INNER"):
        if frozen_install_if_needed():
            return
        newer = frozen_prepare_body()
        if newer:                                  # 업데이트된 본체로 실행
            import runpy
            os.environ["PPM_INNER"] = "1"
            runpy.run_path(newer, run_name="__main__")
            return
    try:   # 작업표시줄에 파이썬 아이콘 대신 이 프로그램 아이콘이 보이게
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("PhonePolicyManager")
    except Exception:
        pass
    app = QApplication.instance() or QApplication(sys.argv)
    app.setWindowIcon(app_icon())
    app.setStyle("Fusion")
    try:
        app.styleHints().setColorScheme(Qt.ColorScheme.Light)
    except Exception:
        pass
    app.setFont(QFont("Malgun Gothic", 10))
    app.setStyleSheet(STYLE)
    if script_launcher(app):
        return
    try:
        db = DB(DB_PATH)
        if FROZEN and first_run_setup(db) == "reload":
            db = DB(DB_PATH)
    except Exception as ex:
        QMessageBox.critical(None, APP_NAME, f"데이터 파일을 열 수 없습니다:\n{ex}")
        return
    step("DB 열기 완료")
    w = MainWindow(db)
    step("화면 만들기 완료")
    w.showMaximized()
    w.raise_()
    w.activateWindow()
    step("화면 표시")
    sys.exit(app.exec())


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        if "--selftest" in sys.argv:
            print(traceback.format_exc())
            sys.exit(1)
        fatal("폰정책매니저를 켜는 중 오류가 났습니다.", traceback.format_exc())

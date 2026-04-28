"""
core.py
=======
데이터 모델 / 상태 / DB / 공통 유틸리티
UI에 일절 의존하지 않는다.
"""

from __future__ import annotations

import logging
import os
import shutil
import sqlite3
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger("QuoteApp")


# ──────────────────────────────────────────────
# 환경 유틸
# ──────────────────────────────────────────────
def exe_dir() -> str:
    """실행 파일(또는 스크립트)이 위치한 디렉터리."""
    if getattr(sys, "frozen", False):
        return str(Path(sys.executable).resolve().parent)
    return str(Path(__file__).resolve().parent)


def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


def unique_path(path: str) -> str:
    """이미 존재하면 _2, _3 … 접미사를 붙여 반환."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while True:
        candidate = f"{base}_{n}{ext}"
        if not os.path.exists(candidate):
            return candidate
        n += 1


def safe_filename(text: str) -> str:
    """파일명으로 쓸 수 없는 문자를 '_'로 치환."""
    t = (text or "").strip()
    if not t:
        return "-"
    for ch in '\\/:*?"<>|':
        t = t.replace(ch, "_")
    return t.strip() or "-"


# ──────────────────────────────────────────────
# 문자열 / 숫자 변환 유틸
# ──────────────────────────────────────────────
def s(v: Any) -> str:
    return "" if v is None else str(v).strip()


def to_float(v: Any, default: float = 0.0) -> float:
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(",", "")
    try:
        return float(t) if t else default
    except ValueError:
        return default


def fmt_qty(x: float) -> str:
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.3f}".rstrip("0").rstrip(".")


def fmt_krw(x: float) -> str:
    return f"{int(round(x)):,}"


def normalize_token(t: str) -> str:
    return s(t).upper().replace(" ", "")


# ──────────────────────────────────────────────
# 시트 / 셀 상수
# ──────────────────────────────────────────────
class SheetName:
    ITEMS       = "품목"
    CODE_MAP    = "코드매핑"
    REQ_COPY    = "견적의뢰복사본"
    SPEC        = "사양서"
    SIGN_SPEC   = "현업 사인용 사양서"
    INCOMING    = "입고검수확인서"
    QUOTE_CN    = "견적서_중국"
    QUOTE_US    = "견적서_미국"
    COVER_DATA  = "갑지DATA"
    COVER       = "견적서 갑지"

    REQUIRED = [
        ITEMS, CODE_MAP, REQ_COPY, SPEC,
        SIGN_SPEC, INCOMING, QUOTE_CN, QUOTE_US,
    ]
    ESIGN_TARGET = [COVER, SPEC, SIGN_SPEC, INCOMING]


@dataclass(frozen=True)
class DomesticLayout:
    SPEC_START  : int = 17
    SPEC_END    : int = 41
    SIGN_START  : int = 11
    SIGN_END    : int = 35
    COL_SPEC    : str = "B"
    COL_QTY     : str = "E"
    COL_PRICE   : str = "F"
    COL_AMT     : str = "G"


@dataclass(frozen=True)
class ChinaLayout:
    MAKER_CELL      : str = "A17"
    PROCESS_CELL    : str = "B17"
    TOOL_CELL       : str = "B51"
    PUMP_START      : int = 17
    PUMP_END        : int = 19
    RACK_START      : int = 21
    RACK_END        : int = 40
    COL_SPEC        : str = "C"
    COL_PRICE       : str = "F"
    COL_QTY         : str = "G"


@dataclass(frozen=True)
class USLayout:
    MAKER_CELL      : str = "A17"
    PROCESS_CELL    : str = "B17"
    TOOL_CELL       : str = "B54"
    EXCHANGE_CELL   : str = "A53"
    DATE_CELL       : str = "E53"
    PUMP_SPEC       : str = "C17"
    PUMP_PRICE      : str = "E17"
    PUMP_QTY        : str = "G17"
    RACK_START      : int = 21
    RACK_END        : int = 40
    COL_SPEC        : str = "C"
    COL_PRICE       : str = "E"
    COL_QTY         : str = "G"


DOM = DomesticLayout()
CN  = ChinaLayout()
US  = USLayout()


# ──────────────────────────────────────────────
# 앱 상태 (순수 데이터, UI 없음)
# ──────────────────────────────────────────────
@dataclass
class QuoteState:
    quote_type          : str                   = "국내"   # "국내" | "중국" | "미국"
    template_path       : Optional[str]         = None
    request_path        : Optional[str]         = None
    request_sheet_name  : Optional[str]         = None

    process : Optional[str] = None
    vendor  : Optional[str] = None
    code_5d : Optional[str] = None

    # 의뢰파일 행 목록: row_idx / D E F H J K V X Z
    request_rows : List[Dict[str, Any]] = field(default_factory=list)

    # 품목시트
    items_rows      : List[Dict[str, Any]]                              = field(default_factory=list)
    items_by_class  : Dict[str, List[Dict[str, Any]]]                   = field(default_factory=dict)
    price_by_spec   : Dict[str, float]                                  = field(default_factory=dict)

    # 코드매핑: (공정, 설비사, 5D) → [(ACC, 수량)]
    code_map : Dict[Tuple[str, str, str], List[Tuple[str, float]]] = field(default_factory=dict)

    # Credit
    pump_credit : float = 0.0
    rack_credit : float = 0.0

    # 중국/미국 팝업 입력값
    cn_info : Dict[str, str]  = field(default_factory=dict)
    us_info : Dict[str, Any]  = field(default_factory=dict)

    last_output_dir : Optional[str] = None


# ──────────────────────────────────────────────
# DB 로그 엔트리
# ──────────────────────────────────────────────
@dataclass
class LogEntry:
    created_at      : str
    action          : str           # "QUOTE" | "COVER" | "ERROR"
    quote_type      : str
    pr              : str = ""
    itemno          : str = ""
    lineproc        : str = ""
    investor        : str = ""
    tool            : str = ""
    template_path   : str = ""
    request_path    : str = ""
    output_path     : str = ""
    message         : str = ""


# ──────────────────────────────────────────────
# DB 관리자
# ──────────────────────────────────────────────
class LogDB:
    """
    SQLite 기반 견적 로그 저장소.
    UI에 완전히 독립적이며 단독 테스트 가능.
    """

    _DDL = """
    CREATE TABLE IF NOT EXISTS quote_logs (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        created_at      TEXT,
        action          TEXT,
        quote_type      TEXT,
        template_path   TEXT,
        request_path    TEXT,
        pr              TEXT,
        itemno          TEXT,
        lineproc        TEXT,
        investor        TEXT,
        tool            TEXT,
        output_path     TEXT,
        message         TEXT
    );
    CREATE TABLE IF NOT EXISTS quote_log_items (
        id      INTEGER PRIMARY KEY AUTOINCREMENT,
        log_id  INTEGER,
        seq     INTEGER,
        role    TEXT,
        cat     TEXT,
        spec    TEXT,
        qty     REAL,
        price   REAL,
        amt     REAL
    );
    CREATE INDEX IF NOT EXISTS idx_logs_pr      ON quote_logs(pr);
    CREATE INDEX IF NOT EXISTS idx_logs_action  ON quote_logs(action);
    CREATE INDEX IF NOT EXISTS idx_items_log_id ON quote_log_items(log_id);
    """

    def __init__(self, db_path: str) -> None:
        self.db_path = os.path.abspath(db_path)
        ensure_dir(os.path.dirname(self.db_path))
        logger.info("LogDB path: %s", self.db_path)
        self._init()

    # ── 연결 ──────────────────────────────────
    def _conn(self) -> sqlite3.Connection:
        con = sqlite3.connect(self.db_path, timeout=10)
        con.execute("PRAGMA journal_mode=WAL;")
        con.execute("PRAGMA synchronous=NORMAL;")
        return con

    def _init(self) -> None:
        with self._conn() as con:
            con.executescript(self._DDL)
            con.commit()

    # ── 쓰기 ──────────────────────────────────
    def insert(self, entry: LogEntry, items: List[Dict[str, Any]]) -> int:
        """
        로그 1건 + 품목 목록을 원자적으로 저장한다.
        반환: 새 log id
        """
        with self._conn() as con:
            cur = con.execute(
                """
                INSERT INTO quote_logs(
                    created_at, action, quote_type,
                    template_path, request_path,
                    pr, itemno, lineproc, investor, tool,
                    output_path, message
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    entry.created_at, entry.action, entry.quote_type,
                    entry.template_path, entry.request_path,
                    entry.pr, entry.itemno, entry.lineproc, entry.investor, entry.tool,
                    entry.output_path, entry.message,
                ),
            )
            log_id = cur.lastrowid
            for seq, it in enumerate(items):
                if it.get("role") == "TOTAL":
                    continue
                con.execute(
                    """
                    INSERT INTO quote_log_items(log_id, seq, role, cat, spec, qty, price, amt)
                    VALUES (?,?,?,?,?,?,?,?)
                    """,
                    (
                        log_id, seq,
                        it.get("role"), it.get("cat"), it.get("spec"),
                        float(it.get("qty") or 0),
                        float(it.get("price") or 0),
                        float(it.get("amt") or 0),
                    ),
                )
            con.commit()
        return int(log_id)

    def insert_simple(self, entry: LogEntry) -> int:
        """품목 없는 단건 저장 (에러/갑지 로그)."""
        return self.insert(entry, [])

    # ── 읽기 ──────────────────────────────────
    def fetch_logs(
        self,
        pr_kw       : str = "",
        tool_kw     : str = "",
        item_kw     : str = "",
        quote_type  : str = "",
    ) -> List[Dict[str, Any]]:
        where : List[str] = ["action='QUOTE'"]
        params: List[Any] = []

        if pr_kw:
            where.append("pr LIKE ?");    params.append(f"%{pr_kw}%")
        if tool_kw:
            where.append("tool LIKE ?");  params.append(f"%{tool_kw}%")
        if quote_type:
            where.append("quote_type=?"); params.append(quote_type)
        if item_kw:
            where.append(
                "EXISTS(SELECT 1 FROM quote_log_items i "
                "WHERE i.log_id=quote_logs.id AND i.spec LIKE ?)"
            )
            params.append(f"%{item_kw}%")

        sql = (
            f"SELECT id,created_at,quote_type,pr,tool,itemno,output_path "
            f"FROM quote_logs "
            f"WHERE {' AND '.join(where)} "
            f"ORDER BY pr ASC, itemno ASC, created_at ASC, id ASC"
        )
        with self._conn() as con:
            rows = con.execute(sql, params).fetchall()
        return [
            dict(id=r[0], created_at=r[1], quote_type=r[2],
                 pr=r[3], tool=r[4], itemno=r[5], output_path=r[6])
            for r in rows
        ]

    def fetch_items(self, log_ids: List[int]) -> List[Dict[str, Any]]:
        if not log_ids:
            return []
        ph = ",".join("?" * len(log_ids))
        sql = (
            f"SELECT log_id,seq,cat,spec,qty,price,amt "
            f"FROM quote_log_items "
            f"WHERE log_id IN ({ph}) "
            f"ORDER BY log_id ASC, seq ASC"
        )
        with self._conn() as con:
            rows = con.execute(sql, tuple(log_ids)).fetchall()
        return [
            dict(log_id=r[0], seq=r[1], cat=r[2],
                 spec=r[3], qty=r[4], price=r[5], amt=r[6])
            for r in rows
        ]

    def fetch_categories(self, log_ids: List[int]) -> List[str]:
        if not log_ids:
            return []
        ph = ",".join("?" * len(log_ids))
        sql = (
            f"SELECT DISTINCT cat FROM quote_log_items "
            f"WHERE log_id IN ({ph}) AND cat IS NOT NULL AND cat!='' "
            f"ORDER BY cat"
        )
        with self._conn() as con:
            return [r[0] for r in con.execute(sql, tuple(log_ids)).fetchall()]

    # ── 유틸 ──────────────────────────────────
    def count_logs(self)  -> int:
        with self._conn() as con:
            return con.execute("SELECT COUNT(*) FROM quote_logs").fetchone()[0]

    def count_items(self) -> int:
        with self._conn() as con:
            return con.execute("SELECT COUNT(*) FROM quote_log_items").fetchone()[0]

    def table_names(self) -> List[str]:
        with self._conn() as con:
            return [
                r[0] for r in con.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
                ).fetchall()
            ]


# ──────────────────────────────────────────────
# DB 마이그레이션 (구버전 경로 → 신버전)
# ──────────────────────────────────────────────
def migrate_db(target_dir: str, exe_dir_path: str) -> Tuple[str, Optional[str]]:
    """
    이전 버전에서 생성된 DB를 현재 위치로 복사한다.
    반환: (status, source_path)
      status: "exists" | "migrated" | "none" | "error:..."
    """
    target = os.path.join(target_dir, "quote_log.db")
    if os.path.exists(target):
        return "exists", None

    candidates = [
        os.path.join(exe_dir_path, "quote_log.db"),
        os.path.join(exe_dir_path, "quote_log", "quote_log.db"),
        os.path.join(os.getcwd(), "quote_log.db"),
        os.path.join(os.getcwd(), "quote_log", "quote_log.db"),
    ]
    for src in candidates:
        if os.path.isfile(src) and os.path.abspath(src) != os.path.abspath(target):
            try:
                ensure_dir(target_dir)
                shutil.copy2(src, target)
                return "migrated", src
            except OSError as e:
                return f"error:{e}", None
    return "none", None

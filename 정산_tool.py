# Add function here.
# 1. Tool should be defined as a function.
# 2. The function argument name should be in the schema as a key.
# ex. ``
# 3. Docstring """ """ should be written to guide proper use.
# 4. The result `content` is returned as string type.
# 5. For now, only python built-in modules can be supported.


import json
import math
import re
from datetime import datetime


# ─────────────────────────────────────────────
# 설정 상수
# ─────────────────────────────────────────────

DB_COL_MAP = {
    "MYDA_ORG_NM":       "제휴사명",
    "SINCDT":            "신청일자",
    "DC_EXE_DT":         "실행일자",
    "DC_SINC_NO":        "대출신청번호",
    "JEHYUSA_DC_PRDT_C": "대출상품코드",
    "DC_PRDT_NM":        "대출상품명",
    "DCAMT":             "대출금액",
    "결과":              "상태",
    "수수료":            "지급수수료",
}

PARTNER_COL_MAP = {
    "신청일자":      "신청일자",
    "실행일자":      "실행일자",
    "대출 신청번호": "대출신청번호",
    "대출신청번호":  "대출신청번호",
    "대출 상품코드": "대출상품코드",
    "대출상품코드":  "대출상품코드",
    "대출 상품명":   "대출상품명",
    "대출상품명":    "대출상품명",
    "대출금액":      "대출금액",
    "상태":          "상태",
    "지급수수료":    "지급수수료",
}

CHECK_COLS = ["신청일자", "실행일자", "대출상품코드", "대출상품명", "대출금액", "상태", "지급수수료"]

STATUS_SYNONYM_GROUPS = [
    ("실행",  frozenset({"실행", "정상", "03-기표(정상)", "기표"})),
    ("철회",  frozenset({"철회", "취소", "해지", "06-철회"})),
    ("심사중", frozenset({"심사중", "처리중"})),
    ("완제",  frozenset({"완제", "종료", "05-종료(완제)"})),
]


# ─────────────────────────────────────────────
# 유틸리티
# ─────────────────────────────────────────────

def _is_empty(v) -> bool:
    if v is None or str(v).strip() in ("", "None", "nan", "-"):
        return True
    try:
        return math.isnan(float(v))
    except (TypeError, ValueError):
        return False


def _to_float(v):
    if _is_empty(v):
        return None
    try:
        return float(str(v).replace(",", "").replace("%", "").strip())
    except (TypeError, ValueError):
        return None


def _fmt_amount(v) -> str:
    f = _to_float(v)
    if f is None:
        return "-"
    return f"{int(f):,}원"


def _normalize_date(v) -> str:
    if _is_empty(v):
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y%m%d")
    s = str(v).strip().replace("-", "").replace("/", "").replace(".", "")
    return s[:8] if len(s) >= 8 else s


def _normalize_status(v: str) -> str:
    s = str(v).strip()
    for rep, synonyms in STATUS_SYNONYM_GROUPS:
        if s in synonyms:
            return rep
    for rep, synonyms in STATUS_SYNONYM_GROUPS:
        for kw in synonyms:
            if kw in s:
                return rep
    return s


def _normalize_month(raw: str) -> str:
    raw = str(raw).strip()
    m = re.match(r"^(\d{4})[-/](\d{1,2})$", raw)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.match(r"^(\d{4})(\d{2})$", raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.match(r"^(\d{1,2})$", raw)
    if m:
        return f"{datetime.today().year}-{int(m.group(1)):02d}"
    return raw


# ─────────────────────────────────────────────
# Markdown 테이블 파싱
# ─────────────────────────────────────────────

def _parse_markdown_table(markdown: str) -> list:
    """
    markdown 문자열에서 테이블(| col | col | ...)을 파싱하여
    [{헤더: 값, ...}, ...] 형태의 리스트로 반환.
    셀 안에 줄바꿈이 포함된 경우(multi-line cell) 연속된 줄을 하나로 병합.
    """
    raw_lines = markdown.splitlines()

    # 연속 줄 병합: '|'로 끝나지 않는 줄은 다음 줄과 이어붙임(셀 내 줄바꿈 처리)
    merged = []
    buf = ""
    for line in raw_lines:
        stripped = line.strip()
        if buf:
            buf = buf + " " + stripped
        else:
            buf = stripped
        if buf.startswith("|") and buf.endswith("|"):
            merged.append(buf)
            buf = ""
        elif not buf.startswith("|"):
            buf = ""

    if not merged:
        return []

    def split_row(line: str) -> list:
        return [c.strip() for c in line.strip().strip("|").split("|")]

    def is_separator(line: str) -> bool:
        cells = split_row(line)
        return bool(cells) and all(re.match(r"^-+$", c.strip()) for c in cells if c.strip())

    separator_idx = None
    for i, line in enumerate(merged):
        if is_separator(line):
            separator_idx = i
            break

    if separator_idx is None or separator_idx == 0:
        return []

    headers = split_row(merged[separator_idx - 1])

    rows = []
    for line in merged[separator_idx + 1:]:
        cells = split_row(line)
        if not cells:
            continue
        while len(cells) < len(headers):
            cells.append("")
        record = {headers[i]: cells[i] for i in range(len(headers))}
        rows.append(record)

    return rows


# ─────────────────────────────────────────────
# 데이터 로드 및 정규화
# ─────────────────────────────────────────────

def _standardize_row(rec: dict, col_map: dict) -> dict:
    return {col_map.get(k, k): v for k, v in rec.items()}


def _normalize_row(rec: dict) -> dict:
    out = dict(rec)
    for k, v in out.items():
        if k in ("신청일자", "실행일자"):
            out[k] = _normalize_date(v)
        elif k in ("대출금액", "지급수수료"):
            out[k] = _to_float(v)
        elif k in ("대출신청번호", "대출상품코드", "대출상품명", "상태"):
            out[k] = str(v).strip() if not _is_empty(v) else ""
    return out


def _load_db_rows(markdown: str, affiliate: str) -> list:
    raw = _parse_markdown_table(markdown)
    std = [_standardize_row(r, DB_COL_MAP) for r in raw]

    exact = [r for r in std if str(r.get("제휴사명", "")).strip() == affiliate]
    if exact:
        return exact

    escaped = re.escape(affiliate)
    partial = [r for r in std if re.search(escaped, str(r.get("제휴사명", "")))]
    if partial:
        return partial

    all_names = sorted({str(r.get("제휴사명", "")).strip() for r in std if r.get("제휴사명")})
    raise ValueError(f"DB에서 '{affiliate}'을 찾을 수 없습니다. 존재하는 제휴사: {all_names}")


def _load_partner_rows(markdown: str, db_key_len: int = 14) -> list:
    raw = _parse_markdown_table(markdown)
    std = [_standardize_row(r, PARTNER_COL_MAP) for r in raw]

    result = []
    for rec in std:
        loan_no = str(rec.get("대출신청번호", "")).strip()
        if not re.match(r"^[A-Za-z0-9]{10,}", loan_no):
            continue
        rec["대출신청번호"] = loan_no[:db_key_len]
        result.append(rec)
    return result


# ─────────────────────────────────────────────
# 중복 합산 그룹화
# ─────────────────────────────────────────────

def _group_by_loan_no(rows: list) -> dict:
    grouped = {}
    AMOUNT_COLS = {"대출금액", "지급수수료"}
    for rec in rows:
        key = rec.get("대출신청번호", "")
        if not key:
            continue
        if key not in grouped:
            grouped[key] = dict(rec)
            grouped[key]["_count"] = 1
        else:
            grouped[key]["_count"] += 1
            for col in AMOUNT_COLS:
                curr = rec.get(col)
                if curr is not None and not (isinstance(curr, float) and math.isnan(curr)):
                    prev = grouped[key].get(col) or 0
                    grouped[key][col] = prev + curr
    return grouped


# ─────────────────────────────────────────────
# 비교 로직
# ─────────────────────────────────────────────

def _compare_value(col: str, db_val, pt_val) -> bool:
    if col in ("대출금액", "지급수수료"):
        db_f = _to_float(db_val)
        pt_f = _to_float(pt_val)
        if db_f is None and pt_f is None:
            return True
        if db_f is None or pt_f is None:
            return False
        return abs(db_f - pt_f) < 1
    elif col == "상태":
        return _normalize_status(str(db_val)) == _normalize_status(str(pt_val))
    else:
        return str(db_val).strip() == str(pt_val).strip()


# ─────────────────────────────────────────────
# 검증 핵심 로직
# ─────────────────────────────────────────────

def _run_verification(db_rows: list, partner_rows: list) -> dict:
    db_rows      = [_normalize_row(r) for r in db_rows]
    partner_rows = [_normalize_row(r) for r in partner_rows]

    db_map = _group_by_loan_no(db_rows)
    pt_map = _group_by_loan_no(partner_rows)

    all_keys = sorted(set(db_map.keys()) | set(pt_map.keys()))

    duplicate_keys   = []
    matched          = []
    mismatch_records = []
    only_in_db       = []
    only_in_partner  = []

    for key in all_keys:
        db_cnt = db_map[key].get("_count", 1) if key in db_map else 0
        pt_cnt = pt_map[key].get("_count", 1) if key in pt_map else 0
        if db_cnt > 1 or pt_cnt > 1:
            duplicate_keys.append({
                "대출신청번호":       key,
                "DB건수":            db_cnt,
                "제휴사건수":         pt_cnt,
                "DB_합산대출금액":    _fmt_amount(db_map[key].get("대출금액")) if key in db_map else "-",
                "제휴_합산대출금액":  _fmt_amount(pt_map[key].get("대출금액")) if key in pt_map else "-",
                "DB_합산지급수수료":  _fmt_amount(db_map[key].get("지급수수료")) if key in db_map else "-",
                "제휴_합산지급수수료": _fmt_amount(pt_map[key].get("지급수수료")) if key in pt_map else "-",
            })

    for key in all_keys:
        in_db = key in db_map
        in_pt = key in pt_map

        if in_db and not in_pt:
            r = db_map[key]
            only_in_db.append({
                "대출신청번호":  key,
                "신청일자":      r.get("신청일자", ""),
                "대출상품명":    r.get("대출상품명", ""),
                "대출금액":      _fmt_amount(r.get("대출금액")),
                "상태":          r.get("상태", ""),
                "지급수수료":    _fmt_amount(r.get("지급수수료")),
            })
            continue

        if in_pt and not in_db:
            r = pt_map[key]
            only_in_partner.append({
                "대출신청번호":  key,
                "신청일자":      r.get("신청일자", ""),
                "대출상품명":    r.get("대출상품명", ""),
                "대출금액":      _fmt_amount(r.get("대출금액")),
                "상태":          r.get("상태", ""),
                "지급수수료":    _fmt_amount(r.get("지급수수료")),
            })
            continue

        db_r = db_map[key]
        pt_r = pt_map[key]

        same_code = (
            str(db_r.get("대출상품코드", "")).strip() ==
            str(pt_r.get("대출상품코드", "")).strip() and
            str(db_r.get("대출상품코드", "")).strip() != ""
        )

        col_mismatches = []
        for col in CHECK_COLS:
            if col == "대출상품명" and same_code:
                continue
            db_val = db_r.get(col, "")
            pt_val = pt_r.get(col, "")
            if not _compare_value(col, db_val, pt_val):
                col_mismatches.append({
                    "항목":    col,
                    "DB값":    _fmt_amount(db_val) if col in ("대출금액", "지급수수료") else str(db_val),
                    "제휴사값": _fmt_amount(pt_val) if col in ("대출금액", "지급수수료") else str(pt_val),
                })

        if col_mismatches:
            mismatch_records.append({"대출신청번호": key, "불일치항목": col_mismatches})
        else:
            matched.append(key)

    return {
        "db_total":              len(db_map),
        "partner_total":         len(pt_map),
        "matched_count":         len(matched),
        "mismatch_count":        len(mismatch_records),
        "only_in_db_count":      len(only_in_db),
        "only_in_partner_count": len(only_in_partner),
        "duplicate_count":       len(duplicate_keys),
        "mismatch_records":      mismatch_records,
        "only_in_db":            only_in_db,
        "only_in_partner":       only_in_partner,
        "duplicate_keys":        duplicate_keys,
    }


# ─────────────────────────────────────────────
# 결과 포맷팅 (마크다운 테이블)
# ─────────────────────────────────────────────

def _md_table(headers: list, rows: list) -> str:
    """리스트를 마크다운 테이블 문자열로 변환."""
    sep = "| " + " | ".join(["---"] * len(headers)) + " |"
    header_line = "| " + " | ".join(headers) + " |"
    data_lines = []
    for row in rows:
        cells = [str(row.get(h, "")) for h in headers]
        data_lines.append("| " + " | ".join(cells) + " |")
    return "\n".join([header_line, sep] + data_lines)


def _format_report(affiliate: str, settlement_month: str, result: dict) -> str:
    today = datetime.today().strftime("%Y-%m-%d")
    total_mis = (
        result["mismatch_count"] +
        result["only_in_db_count"] +
        result["only_in_partner_count"]
    )

    lines = []

    # ── 헤더 ──────────────────────────────────────────────
    lines.append(f"# [{affiliate}] 대출 정산 검증 보고서 ({settlement_month})")
    lines.append(f"기준일: {today}")
    lines.append("")

    # ── 요약 테이블 ────────────────────────────────────────
    lines.append("## 요약")
    lines.append(_md_table(
        ["구분", "건수"],
        [
            {"구분": "제휴사 회신자료 총 건수",    "건수": f"{result['partner_total']}건"},
            {"구분": "DB 자료 총 건수",            "건수": f"{result['db_total']}건"},
            {"구분": "완전 일치",                  "건수": f"{result['matched_count']}건"},
            {"구분": "불일치 합계",                "건수": f"{total_mis}건"},
            {"구분": "┣ 값 불일치 (양쪽 존재)",   "건수": f"{result['mismatch_count']}건"},
            {"구분": "┣ DB에만 존재",              "건수": f"{result['only_in_db_count']}건"},
            {"구분": "┗ 제휴사 회신에만 존재",     "건수": f"{result['only_in_partner_count']}건"},
        ]
    ))
    lines.append("")

    # ── 중복 신청번호 ──────────────────────────────────────
    if result["duplicate_count"] > 0:
        lines.append(f"## ⚠ 중복 신청번호 — 금액 합산 처리 ({result['duplicate_count']}건)")
        lines.append(_md_table(
            ["대출신청번호", "DB건수", "제휴사건수", "DB합산대출금액", "제휴합산대출금액", "DB합산지급수수료", "제휴합산지급수수료"],
            [
                {
                    "대출신청번호":   d["대출신청번호"],
                    "DB건수":        str(d["DB건수"]),
                    "제휴사건수":     str(d["제휴사건수"]),
                    "DB합산대출금액":  d["DB_합산대출금액"],
                    "제휴합산대출금액": d["제휴_합산대출금액"],
                    "DB합산지급수수료": d["DB_합산지급수수료"],
                    "제휴합산지급수수료": d["제휴_합산지급수수료"],
                }
                for d in result["duplicate_keys"]
            ]
        ))
        lines.append("")

    # ── 모두 일치 ──────────────────────────────────────────
    if total_mis == 0:
        lines.append("> ✔ 모든 항목이 일치합니다.")
        return "\n".join(lines)

    # ── 값 불일치 명세 ─────────────────────────────────────
    if result["mismatch_records"]:
        lines.append(f"## 값 불일치 명세 ({result['mismatch_count']}건)")
        # 대출신청번호 1건당 불일치 항목을 행으로 풀어서 표시
        flat_rows = []
        for rec in result["mismatch_records"]:
            loan_no = rec["대출신청번호"]
            for i, mis in enumerate(rec["불일치항목"]):
                flat_rows.append({
                    "대출신청번호": loan_no if i == 0 else "",
                    "불일치항목":   mis["항목"],
                    "DB값":         mis["DB값"],
                    "제휴사값":     mis["제휴사값"],
                })
        lines.append(_md_table(["대출신청번호", "불일치항목", "DB값", "제휴사값"], flat_rows))
        lines.append("")

    # ── DB에만 존재 ────────────────────────────────────────
    if result["only_in_db"]:
        lines.append(f"## DB에만 존재 — 제휴사 회신 누락 ({result['only_in_db_count']}건)")
        lines.append(_md_table(
            ["대출신청번호", "신청일자", "대출상품명", "대출금액", "상태", "지급수수료"],
            result["only_in_db"]
        ))
        lines.append("")

    # ── 제휴사에만 존재 ────────────────────────────────────
    if result["only_in_partner"]:
        lines.append(f"## 제휴사에만 존재 — DB 미등록 ({result['only_in_partner_count']}건)")
        lines.append(_md_table(
            ["대출신청번호", "신청일자", "대출상품명", "대출금액", "상태", "지급수수료"],
            result["only_in_partner"]
        ))
        lines.append("")

    return "\n".join(lines)


# ─────────────────────────────────────────────
# Tool 진입점
# ─────────────────────────────────────────────

def run_settlement_check(content: list, affiliate: str, settlement_month: str) -> str:
    """
    제휴사 대출 정산 대사(對査)를 수행하는 Tool 함수.

    플랫폼이 Excel 파일을 마크다운 테이블 문자열로 변환하여 content 배열로 전달합니다.
    제휴사명과 정산월은 앞단 LLM이 추출하여 별도 인자로 전달합니다.

    Parameters
    ----------
    content : list
        플랫폼이 전달하는 파일 목록. 각 항목은 아래 필드를 포함합니다.
        - doc_nm  : 원본 파일명 (예: "2604_F_기반_Agent_자료.xlsx")
        - content : Excel을 변환한 마크다운 테이블 문자열

        DB 총본 판별  : content 문자열에 'MYDA_ORG_NM' 또는 'SINCDT' 컬럼이 존재
        제휴사 파일 판별 : content 문자열에 '신청일자' 컬럼이 존재

    affiliate : str
        제휴사명 (예: "BNK저축은행", "KB저축은행")
        앞단 LLM이 사용자 질의에서 추출하여 전달.

    settlement_month : str
        정산 대상 연월 (예: "2026-04", "2026-05", "04")
        앞단 LLM이 사용자 질의에서 추출하여 전달.

    Returns
    -------
    str
        정산 검증 결과 보고서 (마크다운 테이블 형식).
        LLM이 바로 읽고 해석할 수 있는 구조로 반환됩니다.
        오류 발생 시 오류 메시지 문자열 반환.

    Examples
    --------
    content 항목 구조:
        [
            {
                "doc_nm": "2604_F_기반_Agent_자료.xlsx",
                "content": "| MYDA_ORG_NM | SINCDT | DC_SINC_NO | ... |\\n| --- | ... |\\n| BNK저축은행 | ..."
            },
            {
                "doc_nm": "BNK저축은행_2604_대출.xlsx",
                "content": "| 순번 | 신청일자 | 실행일자 | 대출 신청번호 | ... |\\n| --- | ... |"
            }
        ]

    affiliate  = "BNK저축은행"
    settlement_month = "2026-04"
    """
    try:
        affiliate        = str(affiliate).strip()
        settlement_month = _normalize_month(str(settlement_month).strip())

        if not affiliate:
            return "[오류] affiliate(제휴사명)가 비어 있습니다."

        db_markdown      = ""
        partner_markdown = ""

        for item in content:
            # 플랫폼 필드명: 'content' 우선, 없으면 'markdown' fallback
            md = str(item.get("content") or item.get("markdown") or "")

            if "MYDA_ORG_NM" in md or "SINCDT" in md:
                db_markdown = md
            elif "신청일자" in md:
                partner_markdown = md

        if not db_markdown:
            return "[오류] DB 총본 파일을 찾을 수 없습니다. 'MYDA_ORG_NM' 또는 'SINCDT' 컬럼이 포함된 항목을 content에 포함해주세요."
        if not partner_markdown:
            return "[오류] 제휴사 회신 파일을 찾을 수 없습니다. '신청일자' 컬럼이 포함된 항목을 content에 포함해주세요."

        db_rows = _load_db_rows(db_markdown, affiliate)

        key_lengths = [
            len(str(r.get("대출신청번호", "") or ""))
            for r in db_rows if r.get("대출신청번호")
        ]
        db_key_len = max(set(key_lengths), key=key_lengths.count) if key_lengths else 14

        partner_rows = _load_partner_rows(partner_markdown, db_key_len=db_key_len)

        result = _run_verification(db_rows, partner_rows)

        return _format_report(affiliate, settlement_month or "미상", result)

    except ValueError as e:
        return f"[오류] {e}"
    except Exception as e:
        return f"[오류] 검증 중 예외 발생: {e}"

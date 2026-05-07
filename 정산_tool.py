# Add function here.
# 1. Tool should be defined as a function.
# 2. The function argument name should be in the schema as a key.
# ex. ``
# 3. Docstring """ """ should be written to guide proper use.
# 4. The result `content` is returned as string type.
# 5. For now, only python built-in modules can be supported.


import json
import math
import os
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
    if "." in s:
        s = s.split(".")[0]
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

    # ── 연속 줄 병합 ───────────────────────────────────────
    # 어떤 줄이 '|'로 끝나지 않으면 다음 줄과 이어붙임(셀 내 줄바꿈 처리)
    merged: list[str] = []
    buf = ""
    for line in raw_lines:
        stripped = line.strip()
        if buf:
            buf = buf + " " + stripped
        else:
            buf = stripped
        # 완전한 행: '|'로 시작하고 '|'로 끝남
        if buf.startswith("|") and buf.endswith("|"):
            merged.append(buf)
            buf = ""
        elif not buf.startswith("|"):
            # 테이블 행이 아님 → 버림
            buf = ""

    if not merged:
        return []

    def split_row(line: str) -> list:
        return [c.strip() for c in line.strip().strip("|").split("|")]

    def is_separator(line: str) -> bool:
        cells = split_row(line)
        return bool(cells) and all(re.match(r"^-+$", c.strip()) for c in cells if c.strip())

    # 구분선 위치 탐지
    separator_idx = None
    for i, line in enumerate(merged):
        if is_separator(line):
            separator_idx = i
            break

    if separator_idx is None or separator_idx == 0:
        return []

    # 헤더: 구분선 바로 위 줄
    headers = split_row(merged[separator_idx - 1])

    # 데이터 행: 구분선 아래
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

    # 정확한 이름 일치 우선
    exact = [r for r in std if str(r.get("제휴사명", "")).strip() == affiliate]
    if exact:
        return exact

    # 부분 일치 fallback
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
        # 영숫자 10자 이상인 실제 신청번호만 유지
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
    db_rows     = [_normalize_row(r) for r in db_rows]
    partner_rows = [_normalize_row(r) for r in partner_rows]

    db_map = _group_by_loan_no(db_rows)
    pt_map = _group_by_loan_no(partner_rows)

    all_keys = sorted(set(db_map.keys()) | set(pt_map.keys()))

    duplicate_keys  = []
    matched         = []
    mismatch_records = []
    only_in_db      = []
    only_in_partner  = []

    for key in all_keys:
        db_cnt = db_map[key].get("_count", 1) if key in db_map else 0
        pt_cnt = pt_map[key].get("_count", 1) if key in pt_map else 0
        if db_cnt > 1 or pt_cnt > 1:
            duplicate_keys.append({
                "대출신청번호":      key,
                "DB건수":           db_cnt,
                "제휴사건수":        pt_cnt,
                "DB_합산대출금액":   _fmt_amount(db_map[key].get("대출금액")) if key in db_map else "-",
                "제휴_합산대출금액": _fmt_amount(pt_map[key].get("대출금액")) if key in pt_map else "-",
                "DB_합산지급수수료":  _fmt_amount(db_map[key].get("지급수수료")) if key in db_map else "-",
                "제휴_합산지급수수료": _fmt_amount(pt_map[key].get("지급수수료")) if key in pt_map else "-",
                "비고": "금액 합산 후 비교",
            })

    for key in all_keys:
        in_db = key in db_map
        in_pt = key in pt_map

        if in_db and not in_pt:
            r = db_map[key]
            only_in_db.append({
                "대출신청번호":  key,
                "DB_신청일자":   r.get("신청일자", ""),
                "DB_대출상품명": r.get("대출상품명", ""),
                "DB_대출금액":   _fmt_amount(r.get("대출금액")),
                "DB_상태":       r.get("상태", ""),
                "DB_지급수수료": _fmt_amount(r.get("지급수수료")),
            })
            continue

        if in_pt and not in_db:
            r = pt_map[key]
            only_in_partner.append({
                "대출신청번호":      key,
                "제휴사_신청일자":   r.get("신청일자", ""),
                "제휴사_대출상품명": r.get("대출상품명", ""),
                "제휴사_대출금액":   _fmt_amount(r.get("대출금액")),
                "제휴사_상태":       r.get("상태", ""),
                "제휴사_지급수수료": _fmt_amount(r.get("지급수수료")),
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
        "db_total":         len(db_map),
        "partner_total":    len(pt_map),
        "matched_count":    len(matched),
        "mismatch_count":   len(mismatch_records),
        "only_in_db_count": len(only_in_db),
        "only_in_partner_count": len(only_in_partner),
        "duplicate_count":  len(duplicate_keys),
        "mismatch_records": mismatch_records,
        "only_in_db":       only_in_db,
        "only_in_partner":  only_in_partner,
        "duplicate_keys":   duplicate_keys,
    }


# ─────────────────────────────────────────────
# 결과 포맷팅 (문자열 리포트)
# ─────────────────────────────────────────────

def _format_report(affiliate: str, settlement_month: str, result: dict) -> str:
    SEP = "=" * 65
    SUB = "-" * 65
    lines = []

    lines.append(SEP)
    lines.append(f"  [{affiliate}] 대출 정산 검증 보고서  ({settlement_month})")
    lines.append(f"  기준일: {datetime.today().strftime('%Y-%m-%d')}")
    lines.append(SEP)

    total_mis = result["mismatch_count"] + result["only_in_db_count"] + result["only_in_partner_count"]

    lines.append("\n[ 요약 ]")
    lines.append(SUB)
    lines.append(f"  제휴사 회신자료 기준 총 건수  : {result['partner_total']:>4}건")
    lines.append(f"  DB 자료 기준 총 건수          : {result['db_total']:>4}건")
    lines.append(f"  완전 일치                     : {result['matched_count']:>4}건")
    lines.append(f"  불일치 (합계)                 : {total_mis:>4}건")
    lines.append(f"    ├ 값 불일치 (양쪽 존재)     : {result['mismatch_count']:>4}건")
    lines.append(f"    ├ DB에만 존재               : {result['only_in_db_count']:>4}건")
    lines.append(f"    └ 제휴사 회신에만 존재      : {result['only_in_partner_count']:>4}건")
    if result["duplicate_count"] > 0:
        lines.append(f"\n  ⚠ 중복 신청번호 (금액 합산 처리) : {result['duplicate_count']}건")
        for d in result["duplicate_keys"]:
            lines.append(f"    • {d['대출신청번호']}  DB {d['DB건수']}건({d['DB_합산대출금액']}) / 제휴 {d['제휴사건수']}건({d['제휴_합산대출금액']})")
    lines.append(SUB)

    if total_mis == 0:
        lines.append("\n  ✔ 모든 항목이 일치합니다.")
        lines.append(SEP)
        return "\n".join(lines)

    # 값 불일치 명세 (대출신청번호 1건 = 1행 형식)
    if result["mismatch_records"]:
        lines.append(f"\n【 값 불일치 명세 ({result['mismatch_count']}건) 】")
        lines.append(f"  {'대출신청번호':<16} {'불일치항목':<30} {'DB값':<20} {'제휴사값'}")
        lines.append(f"  {'-'*16} {'-'*30} {'-'*20} {'-'*20}")
        for rec in result["mismatch_records"]:
            loan_no = rec["대출신청번호"]
            for i, mis in enumerate(rec["불일치항목"]):
                prefix = loan_no if i == 0 else ""
                lines.append(f"  {prefix:<16} {mis['항목']:<30} {mis['DB값']:<20} {mis['제휴사값']}")

    # DB에만 존재
    if result["only_in_db"]:
        lines.append(f"\n【 DB에만 존재 — 제휴사 회신 누락 ({result['only_in_db_count']}건) 】")
        lines.append(f"  {'대출신청번호':<16} {'상품명':<20} {'대출금액':<14} {'상태':<10} {'지급수수료'}")
        lines.append(f"  {'-'*16} {'-'*20} {'-'*14} {'-'*10} {'-'*12}")
        for r in result["only_in_db"]:
            lines.append(f"  {r['대출신청번호']:<16} {r['DB_대출상품명']:<20} {r['DB_대출금액']:<14} {r['DB_상태']:<10} {r['DB_지급수수료']}")

    # 제휴사에만 존재
    if result["only_in_partner"]:
        lines.append(f"\n【 제휴사에만 존재 — DB 미등록 ({result['only_in_partner_count']}건) 】")
        lines.append(f"  {'대출신청번호':<16} {'상품명':<20} {'대출금액':<14} {'상태':<10} {'지급수수료'}")
        lines.append(f"  {'-'*16} {'-'*20} {'-'*14} {'-'*10} {'-'*12}")
        for r in result["only_in_partner"]:
            lines.append(f"  {r['대출신청번호']:<16} {r['제휴사_대출상품명']:<20} {r['제휴사_대출금액']:<14} {r['제휴사_상태']:<10} {r['제휴사_지급수수료']}")

    lines.append(f"\n{SEP}")
    return "\n".join(lines)


# ─────────────────────────────────────────────
# Tool 진입점
# ─────────────────────────────────────────────

def run_settlement_check(content: list) -> str:
    """
    제휴사 대출 정산 대사(對査)를 수행하는 Tool 함수.

    플랫폼이 LLM을 통해 추출한 제휴사명·정산월과 함께
    두 개의 Excel 파일(DB 총본, 제휴사 회신)을 markdown으로 변환하여
    content 배열로 전달합니다.

    Parameters
    ----------
    content : list
        플랫폼이 전달하는 파일 목록. 각 항목은 아래 필드를 포함합니다.
        - uuid       : 파일 고유 ID
        - fileName   : 원본 파일명 (예: "2604_F_기반_Agent_자료.xlsx")
        - markdown   : 파일 내용을 변환한 markdown 테이블 문자열
        - ts         : 업로드 타임스탬프

        content[0] : DB 총본 파일 (플랫폼 기반자료)
        content[1] : 제휴사 회신 파일

        단, fileName에 'affiliate'와 'settlement_month' 키가 포함된
        JSON 메타데이터가 함께 전달될 수도 있습니다.
        예) fileName = "KB저축은행_26년4월_대출정산.xlsx"
            또는 별도 메타 content 항목에
            {"affiliate": "KB저축은행", "settlement_month": "2026-04"} 포함.

    Returns
    -------
    str
        정산 검증 결과 보고서 (텍스트 형식).
        오류 발생 시 오류 메시지 반환.

    Examples
    --------
    입력 JSON 스키마 (플랫폼 전달 형식):
        {
            "content": [
                {
                    "uuid": "aaa-111",
                    "fileName": "2604_F_기반_Agent_자료.xlsx",
                    "markdown": "| MYDA_ORG_NM | SINCDT | ... |\\n|---|...",
                    "ts": "2026-05-07T09:00:00"
                },
                {
                    "uuid": "bbb-222",
                    "fileName": "KB저축은행_26년4월_대출정산.xlsx",
                    "markdown": "| 순번 | 신청일자 | ... |\\n|---|...",
                    "ts": "2026-05-07T09:00:01"
                },
                {
                    "uuid": "ccc-333",
                    "fileName": "meta",
                    "markdown": "{\"affiliate\": \"KB저축은행\", \"settlement_month\": \"2026-04\"}",
                    "ts": "2026-05-07T09:00:02"
                }
            ]
        }
    """
    try:
        # ── Step 1: content 항목 분류 ──────────────────────────
        # affiliate/settlement_month 메타 추출
        affiliate       = ""
        settlement_month = ""
        db_markdown     = ""
        partner_markdown = ""

        for item in content:
            file_name = str(item.get("fileName", ""))
            markdown  = str(item.get("markdown", ""))

            # JSON 메타데이터 항목 판별 (markdown이 JSON 문자열인 경우)
            stripped = markdown.strip()
            if stripped.startswith("{"):
                try:
                    meta = json.loads(stripped)
                    if "affiliate" in meta:
                        affiliate = str(meta["affiliate"]).strip()
                    if "settlement_month" in meta:
                        settlement_month = _normalize_month(str(meta["settlement_month"]))
                    continue
                except (json.JSONDecodeError, ValueError):
                    pass

            # DB 총본 판별: 헤더에 'MYDA_ORG_NM' 또는 'SINCDT' 포함
            if "MYDA_ORG_NM" in markdown or "SINCDT" in markdown:
                db_markdown = markdown
                continue

            # 제휴사 회신 파일 판별: 헤더에 '신청일자' 포함
            if "신청일자" in markdown and db_markdown != markdown:
                partner_markdown = markdown
                continue

        # affiliate가 메타에 없으면 partner fileName에서 추출 시도
        if not affiliate:
            for item in content:
                fn = str(item.get("fileName", ""))
                # fileName에서 알려진 저축은행명 패턴 탐지
                m = re.search(r"([가-힣A-Za-z]+저축은행|[가-힣A-Za-z]+캐피탈|[가-힣A-Za-z]+은행)", fn)
                if m:
                    affiliate = m.group(1)
                    break

        # settlement_month가 없으면 파일명에서 연월 추출 시도
        if not settlement_month:
            for item in content:
                fn = str(item.get("fileName", ""))
                m = re.search(r"(\d{2,4})[-_년]?\s*(\d{1,2})[-_월]?", fn)
                if m:
                    year_part  = m.group(1)
                    month_part = m.group(2)
                    if len(year_part) == 2:
                        year_part = "20" + year_part
                    settlement_month = f"{year_part}-{int(month_part):02d}"
                    break

        # ── Step 2: 필수 값 검증 ──────────────────────────────
        if not affiliate:
            return "[오류] 제휴사명을 찾을 수 없습니다. content에 affiliate 메타데이터 또는 제휴사명이 포함된 fileName을 전달해주세요."
        if not db_markdown:
            return "[오류] DB 총본 파일을 찾을 수 없습니다. MYDA_ORG_NM 컬럼이 포함된 markdown을 content에 포함해주세요."
        if not partner_markdown:
            return "[오류] 제휴사 회신 파일을 찾을 수 없습니다. 신청일자 컬럼이 포함된 markdown을 content에 포함해주세요."

        # ── Step 3: 데이터 로드 ───────────────────────────────
        db_rows = _load_db_rows(db_markdown, affiliate)

        # DB 신청번호 길이 자동 파악
        key_lengths = [
            len(str(r.get("대출신청번호", "") or ""))
            for r in db_rows if r.get("대출신청번호")
        ]
        db_key_len = max(set(key_lengths), key=key_lengths.count) if key_lengths else 14

        partner_rows = _load_partner_rows(partner_markdown, db_key_len=db_key_len)

        # ── Step 4: 검증 실행 ─────────────────────────────────
        result = _run_verification(db_rows, partner_rows)

        # ── Step 5: 보고서 반환 ───────────────────────────────
        return _format_report(affiliate, settlement_month or "미상", result)

    except ValueError as e:
        return f"[오류] {e}"
    except Exception as e:
        return f"[오류] 검증 중 예외 발생: {e}"

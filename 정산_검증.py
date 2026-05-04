"""
제휴사 정산 검증 스크립트
=======================
사용법:
    python 정산_검증.py <DB파일.xlsx> <제휴사회신파일.xlsx> <제휴사명>

예시:
    python 정산_검증.py 2604_F_기반_Agent_자료.xlsx BNK저축은행_2604_대출.xlsx BNK저축은행
"""

import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime


# ─────────────────────────────────────────────
# 1. 파일 로딩 설정
# ─────────────────────────────────────────────

# DB 파일 컬럼 → 표준명 매핑
DB_COL_MAP = {
    "SINCDT":           "신청일자",
    "DC_EXE_DT":        "실행일자",
    "DC_SINC_NO":       "대출신청번호",
    "JEHYUSA_DC_PRDT_C":"대출상품코드",
    "DC_PRDT_NM":       "대출상품명",
    "DCAMT":            "대출금액",
    "결과":             "상태",
    "수수료":           "지급수수료",
    "MYDA_ORG_NM":      "제휴사명",
}

# 제휴사 회신 파일 컬럼 → 표준명 매핑
PARTNER_COL_MAP = {
    "신청일자":     "신청일자",
    "실행일자":     "실행일자",
    "대출 신청번호": "대출신청번호",
    "대출 신청번호":"대출신청번호",
    "대출신청번호": "대출신청번호",
    "대출 상품코드": "대출상품코드",
    "대출상품코드": "대출상품코드",
    "대출 상품명":  "대출상품명",
    "대출상품명":   "대출상품명",
    "대출금액":     "대출금액",
    "상태":         "상태",
    "지급수수료":   "지급수수료",
}

# 검증 대상 컬럼 (표준명 기준)
CHECK_COLS = ["신청일자", "실행일자", "대출상품코드", "대출상품명", "대출금액", "상태", "지급수수료"]


# ─────────────────────────────────────────────
# 2. 데이터 로딩 함수
# ─────────────────────────────────────────────

def find_header_row(file_path: str, sheet_name=0, search_keyword: str = "신청일자", max_rows: int = 10) -> int:
    """헤더 행 위치를 자동 탐지"""
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=max_rows)
    for i, row in df_raw.iterrows():
        if row.astype(str).str.contains(search_keyword, na=False).any():
            return i
    return 0


def load_db(file_path: str, partner_name: str) -> pd.DataFrame:
    """DB 파일에서 특정 제휴사 데이터 로드"""
    df = pd.read_excel(file_path, sheet_name="대출시트")
    df = df.rename(columns=DB_COL_MAP)

    # 제휴사 필터링 (부분 일치)
    mask = df["제휴사명"].str.contains(partner_name, na=False)
    df = df[mask].copy()

    if df.empty:
        raise ValueError(f"DB에서 '{partner_name}' 데이터를 찾을 수 없습니다.\n"
                         f"존재하는 제휴사: {df['제휴사명'].unique().tolist()}")

    df = df.reset_index(drop=True)
    return df


def load_partner(file_path: str, db_key_len: int = 14) -> pd.DataFrame:
    """제휴사 회신 파일 로드 (헤더 행 자동 탐지)"""
    # 헤더 행 탐지
    sheet_name = 0
    header_row = find_header_row(file_path, sheet_name=sheet_name, search_keyword="신청일자")

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

    # 컬럼명 표준화 (rename 먼저)
    df = df.rename(columns=PARTNER_COL_MAP)

    # 대출신청번호가 존재하고 실제 값이 있는 행만 남김
    if "대출신청번호" in df.columns:
        df["대출신청번호"] = df["대출신청번호"].astype(str).str.strip()
        # 'nan', 빈 문자열, 숫자만인 행 제거 → 실제 신청번호(영문+숫자 혼합) 행만 유지
        df = df[df["대출신청번호"].str.match(r"^[A-Za-z0-9]{10,}", na=False)]
        # 상품코드가 붙어 있는 경우 DB 키 길이만큼만 사용
        df["대출신청번호"] = df["대출신청번호"].str[:db_key_len]
    else:
        # 순번 컬럼 기준 필터링 (대출신청번호 컬럼명이 다를 경우 fallback)
        if "순번" in df.columns:
            df = df[pd.to_numeric(df["순번"], errors="coerce").notna()]

    df = df.reset_index(drop=True)
    return df


# ─────────────────────────────────────────────
# 3. 데이터 정규화 함수
# ─────────────────────────────────────────────

def normalize_date(val) -> str:
    """날짜 값을 YYYYMMDD 문자열로 정규화"""
    if pd.isna(val):
        return ""
    s = str(val).strip().replace("-", "").replace("/", "").replace(".", "")
    # float 형태 (예: 20260407.0) 처리
    if "." in s:
        s = s.split(".")[0]
    return s[:8] if len(s) >= 8 else s


def normalize_amount(val) -> float:
    """금액을 float으로 정규화"""
    if pd.isna(val):
        return np.nan
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return np.nan


def normalize_str(val) -> str:
    """문자열 정규화"""
    if pd.isna(val):
        return ""
    return str(val).strip()


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """데이터프레임 전체 정규화"""
    df = df.copy()
    date_cols = ["신청일자", "실행일자"]
    amount_cols = ["대출금액", "지급수수료"]
    str_cols = ["대출신청번호", "대출상품코드", "대출상품명", "상태"]

    for col in date_cols:
        if col in df.columns:
            df[col] = df[col].apply(normalize_date)

    for col in amount_cols:
        if col in df.columns:
            df[col] = df[col].apply(normalize_amount)

    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].apply(normalize_str)

    return df


# ─────────────────────────────────────────────
# 4. 검증 로직
# ─────────────────────────────────────────────

# 상태값 동의어 그룹 (같은 그룹 내 값은 동일로 처리)
STATUS_SYNONYM_GROUPS = [
    {"실행", "정상"},
    {"철회", "취소", "해지"},
    {"심사중", "처리중"},
]

def normalize_status(val: str) -> str:
    """상태값을 동의어 그룹의 대표값(첫 번째 원소)으로 정규화"""
    v = str(val).strip()
    for group in STATUS_SYNONYM_GROUPS:
        if v in group:
            return sorted(group)[0]  # 알파벳/가나다 정렬 첫 번째를 대표값으로
    return v


def compare_value(col: str, db_val, partner_val) -> bool:
    """두 값이 일치하는지 비교"""
    if col in ["대출금액", "지급수수료"]:
        if pd.isna(db_val) and pd.isna(partner_val):
            return True
        if pd.isna(db_val) or pd.isna(partner_val):
            return False
        return abs(float(db_val) - float(partner_val)) < 1  # 1원 미만 오차 허용
    elif col == "상태":
        return normalize_status(db_val) == normalize_status(partner_val)
    else:
        return str(db_val).strip() == str(partner_val).strip()


def run_verification(db_df: pd.DataFrame, partner_df: pd.DataFrame) -> dict:
    """
    대출신청번호를 기준으로 DB vs 제휴사 회신자료 검증
    Returns: 결과 딕셔너리
    """
    db_df = normalize_df(db_df)
    partner_df = normalize_df(partner_df)

    db_key = "대출신청번호"
    db_indexed = db_df.set_index(db_key)
    partner_indexed = partner_df.set_index(db_key)

    all_keys = set(db_indexed.index) | set(partner_indexed.index)

    mismatch_records = []   # 값 불일치
    only_in_db = []         # DB에만 있음
    only_in_partner = []    # 제휴사에만 있음
    matched = []            # 완전 일치

    for key in sorted(all_keys):
        in_db = key in db_indexed.index
        in_partner = key in partner_indexed.index

        if in_db and not in_partner:
            row = db_indexed.loc[key]
            only_in_db.append({
                "대출신청번호": key,
                "DB_신청일자": row.get("신청일자", ""),
                "DB_대출상품명": row.get("대출상품명", ""),
                "DB_대출금액": row.get("대출금액", ""),
                "DB_상태": row.get("상태", ""),
                "DB_지급수수료": row.get("지급수수료", ""),
            })
            continue

        if in_partner and not in_db:
            row = partner_indexed.loc[key]
            only_in_partner.append({
                "대출신청번호": key,
                "제휴사_신청일자": row.get("신청일자", ""),
                "제휴사_대출상품명": row.get("대출상품명", ""),
                "제휴사_대출금액": row.get("대출금액", ""),
                "제휴사_상태": row.get("상태", ""),
                "제휴사_지급수수료": row.get("지급수수료", ""),
            })
            continue

        # 양쪽 모두 있는 경우 → 항목별 비교
        db_row = db_indexed.loc[key]
        pt_row = partner_indexed.loc[key]

        # 상품코드 일치 여부를 먼저 확인 (상품명 비교 면제 조건)
        db_prdt_code = str(db_row.get("대출상품코드", "")).strip()
        pt_prdt_code = str(pt_row.get("대출상품코드", "")).strip()
        same_product_code = (db_prdt_code == pt_prdt_code and db_prdt_code != "")

        col_mismatches = []
        for col in CHECK_COLS:
            # 상품코드가 동일하면 상품명 불일치는 무시
            if col == "대출상품명" and same_product_code:
                continue
            db_val = db_row.get(col, "")
            pt_val = pt_row.get(col, "")
            if not compare_value(col, db_val, pt_val):
                col_mismatches.append({
                    "항목": col,
                    "DB값": db_val,
                    "제휴사값": pt_val,
                })

        if col_mismatches:
            mismatch_records.append({
                "대출신청번호": key,
                "불일치항목": col_mismatches,
            })
        else:
            matched.append(key)

    return {
        "db_total": len(db_indexed),
        "partner_total": len(partner_indexed),
        "matched": matched,
        "mismatch_records": mismatch_records,
        "only_in_db": only_in_db,
        "only_in_partner": only_in_partner,
    }


# ─────────────────────────────────────────────
# 5. 보고서 출력
# ─────────────────────────────────────────────

def fmt_amount(val) -> str:
    try:
        return f"{int(float(val)):,}원"
    except Exception:
        return str(val)


def print_report(result: dict, partner_name: str):
    """검증 결과 보고서 출력"""
    SEP = "=" * 65
    SUB = "-" * 65

    db_total = result["db_total"]
    pt_total = result["partner_total"]
    mismatch_count = (
        len(result["mismatch_records"])
        + len(result["only_in_db"])
        + len(result["only_in_partner"])
    )
    matched_count = len(result["matched"])

    print(SEP)
    print(f"  [{partner_name}] 대출 정산 검증 보고서")
    print(f"  기준: {datetime.today().strftime('%Y-%m-%d')}")
    print(SEP)

    print(f"\n{'[ 요약 ]':}")
    print(SUB)
    print(f"  제휴사 회신자료 기준 총 건수  : {pt_total:>4}건")
    print(f"  DB 자료 기준 총 건수          : {db_total:>4}건")
    print(f"  완전 일치                     : {matched_count:>4}건")
    print(f"  불일치 (합계)                 : {mismatch_count:>4}건")
    print(f"    ├ 값 불일치 (양쪽 존재)     : {len(result['mismatch_records']):>4}건")
    print(f"    ├ DB에만 존재               : {len(result['only_in_db']):>4}건")
    print(f"    └ 제휴사 회신에만 존재      : {len(result['only_in_partner']):>4}건")
    print(SUB)

    if mismatch_count == 0:
        print("\n  ✔ 모든 항목이 일치합니다.")
        print(SEP)
        return

    # ── 값 불일치 상세 ──────────────────────────────────
    if result["mismatch_records"]:
        print(f"\n【 값 불일치 명세 ({len(result['mismatch_records'])}건) 】")
        for i, rec in enumerate(result["mismatch_records"], 1):
            print(f"\n  [{i}] 대출신청번호: {rec['대출신청번호']}")
            for mis in rec["불일치항목"]:
                col = mis["항목"]
                db_v = mis["DB값"]
                pt_v = mis["제휴사값"]
                if col in ["대출금액", "지급수수료"]:
                    db_v = fmt_amount(db_v)
                    pt_v = fmt_amount(pt_v)
                print(f"      ▸ {col}")
                print(f"        DB        : {db_v}")
                print(f"        제휴사    : {pt_v}")

    # ── DB에만 있는 건 ──────────────────────────────────
    if result["only_in_db"]:
        print(f"\n【 DB에만 존재 (제휴사 회신 누락) — {len(result['only_in_db'])}건 】")
        for i, rec in enumerate(result["only_in_db"], 1):
            print(f"\n  [{i}] 대출신청번호: {rec['대출신청번호']}")
            print(f"       신청일자  : {rec['DB_신청일자']}")
            print(f"       상품명    : {rec['DB_대출상품명']}")
            print(f"       대출금액  : {fmt_amount(rec['DB_대출금액'])}")
            print(f"       상태      : {rec['DB_상태']}")
            print(f"       지급수수료: {fmt_amount(rec['DB_지급수수료'])}")

    # ── 제휴사에만 있는 건 ──────────────────────────────
    if result["only_in_partner"]:
        print(f"\n【 제휴사 회신에만 존재 (DB 미등록) — {len(result['only_in_partner'])}건 】")
        for i, rec in enumerate(result["only_in_partner"], 1):
            print(f"\n  [{i}] 대출신청번호: {rec['대출신청번호']}")
            print(f"       신청일자  : {rec['제휴사_신청일자']}")
            print(f"       상품명    : {rec['제휴사_대출상품명']}")
            print(f"       대출금액  : {fmt_amount(rec['제휴사_대출금액'])}")
            print(f"       상태      : {rec['제휴사_상태']}")
            print(f"       지급수수료: {fmt_amount(rec['제휴사_지급수수료'])}")

    print(f"\n{SEP}\n")


# ─────────────────────────────────────────────
# 6. 메인
# ─────────────────────────────────────────────

def main():
    # 인자 처리
    if len(sys.argv) == 4:
        db_file = sys.argv[1]
        partner_file = sys.argv[2]
        partner_name = sys.argv[3]
    else:
        # 기본값 (개발/테스트용)
        base = os.path.dirname(os.path.abspath(__file__))
        db_file = os.path.join(base, "2604_F_기반_Agent_자료.xlsx")
        partner_file = os.path.join(base, "BNK저축은행_2604_대출.xlsx")
        partner_name = "BNK저축은행"
        print(f"[인자 없음] 기본값으로 실행: {partner_name}\n")

    # 파일 존재 확인
    for f in [db_file, partner_file]:
        if not os.path.exists(f):
            print(f"[오류] 파일을 찾을 수 없습니다: {f}")
            sys.exit(1)

    # 데이터 로드
    print(f"DB 파일 로딩 중: {os.path.basename(db_file)}")
    db_df = load_db(db_file, partner_name)
    print(f"  → {len(db_df)}건 로드 완료 ({partner_name})")

    # DB 신청번호 길이 자동 파악 (앞 14자가 기본, 실제 데이터 기준)
    db_key_len = int(db_df["대출신청번호"].dropna().astype(str).str.len().mode()[0]) if not db_df.empty else 14

    print(f"제휴사 파일 로딩 중: {os.path.basename(partner_file)}")
    partner_df = load_partner(partner_file, db_key_len=db_key_len)
    print(f"  → {len(partner_df)}건 로드 완료\n")

    # 검증 실행
    result = run_verification(db_df, partner_df)

    # 보고서 출력
    print_report(result, partner_name)


if __name__ == "__main__":
    main()

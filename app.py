"""
제휴사 정산 검증 웹앱
====================
Streamlit 기반 웹 인터페이스.
기존 정산_검증.py의 검증 로직을 그대로 재사용.
"""

import streamlit as st
import pandas as pd
import tempfile
import os
import io
from datetime import datetime

from 정산_검증 import (
    load_db,
    load_partner,
    run_verification,
    fmt_amount,
)

# ─────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="정산 검증 시스템",
    page_icon="📊",
    layout="wide",
)

# ─────────────────────────────────────────────
# 헬퍼 함수
# ─────────────────────────────────────────────

def save_uploaded_file(uploaded_file) -> str:
    """Streamlit UploadedFile → 임시 파일 경로 (xlsx)."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        return tmp.name


@st.cache_data(show_spinner=False)
def get_partner_names(file_bytes: bytes) -> list[str]:
    """DB 파일에서 제휴사명 목록을 추출 (캐시)."""
    bio = io.BytesIO(file_bytes)
    df = pd.read_excel(bio, sheet_name="대출시트", usecols=["MYDA_ORG_NM"])
    return sorted(df["MYDA_ORG_NM"].dropna().unique().tolist())


def extract_partner_name(prompt: str, known_names: list[str]) -> str | None:
    """
    프롬프트에서 제휴사명을 추출.
    1단계: 알려진 제휴사명 중 길이 역순으로 prompt 내 포함 여부 확인.
    2단계: 첫 번째 공백 토큰으로 fallback.
    """
    for name in sorted(known_names, key=len, reverse=True):
        if name in prompt:
            return name
    tokens = prompt.strip().split()
    return tokens[0] if tokens else None


def build_mismatch_table(mismatch_records: list) -> pd.DataFrame:
    """
    불일치 명세 → 대출신청번호 1건 = 1행.
    불일치 항목마다 'DB_항목명' / '제휴_항목명' 열 쌍으로 펼침.
    """
    rows = []
    for rec in mismatch_records:
        row: dict = {"대출신청번호": rec["대출신청번호"]}
        col_labels = []
        for item in rec["불일치항목"]:
            col = item["항목"]
            db_v = item["DB값"]
            pt_v = item["제휴사값"]
            if col in ["대출금액", "지급수수료"]:
                db_v = fmt_amount(db_v)
                pt_v = fmt_amount(pt_v)
            row[f"DB_{col}"] = str(db_v)
            row[f"제휴_{col}"] = str(pt_v)
            col_labels.append(col)
        row["불일치항목"] = ", ".join(col_labels)
        rows.append(row)

    if not rows:
        return pd.DataFrame(columns=["대출신청번호", "불일치항목"])

    df = pd.DataFrame(rows)

    # 열 순서: 대출신청번호 → 불일치항목 → (DB_X, 제휴_X) 쌍 반복
    fixed_cols = ["대출신청번호", "불일치항목"]
    value_cols = [c for c in df.columns if c not in fixed_cols]
    # DB_ / 제휴_ 쌍으로 정렬
    db_cols = [c for c in value_cols if c.startswith("DB_")]
    paired = []
    for db_col in db_cols:
        partner_col = db_col.replace("DB_", "제휴_")
        paired.append(db_col)
        if partner_col in value_cols:
            paired.append(partner_col)

    return df[fixed_cols + paired].fillna("-")


def build_only_db_table(records: list) -> pd.DataFrame:
    """DB에만 존재하는 건 목록 → DataFrame."""
    rows = []
    for rec in records:
        rows.append({
            "대출신청번호": rec["대출신청번호"],
            "신청일자": rec["DB_신청일자"],
            "상품명": rec["DB_대출상품명"],
            "대출금액": fmt_amount(rec["DB_대출금액"]),
            "상태": rec["DB_상태"],
            "지급수수료": fmt_amount(rec["DB_지급수수료"]),
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def build_only_partner_table(records: list) -> pd.DataFrame:
    """제휴사에만 존재하는 건 목록 → DataFrame."""
    rows = []
    for rec in records:
        rows.append({
            "대출신청번호": rec["대출신청번호"],
            "신청일자": rec["제휴사_신청일자"],
            "상품명": rec["제휴사_대출상품명"],
            "대출금액": fmt_amount(rec["제휴사_대출금액"]),
            "상태": rec["제휴사_상태"],
            "지급수수료": fmt_amount(rec["제휴사_지급수수료"]),
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ─────────────────────────────────────────────
# Sidebar — 입력 영역
# ─────────────────────────────────────────────

with st.sidebar:
    st.title("📊 정산 검증 시스템")
    st.caption("제휴사 대출 정산 대사 도구")
    st.markdown("---")

    db_file = st.file_uploader(
        "① DB 기반 자료 (플랫폼 총본)",
        type=["xlsx"],
        help="신한은행 DB에서 추출한 전체 제휴사 정산 파일",
    )

    partner_file = st.file_uploader(
        "② 제휴사 회신 파일",
        type=["xlsx"],
        help="제휴사가 회신한 전월 대출 취급 명세 파일",
    )

    st.markdown("---")

    prompt = st.text_area(
        "검증 요청 입력",
        placeholder="예: BNK저축은행 4월 정산 대사를 진행해줘",
        height=80,
        help="제휴사명이 포함된 자연어 문장을 입력하세요.",
    )

    ready = (db_file is not None) and (partner_file is not None) and bool(prompt.strip())
    run_btn = st.button("🔍 검증 실행", type="primary", disabled=not ready, use_container_width=True)

    if not ready:
        missing = []
        if db_file is None:
            missing.append("DB 파일")
        if partner_file is None:
            missing.append("제휴사 회신 파일")
        if not prompt.strip():
            missing.append("검증 요청 문장")
        st.caption(f"미입력: {', '.join(missing)}")


# ─────────────────────────────────────────────
# Main Area — 헤더
# ─────────────────────────────────────────────

st.title("제휴사 대출 정산 검증 결과")
st.caption(f"기준일: {datetime.today().strftime('%Y년 %m월 %d일')}")


# ─────────────────────────────────────────────
# 검증 실행 로직
# ─────────────────────────────────────────────

if run_btn:
    # 제휴사명 추출
    known_names = get_partner_names(db_file.getvalue())
    partner_name = extract_partner_name(prompt, known_names)

    if not partner_name:
        st.error("프롬프트에서 제휴사명을 인식하지 못했습니다. 제휴사명을 앞부분에 입력해주세요.")
        st.stop()

    db_path = None
    partner_path = None

    with st.spinner(f"**{partner_name}** 정산 데이터 로딩 및 검증 중..."):
        try:
            db_path = save_uploaded_file(db_file)
            partner_path = save_uploaded_file(partner_file)

            db_df = load_db(db_path, partner_name)

            db_key_len = (
                int(db_df["대출신청번호"].dropna().astype(str).str.len().mode()[0])
                if not db_df.empty
                else 14
            )

            partner_df = load_partner(partner_path, db_key_len=db_key_len)
            result = run_verification(db_df, partner_df)

            st.session_state["result"] = result
            st.session_state["partner_name"] = partner_name

        except ValueError as e:
            st.error(str(e))
            st.stop()
        except Exception as e:
            st.error(f"오류 발생: {e}")
            st.stop()
        finally:
            if db_path and os.path.exists(db_path):
                os.unlink(db_path)
            if partner_path and os.path.exists(partner_path):
                os.unlink(partner_path)


# ─────────────────────────────────────────────
# 결과 표시 (session_state 기반)
# ─────────────────────────────────────────────

if "result" in st.session_state:
    result: dict = st.session_state["result"]
    partner_name: str = st.session_state["partner_name"]

    total_mismatch = (
        len(result["mismatch_records"])
        + len(result["only_in_db"])
        + len(result["only_in_partner"])
    )
    matched_count = len(result["matched"])

    st.info(f"검출된 제휴사: **{partner_name}**")

    # ── 요약 메트릭 ──
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("DB 자료 건수", f"{result['db_total']}건")
    col2.metric("제휴사 회신 건수", f"{result['partner_total']}건")
    col3.metric(
        "완전 일치",
        f"{matched_count}건",
        delta=f"+{matched_count}" if matched_count > 0 else None,
        delta_color="normal",
    )
    col4.metric(
        "불일치 합계",
        f"{total_mismatch}건",
        delta=f"-{total_mismatch}" if total_mismatch > 0 else None,
        delta_color="inverse",
    )

    st.markdown("---")

    # ── 완전 일치 ──
    if total_mismatch == 0:
        st.success("모든 항목이 일치합니다.")
    else:
        # ── 불일치 상세 탭 ──
        tab1, tab2, tab3 = st.tabs([
            f"📋 값 불일치 ({len(result['mismatch_records'])}건)",
            f"🔴 DB에만 존재 ({len(result['only_in_db'])}건)",
            f"🟡 제휴사에만 존재 ({len(result['only_in_partner'])}건)",
        ])

        with tab1:
            if result["mismatch_records"]:
                st.caption("양쪽 파일에 모두 있으나 값이 다른 건 — 대출신청번호 1건 = 1행, DB / 제휴 값 비교")
                df_mis = build_mismatch_table(result["mismatch_records"])

                # DB_ 열은 파란색, 제휴_ 열은 주황색 헤더로 구분
                db_cols   = [c for c in df_mis.columns if c.startswith("DB_")]
                pt_cols   = [c for c in df_mis.columns if c.startswith("제휴_")]

                col_config = {
                    "대출신청번호": st.column_config.TextColumn("대출신청번호", width="medium"),
                    "불일치항목":   st.column_config.TextColumn("불일치 항목", width="medium"),
                }
                for c in db_cols:
                    col_config[c] = st.column_config.TextColumn(c, width="small")
                for c in pt_cols:
                    col_config[c] = st.column_config.TextColumn(c, width="small")

                st.dataframe(df_mis, use_container_width=True, hide_index=True, column_config=col_config)
            else:
                st.info("값 불일치 항목 없음")

        with tab2:
            if result["only_in_db"]:
                st.caption("DB에는 있으나 제휴사 회신 파일에 누락된 건")
                df_db_only = build_only_db_table(result["only_in_db"])
                st.dataframe(df_db_only, use_container_width=True, hide_index=True)
            else:
                st.info("DB에만 존재하는 건 없음")

        with tab3:
            if result["only_in_partner"]:
                st.caption("제휴사 회신 파일에는 있으나 DB에 미등록된 건")
                df_pt_only = build_only_partner_table(result["only_in_partner"])
                st.dataframe(df_pt_only, use_container_width=True, hide_index=True)
            else:
                st.info("제휴사에만 존재하는 건 없음")

else:
    # 초기 안내 화면
    st.markdown(
        """
        ### 사용 방법

        1. 왼쪽 사이드바에서 **① DB 기반 자료** (플랫폼 총본 xlsx)를 업로드하세요.
        2. **② 제휴사 회신 파일** (제휴사 대출 명세 xlsx)을 업로드하세요.
        3. 검증 요청 문장을 입력하세요.
           > 예: `BNK저축은행 4월 정산 대사를 진행해줘`
        4. **🔍 검증 실행** 버튼을 누르면 결과가 여기에 표시됩니다.

        ---

        ### 검증 항목

        | 항목 | 설명 |
        |------|------|
        | 신청일자 | 대출 신청 일자 일치 여부 |
        | 실행일자 | 대출 실행 일자 일치 여부 |
        | 대출상품코드 | 상품 코드 일치 여부 |
        | 대출상품명 | 상품명 일치 여부 |
        | 대출금액 | 금액 일치 여부 (1원 미만 오차 허용) |
        | 상태 | 실행/철회 등 상태 일치 여부 |
        | 지급수수료 | 수수료 일치 여부 (1원 미만 오차 허용) |
        """,
        unsafe_allow_html=False,
    )

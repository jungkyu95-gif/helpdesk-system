import streamlit as st
import pandas as pd
import os
from modules.data_processor import process_uploaded_file, load_from_db

st.set_page_config(
    page_title="IT Helpdesk 관리 시스템",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 사이드바 ──────────────────────────────────────────────
with st.sidebar:
    st.title("🔧 IT Helpdesk")
    st.markdown("---")

    st.subheader("📂 데이터 업로드")
    uploaded = st.file_uploader(
        "관리대장 Excel 업로드",
        type=["xlsx", "xls"],
        help="관리대장 시트가 포함된 Excel 파일을 업로드하세요.",
    )

    if uploaded:
        with st.spinner("파일 처리 중..."):
            try:
                df = process_uploaded_file(uploaded)
                st.session_state["df"] = df
                st.success(f"✅ 로드 완료! ({len(df):,}건)")
            except Exception as e:
                st.error(f"❌ {str(e)}")

    # 샘플 데이터 로드
    sample_path = "sample_data/관리대장_더미데이터.xlsx"
    if st.button("📋 샘플 데이터 사용", use_container_width=True):
        if os.path.exists(sample_path):
            with st.spinner("샘플 데이터 로딩 중..."):
                try:
                    with open(sample_path, "rb") as f:
                        df = process_uploaded_file(f)
                    st.session_state["df"] = df
                    st.success(f"✅ 샘플 데이터 로드! ({len(df):,}건)")
                except Exception as e:
                    st.error(f"❌ {str(e)}")
        else:
            st.warning("샘플 데이터 파일이 없습니다.")

    st.markdown("---")
    st.markdown("**📖 사용 가이드**")
    st.markdown("""
1. Excel 파일 업로드 또는 샘플 데이터 선택
2. **🔍 검색** 탭에서 오류 검색
3. **📊 보고서** 탭에서 자동 보고서 생성
""")
    st.markdown("---")
    st.caption("IT Helpdesk 관리 시스템 v1.0")

# ── 메인 ──────────────────────────────────────────────────
st.title("🔧 IT Helpdesk 관리 시스템")
st.markdown("Excel 관리대장 자동 분석 · 오류 검색 · 보고서 생성")
st.markdown("---")

df = st.session_state.get("df")

if df is None:
    # 환영 화면
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("**🔍 오류 검색**\n\nNAC, VPN 등 키워드로 과거 사례와 해결방안을 즉시 검색")
    with col2:
        st.info("**📊 자동 보고서**\n\n월간 처리 현황, 시스템별 통계, 추이 분석 자동 생성")
    with col3:
        st.info("**⬇️ Excel 다운로드**\n\n보고서를 Excel 파일로 다운로드하여 공유")
    st.markdown("---")
    st.markdown("### 👆 왼쪽 사이드바에서 파일을 업로드하거나 샘플 데이터를 선택하세요!")
else:
    # 데이터 로드된 상태 - 현황 요약
    total = len(df)
    completed = len(df[df['진행상태'] == '1차 완료']) if '진행상태' in df.columns else 0
    rate = completed / total if total > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📋 총 접수", f"{total:,}건")
    col2.metric("✅ 1차 완료", f"{completed:,}건")
    col3.metric("📈 처리율", f"{rate:.1%}")
    col4.metric("🏢 고객사", f"{df['회사명'].nunique() if '회사명' in df.columns else '-'}개")

    st.markdown("---")
    st.markdown("### 👈 왼쪽 메뉴에서 **🔍 검색** 또는 **📊 보고서** 페이지로 이동하세요.")

    # 최근 데이터 미리보기
    with st.expander("📄 데이터 미리보기 (최근 10건)"):
        preview_cols = [c for c in ['접수번호', '통화시작시간', '회사명', '요청유형', '시스템/장비명', '조치유형', '진행상태', '처리소요시간'] if c in df.columns]
        st.dataframe(df[preview_cols].head(10), use_container_width=True)

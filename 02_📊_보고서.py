import streamlit as st
import pandas as pd
from modules.report_generator import analyze_data, generate_report
from datetime import datetime

st.set_page_config(page_title="자동 보고서", page_icon="📊", layout="wide")
st.title("📊 자동 보고서 생성")
st.markdown("데이터를 분석하여 월간 보고서를 자동으로 생성합니다.")
st.markdown("---")

df = st.session_state.get("df")

if df is None:
    st.warning("⚠️ 데이터가 없습니다. 메인 페이지에서 파일을 업로드해 주세요.")
    st.stop()

# ── 필터 옵션 ─────────────────────────────────────────
with st.expander("⚙️ 보고서 필터 설정", expanded=False):
    col1, col2 = st.columns(2)
    with col1:
        companies = ['전체'] + sorted(df['회사명'].dropna().unique().tolist()) if '회사명' in df.columns else ['전체']
        selected_company = st.selectbox("고객사 선택", companies)
    with col2:
        statuses = ['전체'] + sorted(df['진행상태'].dropna().unique().tolist()) if '진행상태' in df.columns else ['전체']
        selected_status = st.selectbox("진행상태 필터", statuses)

df_filtered = df.copy()
if selected_company != '전체' and '회사명' in df.columns:
    df_filtered = df_filtered[df_filtered['회사명'] == selected_company]
if selected_status != '전체' and '진행상태' in df.columns:
    df_filtered = df_filtered[df_filtered['진행상태'] == selected_status]

st.markdown(f"분석 대상: **{len(df_filtered):,}건**")

# ── 분석 실행 ─────────────────────────────────────────
if st.button("📈 보고서 생성", type="primary", use_container_width=False):
    with st.spinner("데이터 분석 중..."):
        analysis = analyze_data(df_filtered)

    st.success("✅ 보고서 생성 완료!")
    st.markdown("---")

    # ── 핵심 지표 ──────────────────────────────────────
    st.subheader("① 처리 현황 요약")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📋 총 접수", f"{analysis['total_tickets']:,}건")
    col2.metric("✅ 1차 완료", f"{analysis['completed_count']:,}건")
    col3.metric("📈 처리율", f"{analysis['completion_rate']:.1%}")
    col4.metric("⏱ 평균 처리시간", analysis['avg_resolution_str'])
    st.markdown("---")

    # ── 시스템별 현황 ──────────────────────────────────
    st.subheader("② 시스템/장비별 접수 현황")
    if analysis['error_distribution']:
        sys_df = pd.DataFrame(list(analysis['error_distribution'].items()), columns=['시스템/장비명', '접수건수'])
        sys_df = sys_df.sort_values('접수건수', ascending=False)
        col1, col2 = st.columns([1, 1])
        with col1:
            st.dataframe(sys_df.reset_index(drop=True), use_container_width=True, hide_index=True)
        with col2:
            st.bar_chart(sys_df.set_index('시스템/장비명')['접수건수'])
    st.markdown("---")

    # ── 조치유형 현황 ──────────────────────────────────
    st.subheader("③ 조치유형별 현황")
    if analysis['action_distribution']:
        act_df = pd.DataFrame(list(analysis['action_distribution'].items()), columns=['조치유형', '건수'])
        act_df = act_df.sort_values('건수', ascending=False)
        col1, col2 = st.columns([1, 1])
        with col1:
            st.dataframe(act_df.reset_index(drop=True), use_container_width=True, hide_index=True)
        with col2:
            st.bar_chart(act_df.set_index('조치유형')['건수'])
    st.markdown("---")

    # ── 월별 추이 ──────────────────────────────────────
    st.subheader("④ 월별 접수 추이")
    if analysis['monthly_trend']:
        trend_df = pd.DataFrame(list(analysis['monthly_trend'].items()), columns=['월', '건수'])
        trend_df = trend_df.sort_values('월')
        st.line_chart(trend_df.set_index('월')['건수'])
    st.markdown("---")

    # ── 고객사별 현황 ──────────────────────────────────
    if analysis['company_distribution']:
        st.subheader("⑤ 고객사별 현황")
        comp_df = pd.DataFrame(list(analysis['company_distribution'].items()), columns=['고객사', '건수'])
        comp_df = comp_df.sort_values('건수', ascending=False)
        col1, col2 = st.columns([1, 1])
        with col1:
            st.dataframe(comp_df.reset_index(drop=True), use_container_width=True, hide_index=True)
        with col2:
            st.bar_chart(comp_df.set_index('고객사')['건수'])
        st.markdown("---")

    # ── Excel 다운로드 ──────────────────────────────────
    st.subheader("⬇️ 보고서 다운로드")
    with st.spinner("Excel 파일 생성 중..."):
        excel_buf = generate_report(df_filtered)

    filename = f"IT_Helpdesk_보고서_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        label="📥 Excel 보고서 다운로드",
        data=excel_buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    # 버튼 클릭 전 간단 현황
    st.info("👆 **보고서 생성** 버튼을 클릭하면 분석이 시작됩니다.")
    st.markdown("---")
    st.subheader("📋 데이터 미리보기")
    preview_cols = [c for c in ['접수번호', '회사명', '요청유형', '시스템/장비명', '조치유형', '진행상태', '처리소요시간'] if c in df_filtered.columns]
    st.dataframe(df_filtered[preview_cols].head(10), use_container_width=True, hide_index=True)

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

st.set_page_config(page_title="오류 검색", page_icon="🔍", layout="wide")
st.title("🔍 오류 검색")
st.markdown("키워드로 과거 IT 헬프데스크 사례와 해결방안을 검색합니다.")
st.markdown("---")

df = st.session_state.get("df")

if df is None:
    st.warning("⚠️ 데이터가 없습니다. 메인 페이지에서 파일을 업로드해 주세요.")
    st.stop()

# ── 검색 UI ────────────────────────────────────────────
col1, col2, col3 = st.columns([3, 1, 1])

with col1:
    keyword = st.text_input(
        "🔎 검색어 입력",
        placeholder="예: NAC, VPN, 팀즈, 계정, 프린터...",
        help="요청유형, 시스템/장비명, 조치내용에서 검색합니다."
    )

with col2:
    match_type = st.radio("검색 방식", ["부분 일치", "정확 일치"], horizontal=True)
    match_code = 'partial' if match_type == "부분 일치" else 'exact'

with col3:
    sort_col = st.selectbox("정렬 기준", ["통화시작시간", "처리소요시간", "회사명"])
    sort_order = st.radio("정렬 순서", ["최신순", "오래된순"], horizontal=True)
    sort_asc = 'asc' if sort_order == "오래된순" else 'desc'

# 빠른 선택 버튼
st.markdown("**⚡ 빠른 검색:**")
quick_terms = ['NAC', 'VPN', '팀즈', 'NERP', '장비', '계정', 'PC관리', '기타']
cols = st.columns(len(quick_terms))
for i, term in enumerate(quick_terms):
    if cols[i].button(term, use_container_width=True):
        keyword = term

search_btn = st.button("🔍 검색", type="primary", use_container_width=False)
st.markdown("---")

# ── 검색 실행 ───────────────────────────────────────────
if keyword:
    with st.spinner(f"'{keyword}' 검색 중..."):
        results = search_symptoms(keyword, df, match_type=match_code)
        results = sort_results(results, sort_by=sort_col, order=sort_asc)

    if results.empty:
        st.info(f"🔍 '{keyword}'에 대한 검색 결과가 없습니다.")
    else:
        st.success(f"✅ **{len(results):,}건** 검색됨 (검색어: '{keyword}')")

        # 통계 요약
        col1, col2, col3 = st.columns(3)
        completed = len(results[results['진행상태'] == '1차 완료']) if '진행상태' in results.columns else 0
        col1.metric("검색 결과", f"{len(results):,}건")
        col2.metric("1차 완료", f"{completed:,}건")
        col3.metric("처리율", f"{completed/len(results):.1%}" if len(results) > 0 else "-")

        # 결과 테이블
        display_cols = [c for c in [
            '접수번호', '통화시작시간', '회사명', '요청유형', '시스템/장비명',
            '조치유형', '조치내용', '처리자', '처리소요시간', '진행상태'
        ] if c in results.columns]

        st.dataframe(
            results[display_cols].reset_index(drop=True),
            use_container_width=True,
            height=400,
        )

        # CSV 다운로드
        csv = results[display_cols].to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="⬇️ 검색 결과 CSV 다운로드",
            data=csv.encode('utf-8-sig'),
            file_name=f"검색결과_{keyword}.csv",
            mime="text/csv",
        )

        # 조치내용 상세 보기
        with st.expander("📋 조치내용 상세 보기"):
            for _, row in results.head(20).iterrows():
                sys_name = row.get('시스템/장비명', '-')
                req_type = row.get('요청유형', '-')
                action = row.get('조치내용', '-')
                handler = row.get('처리자', '-')
                st.markdown(f"**[{sys_name}] {req_type}**  \n처리자: {handler}  \n조치: {action}")
                st.markdown("---")
else:
    # 검색 전 통계 현황
    st.markdown("### 📊 전체 데이터 현황")
    if '시스템/장비명' in df.columns:
        sys_counts = df['시스템/장비명'].value_counts().head(10).reset_index()
        sys_counts.columns = ['시스템/장비명', '접수건수']
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**시스템/장비별 Top 10**")
            st.dataframe(sys_counts, use_container_width=True, hide_index=True)
        with col2:
            if '조치유형' in df.columns:
                action_counts = df['조치유형'].value_counts().reset_index()
                action_counts.columns = ['조치유형', '건수']
                st.markdown("**조치유형별 현황**")
                st.dataframe(action_counts, use_container_width=True, hide_index=True)

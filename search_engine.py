import pandas as pd


def search_symptoms(keyword, df, match_type='partial', search_cols=None):
    """
    오류증상 / 조치내용 검색
    match_type: 'partial' (부분일치) | 'exact' (정확일치)
    """
    if not keyword or df is None or df.empty:
        return pd.DataFrame()

    if search_cols is None:
        search_cols = ['요청유형', '시스템/장비명', '조치내용', '조치유형']

    existing_cols = [c for c in search_cols if c in df.columns]
    if not existing_cols:
        return pd.DataFrame()

    if match_type == 'exact':
        mask = df[existing_cols].apply(
            lambda col: col.astype(str).str.lower() == keyword.lower()
        ).any(axis=1)
    else:
        mask = df[existing_cols].apply(
            lambda col: col.astype(str).str.contains(keyword, case=False, na=False)
        ).any(axis=1)

    return df[mask].copy()


def get_unique_symptoms(df):
    results = set()
    for col in ['요청유형', '시스템/장비명']:
        if col in df.columns:
            results.update(df[col].dropna().unique().tolist())
    return sorted(results)


def get_solution_for_symptom(symptom, df):
    if '요청유형' not in df.columns:
        return pd.DataFrame()
    result = df[df['요청유형'].astype(str).str.contains(symptom, case=False, na=False)]
    return result[['요청유형', '시스템/장비명', '조치유형', '조치내용', '처리자', '처리소요시간', '진행상태']].copy()


def sort_results(df, sort_by='통화시작시간', order='desc'):
    if df.empty or sort_by not in df.columns:
        return df
    ascending = (order == 'asc')
    try:
        return df.sort_values(by=sort_by, ascending=ascending, na_position='last')
    except:
        return df

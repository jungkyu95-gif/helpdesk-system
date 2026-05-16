import pandas as pd
import sqlite3
import os
from utils import parse_duration_to_seconds

REQUIRED_COLUMNS = ['접수번호', '회사명', '요청유형', '시스템/장비명', '조치유형', '조치내용', '진행상태', '처리소요시간']
DB_PATH = "helpdesk.db"


def read_excel_file(file_path_or_buffer):
    df = pd.read_excel(file_path_or_buffer, sheet_name='관리대장', dtype=str)
    df.columns = df.columns.str.strip()
    return df


def validate_data(df):
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        return False, f"필수 컬럼 누락: {', '.join(missing)}"
    if len(df) == 0:
        return False, "데이터가 없습니다."
    return True, "검증 통과"


def clean_data(df):
    df = df.copy()
    # 공백 제거
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.strip()
    # 날짜 컬럼 처리
    for date_col in ['통화시작시간', '완료일시']:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    # 처리소요시간 → 초 변환 (분석용)
    if '처리소요시간' in df.columns:
        df['처리소요초'] = df['처리소요시간'].apply(parse_duration_to_seconds)
    # 빈 문자열 → NaN
    df.replace('', pd.NA, inplace=True)
    return df


def save_to_db(df, db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    # datetime 컬럼 문자열 변환 후 저장
    df_to_save = df.copy()
    for col in df_to_save.select_dtypes(include=['datetime64[ns]', 'datetimetz']).columns:
        df_to_save[col] = df_to_save[col].astype(str)
    df_to_save.to_sql('helpdesk_records', conn, if_exists='replace', index=False)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_company ON helpdesk_records(회사명)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_status ON helpdesk_records(진행상태)")
    conn.commit()
    conn.close()
    return True


def load_from_db(db_path=DB_PATH):
    if not os.path.exists(db_path):
        return None
    conn = sqlite3.connect(db_path)
    df = pd.read_sql("SELECT * FROM helpdesk_records", conn)
    conn.close()
    return df


def process_uploaded_file(uploaded_file):
    """업로드 파일 전체 파이프라인: 읽기 → 검증 → 정제 → DB 저장"""
    df_raw = read_excel_file(uploaded_file)
    ok, msg = validate_data(df_raw)
    if not ok:
        raise ValueError(msg)
    df_clean = clean_data(df_raw)
    save_to_db(df_clean)
    return df_clean

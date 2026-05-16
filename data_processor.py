import pandas as pd

REQUIRED_COLUMNS = ['접수번호', '회사명', '요청유형', '진행상태']


def read_excel_file(file_path_or_buffer):
    try:
        df = pd.read_excel(file_path_or_buffer, sheet_name='관리대장', dtype=str)
    except Exception:
        try:
            df = pd.read_excel(file_path_or_buffer, sheet_name=0, dtype=str)
        except Exception as e:
            raise ValueError(f"Excel 파일을 읽을 수 없습니다: {str(e)}")
    df.columns = df.columns.str.strip()
    return df


def validate_data(df):
    existing = [col for col in REQUIRED_COLUMNS if col in df.columns]
    if len(existing) == 0:
        return False, f"올바른 관리대장 파일이 아닙니다. 컬럼: {list(df.columns)}"
    if len(df) == 0:
        return False, "데이터가 없습니다."
    return True, "검증 통과"


def clean_data(df):
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace('nan', '')
            df[col] = df[col].replace('None', '')
    for date_col in ['통화시작시간', '완료일시']:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    if '처리소요시간' in df.columns:
        df['처리소요초'] = df['처리소요시간'].apply(_parse_duration)
    return df


def _parse_duration(duration_str):
    try:
        parts = str(duration_str).strip().split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(float(parts[2]))
        elif len(parts) == 2:
            return int(parts[0]) * 60 + int(float(parts[1]))
    except Exception:
        pass
    return None


def process_uploaded_file(uploaded_file):
    df_raw = read_excel_file(uploaded_file)
    ok, msg = validate_data(df_raw)
    if not ok:
        raise ValueError(msg)
    df_clean = clean_data(df_raw)
    return df_clean


def load_from_db():
    return None

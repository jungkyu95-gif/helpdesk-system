import pandas as pd
from datetime import datetime


def format_time(seconds):
    if seconds is None or pd.isna(seconds):
        return "-"
    try:
        seconds = int(seconds)
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"
    except:
        return str(seconds)


def format_date(date_obj):
    if date_obj is None or pd.isna(date_obj):
        return "-"
    try:
        if isinstance(date_obj, str):
            return date_obj
        return date_obj.strftime("%Y-%m-%d %H:%M")
    except:
        return str(date_obj)


def get_error_categories(df):
    if '요청유형' not in df.columns:
        return []
    return sorted(df['요청유형'].dropna().unique().tolist())


def get_system_categories(df):
    if '시스템/장비명' not in df.columns:
        return []
    return sorted(df['시스템/장비명'].dropna().unique().tolist())


def parse_duration_to_seconds(duration_str):
    """'HH:MM:SS' 문자열을 초로 변환"""
    if not duration_str or pd.isna(duration_str):
        return None
    try:
        parts = str(duration_str).strip().split(':')
        if len(parts) == 3:
            h, m, s = int(parts[0]), int(parts[1]), int(float(parts[2]))
            return h * 3600 + m * 60 + s
        elif len(parts) == 2:
            m, s = int(parts[0]), int(float(parts[1]))
            return m * 60 + s
    except:
        pass
    return None


def handle_error(error_message):
    return f"오류가 발생했습니다: {str(error_message)}"

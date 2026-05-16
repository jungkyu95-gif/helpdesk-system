# 🔧 IT Helpdesk 관리 시스템

**포트폴리오 프로젝트 | 더미 데이터 사용**

## 기능

- ✅ Excel 관리대장 자동 분석
- ✅ 오류증상 검색 기능 (NAC, VPN, 팀즈 등)
- ✅ 자동 보고서 생성 (그래프, 표, 통계)
- ✅ Excel 보고서 다운로드

## 기술 스택

- **Frontend**: Streamlit
- **Data Processing**: pandas, openpyxl
- **Deployment**: Streamlit Cloud

## 웹에서 바로 사용

> (배포 후 URL을 여기에 입력하세요)

## 로컬에서 실행

```bash
git clone https://github.com/your-name/helpdesk-system.git
cd helpdesk-system
pip install -r requirements.txt
streamlit run main.py
```

## 프로젝트 구조

```
├── main.py
├── pages/
│   ├── 01_🔍_검색.py
│   └── 02_📊_보고서.py
├── modules/
│   ├── data_processor.py
│   ├── search_engine.py
│   ├── report_generator.py
│   └── utils.py
└── sample_data/
    └── 관리대장_더미데이터.xlsx
```

## 주의사항

이 프로젝트는 실제 데이터에서 민감 정보를 제거한 더미 데이터를 사용합니다.
회사명, 직원명, 사번, 연락처는 모두 익명화 처리되었습니다.

## 라이선스

MIT License

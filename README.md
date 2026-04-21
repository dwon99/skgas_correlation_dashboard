# Fundamental Driver Correlation Dashboard (v2)

SK Gas Trading AX — Position Play AI

## v2 변경사항

### 1. 윈도우 설정: 연-월-일 지원
- **연도별 커스텀 월-일**: 연도 복수 선택 + 시작/종료 월-일 설정 (예: 매년 4/1~11/30)
- **자유 윈도우**: 개별 윈도우마다 시작/종료 날짜를 연-월-일 단위로 직접 설정

### 2. 월물 설정: Y+a년 b월 형식
- 기존: 드롭다운에서 "2019-01" 선택
- 변경: Y+a (연도 오프셋: 0,1,2,3) + b (월: 1~12) 조합
- 예시: Y+0 2월 → 해당 연도 2월물, Y+1 3월 → 다음 연도 3월물
- 연도별 윈도우 분석 시 각 연도에 맞는 월물이 자동 계산

### 3. Fundamental Driver: 인덱스별 그룹핑 + 복수 선택
- 인덱스(FEI C3, MB C3, TTF 등)를 먼저 선택
- 해당 인덱스에 속하는 Driver만 표시 (예: [FEI C3] PDH operating rate)
- 복수 인덱스 × 복수 Driver 조합으로 분석 가능

## 설치 및 실행

```bash
# 1. 가상환경 생성 (권장)
python -m venv venv
source venv/bin/activate  # Mac/Linux
# venv\Scripts\activate   # Windows

# 2. 패키지 설치
pip install -r requirements.txt

# 3. 실행
streamlit run app.py
# 또는
python -m streamlit run app.py
```

브라우저에서 `http://localhost:8501`로 접속됩니다.

## 입력 데이터 형식

### 1. 인덱스 가격 데이터 (CSV/Excel)
| INDEX_ID | Index명 | 기준일자 | 월물 | 휴일여부 | Value |
|---|---|---|---|---|---|
| 1020012 | MB_NON_TET_C3 | 2019-01-09 | 2019-01 | N | 65.875 |

### 2. Fundamental Driver 데이터 (CSV/Excel)
| Index | Fundamental Driver | Date | Value |
|---|---|---|---|
| FEI C3 | LPG 수출량 | 2016-01-01 | 175559.365 |

날짜 형식은 자동 인식됩니다.

# KT 정산 자동화 도구

KT커머스 드롭십 정산 업무를 자동화하는 Python 도구입니다.

## 파일 구성

| 파일 | 설명 |
|------|------|
| `run_settlement.py` | 정산 핵심 로직 (SettlementRunner) |
| `settlement_gui.py` | Tkinter GUI 래퍼 |
| `build_exe.py` | PyInstaller 빌드 스크립트 |

## 실행 방법

### GUI 실행
```bash
python settlement_gui.py
```

### EXE 빌드
```bash
python build_exe.py
```
빌드 결과물: `dist/정산자동화/정산자동화.exe`

## 필요 파일 (별도 관리, Git 미포함)

- `중견기업목록.xlsx` — 중견기업 명단 및 사업자번호
- `대기업목록.xlsx` — 대기업 명단
- `mapping_master.xlsx` — 협력사 매핑 마스터

## 주요 기능

- KT 매입/반품 데이터와 플랫폼 데이터 대사
- 과세/면세 구분 자동 검증
- 어음 지급 대상 판별 (대기업 90일 / 중견기업 60일)
- 반품 주문 자동 매칭 + return_mapping.xlsx 후보 제안
- 결과 Excel 자동 생성

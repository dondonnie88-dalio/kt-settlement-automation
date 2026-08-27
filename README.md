# kt-settlement-automation

업무 자동화 스크립트 모음. 세 개의 독립된 하위 프로젝트로 구성되어 있습니다.

## 📁 구조

### 정산 자동화 (루트)
- `run_settlement.py` — KT 정산 자동화 시스템 (인보이스 / 매출현황 / 어음확인 / 자금예측 시트 생성)
- `run_reconcile.py` — KT ↔ 플랫폼 정산 대사 스크립트
- `make_sample_data.py` — 테스트용 샘플 데이터 생성
- `run_정산자동화.bat`, `run_매출분석.bat` — 실행용 배치 파일

### `delivery_eta/` — 배송 예측
상품코드 기준 "N일 이내 도착확률" 예측 파이프라인. EDA부터 리스크 점검, 확률분포
학습, 캘리브레이션 검증, 협력사/카테고리 이상탐지, 서빙 API까지 단계별 스크립트
(`01_eda.py` ~ `12_procurement_delay_flag.py`)로 구성되어 있습니다.
- `run_all.py` / `run_all.bat` — 전체 파이프라인 순차 실행
- `serve.py` — 예측 서빙 로직 (`DeliveryPredictor`)
- `gui_predict.py` / `gui_predict.bat` — 간단 조회 GUI
- `API_SPEC.md` — 통합플랫폼 연동용 API 스펙
- `known_risks.md`, `phase3_backlog.md` — 검증된 리스크 기록 / 백로그
- 자세한 재학습 체크리스트는 `known_risks.md` 상단 참고

### `보고자료/` — 보고서 생성
분석 결과를 보고용 문서(docx/xlsx)로 변환하는 스크립트 모음
(경영진 요약, AX 외주사 브리핑, 변경 이력 등).

## 참고
각 하위 프로젝트가 생성하는 결과 파일(csv/report txt/docx/xlsx 등)은 `.gitignore`에서
제외되어 있으며, 소스 스크립트만 버전 관리됩니다.

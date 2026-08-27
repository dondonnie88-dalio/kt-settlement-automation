# -*- coding: utf-8 -*-
"""배송 예측 프로젝트 전체 파이프라인을 순서대로 재실행한다. run_all.bat 더블클릭으로 실행."""
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}

STEPS = [
    ("01_eda.py", "Phase 0: 데이터 적합성 EDA"),
    ("02_risk_check.py", "리스크 점검 (주문시배송리드타임 정체 / 외부사 특이성 / 직접배송 비교)"),
    ("03_distribution_fit.py", "Phase 1: 확률분포 학습 + 캘리브레이션 검증"),
    ("04_recency_diagnosis.py", "재학습 전 진단 (2~5일 과소신 원인 규명)"),
    ("05_growth_vendor_flag.py", "급성장 협력사x상품군 탐지 (재발 방지 모니터링)"),
    ("06_category_granularity_check.py", "중분류 세분화 가치 검증"),
    ("07_calendar_effects_check.py", "전사적 배송 지연 이상일 탐지 (공휴일/파업/기상 후보)"),
    ("08_holiday_adjustment.py", "공휴일 연휴 직전 보정계수 산출 및 검증"),
    ("09_batch_import_vendor_check.py", "Phase 3: 배치성 발주 협력사 탐지 (참고용, 10번으로 대체됨)"),
    ("10_kraljic_delivery_risk.py", "Phase 3: 배송 리스크 Kraljic 매트릭스 분류 (상품코드 단위, 정식)"),
    ("11_sla_gap_check.py", "표준납기일(SLA) vs 모델 예측 불일치 점검"),
    ("12_procurement_delay_flag.py", "발주->출하 지연 협력사x상품군x월 조기경보 (협력사 배송리드타임 관리용)"),
    ("build_analysis_workbook.py", "분석 심화 엑셀 워크북 생성"),
]

OUTPUTS = [
    "eda_report.txt", "risk_check_report.txt", "phase1_report.txt",
    "recency_diagnosis.txt", "growth_vendor_flags.txt", "category_granularity_check.txt",
    "calendar_anomaly_report.txt", "holiday_adjustment_report.txt", "holiday_adjustment_ratios.csv",
    "phase3_batch_vendor_candidates.txt", "kraljic_delivery_risk.csv", "kraljic_delivery_risk_report.txt",
    "sla_gap_report.txt", "procurement_delay_flags.txt",
    "segment_distributions.csv", "용어정리.xlsx", "배송예측_분석보고서.xlsx",
]


def main():
    print("=" * 70)
    print("배송 예측 프로젝트 - 전체 파이프라인 재실행")
    print("첫 실행이면 원본 엑셀 파싱 때문에 10~20분 걸릴 수 있습니다.")
    print("한 번 실행한 뒤로는 cache 폴더에 저장된 파케이 파일을 재사용해서 훨씬 빨라집니다.")
    print("=" * 70)

    for script, desc in STEPS:
        print(f"\n[{script}] {desc}")
        result = subprocess.run(
            [sys.executable, script],
            cwd=HERE,
            env=ENV,
        )
        if result.returncode != 0:
            print(f"  !! {script} 실행 중 오류 발생 (종료코드 {result.returncode}). 여기서 중단합니다.")
            input("\n아무 키나 누르면 창이 닫힙니다...")
            return
        print(f"  -> 완료")

    print("\n" + "=" * 70)
    print("전체 완료. 아래 결과 파일들을 확인하세요:")
    for f in OUTPUTS:
        exists = "✔" if (HERE / f).exists() else "(없음)"
        print(f"  {exists} {f}")
    print("=" * 70)
    input("\n아무 키나 누르면 창이 닫힙니다...")


if __name__ == "__main__":
    main()

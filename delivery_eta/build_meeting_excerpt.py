# -*- coding: utf-8 -*-
"""
박형근 차장님(통합플랫폼) 협의용 발췌본 - 배송예측_분석보고서.xlsx에서
연동 논의에 필요한 시트(요약/용어정리/캘리브레이션_검증/세그먼트_상세/표준납기일_비교/
Kraljic_리스크분류)만 남기고 나머지(진행과정/데이터현황/리스크점검/상품코드_조회예시/
급성장세그먼트)는 제거한 버전을 만든다. 세그먼트_상세는 3천여 개(구성 변경 때마다
증감 - 정확한 수는 phase1_report.txt 참고) 전체 상품코드/세그먼트가 Excel 표로
들어있어 필터/정렬로 원하는 상품코드를 바로 찾을 수 있음
(상품코드_조회예시의 10개 예시보다 실사용에 낫다는 피드백 반영, 2026-07-07). 표준납기일_비교는
API_SPEC.md 협의 필요 사항 9번(SLA vs 모델 예측 표시 정책)의 근거 자료라 포함
(2026-07-09 추가). Kraljic_리스크분류는 표준납기일_비교의 gap 큰 상품 대부분이 여기
전략재/병목재와 겹친다는 설명 근거라 포함(2026-07-09 추가).

전체 워크북(배송예측_분석보고서.xlsx)은 그대로 마스터/기록용으로 유지 - 이 스크립트는
파일을 복사한 뒤 불필요한 시트만 삭제하는 방식이라(재구성 아님), 서식·병합셀·차트가
원본 그대로 보존된다.
"""
import shutil
from pathlib import Path
from openpyxl import load_workbook

HERE = Path(__file__).parent
SOURCE = HERE / "배송예측_분석보고서.xlsx"
TARGET = HERE / "배송예측_협의용_통합플랫폼.xlsx"

KEEP_SHEETS = ["요약", "용어정리", "캘리브레이션_검증", "세그먼트_상세", "표준납기일_비교", "Kraljic_리스크분류"]


def main():
    shutil.copy(SOURCE, TARGET)
    wb = load_workbook(TARGET)

    for name in list(wb.sheetnames):
        if name not in KEEP_SHEETS:
            del wb[name]

    wb.save(TARGET)
    print(f"Saved: {TARGET.name}  (시트: {wb.sheetnames})")


if __name__ == "__main__":
    main()

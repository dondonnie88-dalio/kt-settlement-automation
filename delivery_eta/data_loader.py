# -*- coding: utf-8 -*-
"""배송예측 프로젝트 공용 데이터 로더. raw xlsx -> parquet 캐시."""
import pandas as pd
from pathlib import Path

DATA_DIR = Path(r"C:\Users\USER\OneDrive - KT Corporation\문서\배송예측")
CACHE_DIR = DATA_DIR / "delivery_eta" / "cache"

SEGMENTS = ["KT", "그룹사", "외부사", "지입자재"]

# 2026-07-10: "그룹사" 세그먼트가 2025년 이후 raw 파일에서 "KT그룹사"로 이름/범위가
# 수정됨(사용자 확인) - 2022~2024년은 여전히 기존 "그룹사" 파일만 존재하므로, 연도별로
# 실제 파일명이 다름. 내부 세그먼트 키("그룹사")는 그대로 유지하고(하위 코드 영향 없게)
# 파일명만 연도에 따라 다르게 찾는다.
FILENAME_OVERRIDES = {
    ("그룹사", 2025): "KT그룹사",
    ("그룹사", 2026): "KT그룹사",
}

USECOLS = [
    "주문번호", "주문일자", "발주일자", "출하일자", "배송완료일자",
    "배송예정일", "배송희망일", "표준납기일", "입고일자", "입고승인일자",
    "정산확정일", "배송지", "배송업체", "송장번호", "상품배송리드타임",
    "주문시배송리드타임", "배송상태", "배송형태", "배송유형", "상품타입",
    "상품유형", "협력사명", "서비스카테고리", "주문상태", "취소수량", "반품수량",
    "상품코드", "상품명",
]

DATE_COLS = [
    "주문일자", "발주일자", "출하일자", "배송완료일자", "배송예정일",
    "배송희망일", "표준납기일", "입고일자", "입고승인일자", "정산확정일",
]


def parse_date_col(s: pd.Series) -> pd.Series:
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    is_num = pd.to_numeric(s, errors="coerce")
    num_mask = is_num.notna()
    if num_mask.any():
        yyyymmdd = is_num[num_mask].astype("Int64")
        valid = (yyyymmdd >= 19000101) & (yyyymmdd <= 20991231)
        parsed = pd.to_datetime(
            yyyymmdd[valid].astype(str), format="%Y%m%d", errors="coerce"
        )
        out.loc[parsed.index] = parsed
    str_mask = ~num_mask & s.notna()
    if str_mask.any():
        parsed = pd.to_datetime(s[str_mask], errors="coerce")
        out.loc[parsed.index] = parsed
    return out


def load_segment_year(segment: str, year: int, force: bool = False) -> pd.DataFrame:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE_DIR / f"{segment}_{year}.parquet"
    if cache_path.exists() and not force:
        return pd.read_parquet(cache_path)

    file_segment = FILENAME_OVERRIDES.get((segment, year), segment)
    fname = f"{year}_주문현황_{file_segment}(요청자_이돈현).xlsx"
    path = DATA_DIR / fname
    if not path.exists():
        raise FileNotFoundError(path)
    print(f"  [raw parse] {fname} ...", flush=True)
    df = pd.read_excel(path, usecols=USECOLS, engine="openpyxl")
    for c in DATE_COLS:
        df[c] = parse_date_col(df[c])
    df["_segment"] = segment
    df["_year"] = year
    df.to_parquet(cache_path)
    print(f"  [cached] {cache_path.name} ({len(df):,} rows)", flush=True)
    return df


def load_many(segments, years) -> pd.DataFrame:
    dfs = []
    for seg in segments:
        for y in years:
            try:
                dfs.append(load_segment_year(seg, y))
            except FileNotFoundError:
                print(f"  [skip] {seg} {y} 파일 없음")
    return pd.concat(dfs, ignore_index=True)

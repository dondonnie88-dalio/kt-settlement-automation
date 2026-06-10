"""유가증권성 거래 식별 진단 스크립트 v2"""
import pandas as pd

df = pd.read_parquet("processed_data.parquet", engine="pyarrow")
print(f"전체: {len(df):,}행  금액합계: {df['_amount'].sum():,.0f}원\n")

KEYWORDS = ["유가증권", "상품권", "기프트", "쿠폰", "포인트", "캐시",
            "바우처", "기프티콘", "문화", "해피", "도서", "온누리", "선불"]

# 타깃 컬럼 (컬럼 목록에서 확인된 문자열 컬럼 직접 지정)
TARGET_COLS = ["서비스카테고리", "마스터카테고리(CMS)", "상품명",
               "매출구분", "예산계정", "과세상태", "제조사"]

SEP = "=" * 70

for col in TARGET_COLS:
    if col not in df.columns:
        continue
    s = df[col].fillna("").astype(str).str.lower().str.replace(" ", "", regex=False)
    pattern = "|".join(kw.lower().replace(" ", "") for kw in KEYWORDS)
    mask = s.str.contains(pattern, regex=True, na=False)
    n = int(mask.sum())
    if n == 0:
        continue
    amt = df.loc[mask, "_amount"].sum()
    print(SEP)
    print(f"■ [{col}]  {n:,}건  /  {amt:,.0f}원  ({amt/df['_amount'].sum()*100:.1f}%)")
    print(SEP)
    top = (df.loc[mask]
             .groupby(col)["_amount"]
             .agg(건수="count", 금액="sum")
             .sort_values("금액", ascending=False)
             .head(20))
    print(top.to_string())
    if "_vendor" in df.columns:
        print("\n  [협력사 TOP10]")
        print(df.loc[mask, "_vendor"].value_counts().head(10).to_string())
    print()

# 키워드 없이 전체 고유값 확인 (건수 적은 컬럼)
print(SEP)
print("■ 매출구분 전체 고유값")
print(SEP)
print(df["매출구분"].value_counts().to_string())

print()
print(SEP)
print("■ 과세상태 전체 고유값")
print(SEP)
print(df["과세상태"].value_counts().to_string())

print()
print(SEP)
print("■ 예산계정 상위 30개")
print(SEP)
print(df["예산계정"].value_counts().head(30).to_string())

print("\n[완료]")
